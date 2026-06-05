"""
Collections Review Generator — Flask application.
"""
from __future__ import annotations

import datetime
import io
import logging
import os
import threading
import uuid

from flask import Flask, jsonify, render_template, request, send_file

from processor.logic import process_collections
from processor.netsuite import pull_netsuite_data, check_credentials
from processor.parser import parse_invoices, parse_credits
from processor.excel_writer import build_workbook

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB upload limit

# ---------------------------------------------------------------------------
# In-memory job store for async NS pulls
# (single-process; fine for Railway's single-replica deploy)
# ---------------------------------------------------------------------------
_jobs: dict[str, dict] = {}
_jobs_lock = threading.Lock()


def _set_job(job_id: str, **kwargs) -> None:
    with _jobs_lock:
        if job_id not in _jobs:
            _jobs[job_id] = {}
        _jobs[job_id].update(kwargs)


def _get_job(job_id: str) -> dict | None:
    with _jobs_lock:
        return dict(_jobs.get(job_id, {}))


# ---------------------------------------------------------------------------
# Error handler
# ---------------------------------------------------------------------------

@app.errorhandler(Exception)
def handle_exception(e):
    """Return JSON for all unhandled exceptions so the frontend never gets HTML."""
    logger.exception("Unhandled exception: %s", e)
    return jsonify({"error": f"{type(e).__name__}: {e}"}), 500


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html", ns_configured=check_credentials())


@app.route("/health")
def health():
    return jsonify({"status": "ok", "ns_configured": check_credentials()})


@app.route("/diagnose")
def diagnose():
    """Run diagnostic SuiteQL queries — check Railway logs for results."""
    from processor.netsuite import diagnose as _diag
    try:
        results = _diag()
        return jsonify({"status": "ok", "results": results})
    except Exception as e:
        logger.exception("Diagnose failed")
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# Async pull endpoints
# ---------------------------------------------------------------------------

def _run_pull_job(job_id: str, month: int, year: int) -> None:
    """Background thread: pull NS data, process, store result in _jobs."""
    try:
        _set_job(job_id, status="running", message="Connecting to NetSuite…")
        invoices, credits = pull_netsuite_data()
        logger.info("Job %s: pull_netsuite_data returned %d invoices, %d credits",
                    job_id, len(invoices), len(credits))

        _set_job(job_id, message=f"Processing {len(invoices)} invoices…")
        accounts = process_collections(invoices, credits, month, year)
        logger.info("Job %s: process_collections returned %d accounts", job_id, len(accounts))

        preview = _build_preview(accounts, month, year)
        logger.info("Job %s: preview built: %s", job_id, preview)
        # Store atomically — status+preview in one update so poll never sees done without preview
        with _jobs_lock:
            _jobs.setdefault(job_id, {}).update({"status": "done", "preview": preview})

    except RuntimeError as e:
        logger.error("Job %s RuntimeError: %s", job_id, e)
        _set_job(job_id, status="error", error=str(e))
    except Exception as e:
        logger.exception("Job %s unexpected error", job_id)
        _set_job(job_id, status="error", error=f"{type(e).__name__}: {e}")


@app.route("/pull/start", methods=["POST"])
def pull_start():
    """Start a background NS pull. Returns {job_id} immediately."""
    data = request.get_json(force=True)
    month = int(data.get("month", datetime.date.today().month))
    year  = int(data.get("year",  datetime.date.today().year))

    job_id = uuid.uuid4().hex
    _set_job(job_id, status="pending", month=month, year=year)

    t = threading.Thread(target=_run_pull_job, args=(job_id, month, year), daemon=True)
    t.start()
    logger.info("Started pull job %s for %d/%d", job_id, month, year)
    return jsonify({"job_id": job_id})


@app.route("/pull/status/<job_id>")
def pull_status(job_id: str):
    """Poll for job status. Returns {status, preview?, error?, message?}."""
    job = _get_job(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify(job)


# Keep old /pull endpoint for backwards compatibility (still async internally)
@app.route("/pull", methods=["POST"])
def pull():
    """Legacy pull endpoint — now delegates to async pull/start + immediate poll."""
    return pull_start()


# ---------------------------------------------------------------------------
# Generate endpoints
# ---------------------------------------------------------------------------

def _run_generate_job(job_id: str, month: int, year: int) -> None:
    """Background thread: pull + generate XLSX, store bytes in _jobs."""
    try:
        _set_job(job_id, status="running", message="Pulling from NetSuite…")
        invoices, credits = pull_netsuite_data()
        _set_job(job_id, message=f"Generating report for {len(invoices)} invoices…")
        accounts   = process_collections(invoices, credits, month, year)
        xlsx_bytes = build_workbook(accounts, month, year)
        _set_job(job_id, status="done", xlsx=xlsx_bytes, month=month, year=year)
        logger.info("Job %s: generate complete (%d bytes)", job_id, len(xlsx_bytes))
    except RuntimeError as e:
        logger.error("Job %s RuntimeError: %s", job_id, e)
        _set_job(job_id, status="error", error=str(e))
    except Exception as e:
        logger.exception("Job %s unexpected error", job_id)
        _set_job(job_id, status="error", error=f"{type(e).__name__}: {e}")


@app.route("/generate/start", methods=["POST"])
def generate_start():
    """Start async generate from NS. Returns {job_id}."""
    data  = request.get_json(force=True)
    month = int(data.get("month", datetime.date.today().month))
    year  = int(data.get("year",  datetime.date.today().year))

    job_id = uuid.uuid4().hex
    _set_job(job_id, status="pending")
    t = threading.Thread(target=_run_generate_job, args=(job_id, month, year), daemon=True)
    t.start()
    return jsonify({"job_id": job_id})


@app.route("/generate/status/<job_id>")
def generate_status(job_id: str):
    """Poll for generate job status."""
    job = _get_job(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    # Don't return xlsx bytes in status poll — just status/error/message
    safe = {k: v for k, v in job.items() if k != "xlsx"}
    return jsonify(safe)


@app.route("/generate/download/<job_id>")
def generate_download(job_id: str):
    """Download the XLSX once the generate job is done."""
    job = _get_job(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    if job.get("status") == "error":
        return jsonify({"error": job.get("error", "Unknown error")}), 500
    if job.get("status") != "done":
        return jsonify({"error": "Not ready yet"}), 202
    xlsx_bytes = job.get("xlsx")
    if not xlsx_bytes:
        return jsonify({"error": "No file data"}), 500
    return _xlsx_response(xlsx_bytes, job["month"], job["year"])


@app.route("/generate", methods=["POST"])
def generate():
    """
    Generate and return the XLSX file.
    Accepts either:
      (a) multipart form with invoice_file + credits_file + month + year
      (b) JSON with {source: "netsuite", month, year}
    """
    if request.content_type and "multipart/form-data" in request.content_type:
        return _generate_from_upload()
    return _generate_from_netsuite()


def _generate_from_upload():
    invoice_file = request.files.get("invoice_file")
    credits_file = request.files.get("credits_file")
    month = int(request.form.get("month", datetime.date.today().month))
    year  = int(request.form.get("year",  datetime.date.today().year))

    if not invoice_file or not credits_file:
        return jsonify({"error": "Both invoice_file and credits_file are required."}), 400

    try:
        invoices = parse_invoices(invoice_file.read())
        credits  = parse_credits(credits_file.read())
        accounts = process_collections(invoices, credits, month, year)
        xlsx_bytes = build_workbook(accounts, month, year)
    except Exception as e:
        logger.exception("Error generating report from upload")
        return jsonify({"error": str(e)}), 500

    return _xlsx_response(xlsx_bytes, month, year)


def _generate_from_netsuite():
    """Synchronous NS generate — kept for upload path compatibility."""
    data  = request.get_json(force=True)
    month = int(data.get("month", datetime.date.today().month))
    year  = int(data.get("year",  datetime.date.today().year))

    try:
        invoices, credits = pull_netsuite_data()
        accounts   = process_collections(invoices, credits, month, year)
        xlsx_bytes = build_workbook(accounts, month, year)
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logger.exception("Error generating report from NetSuite")
        return jsonify({"error": str(e)}), 500

    return _xlsx_response(xlsx_bytes, month, year)


def _xlsx_response(xlsx_bytes: bytes, month: int, year: int):
    month_name = datetime.date(year, month, 1).strftime("%B_%Y")
    filename = f"Collections_Review_{month_name}.xlsx"
    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


def _build_preview(accounts, month, year):
    from collections import Counter
    status_counts = Counter(
        a.suggested_status for a in accounts if a.section == "actionable"
    )
    return {
        "status": "ok",
        "month": month,
        "year": year,
        "total_accounts": len(accounts),
        "section_counts": {
            "actionable":    sum(1 for a in accounts if a.section == "actionable"),
            "current_only":  sum(1 for a in accounts if a.section == "current_only"),
            "autopay":       sum(1 for a in accounts if a.section == "autopay"),
        },
        "status_counts": dict(status_counts),
        "errors": [],
    }


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
