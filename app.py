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

import csv

from processor.logic import process_collections
from processor.netsuite import pull_netsuite_data, pull_metrics_data, check_credentials
from processor.parser import parse_invoices, parse_credits
from processor.excel_writer import build_workbook
from processor.metrics import build_metrics_workbook

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

        def progress(msg: str):
            _set_job(job_id, message=msg)

        invoices, credits = pull_netsuite_data(progress_cb=progress)
        logger.info("Job %s: pull returned %d invoices, %d credits",
                    job_id, len(invoices), len(credits))
        if not credits:
            logger.warning("Job %s: 0 credits returned — credit totals will be blank", job_id)

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

        def progress(msg: str):
            _set_job(job_id, message=msg)

        invoices, credits = pull_netsuite_data(progress_cb=progress)
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


# ---------------------------------------------------------------------------
# Metrics endpoints
# ---------------------------------------------------------------------------

def _parse_review_xlsx(raw_bytes: bytes) -> list | None:
    """
    Parse a Collections Review XLSX and return a list of AccountSummary objects
    with collect_as, suggested_status, collections_assignment, and raw_balances
    populated from the file.  Returns None if parsing fails.

    Column names expected (from excel_writer.OUTPUT_COLUMNS):
      Collect As, Suggested Status, Collections Assignment,
      Current Balance, 30+ Balance, 60+ Balance, 90+ Balance,
      120+ Balance, 150+ Balance, 180+ Balance, Finance Charges
    """
    from processor.logic import AccountSummary
    import io as _io

    BALANCE_COL_MAP = {
        "Current Balance":  "Current",
        "30+ Balance":      "30+",
        "60+ Balance":      "60+",
        "90+ Balance":      "90+",
        "120+ Balance":     "120+",
        "150+ Balance":     "150+",
        "180+ Balance":     "180+",
        "Finance Charges":  "Finance Charges",
    }

    try:
        import openpyxl as _xl
        wb = _xl.load_workbook(_io.BytesIO(raw_bytes), read_only=True, data_only=True)
        # First sheet = the main review sheet
        ws = wb.worksheets[0]
        rows = list(ws.rows)
        if not rows:
            return None

        # Find header row (first row whose cells contain "Collect As")
        header_idx = None
        headers: list[str] = []
        for i, row in enumerate(rows):
            vals = [str(c.value or "").strip() for c in row]
            if "Collect As" in vals:
                header_idx = i
                headers = vals
                break
        if header_idx is None:
            logger.warning("review XLSX: could not find 'Collect As' header row")
            return None

        def col_idx(name: str) -> int | None:
            try:
                return headers.index(name)
            except ValueError:
                # Try partial / case-insensitive match for "Suggested Status" which
                # the writer suffixes with the month name
                name_lower = name.lower()
                for j, h in enumerate(headers):
                    if h.lower().startswith(name_lower):
                        return j
            return None

        ci_ca    = col_idx("Collect As")
        ci_st    = col_idx("Suggested Status")
        ci_as    = col_idx("Collections Assignment")
        if ci_ca is None:
            return None

        balance_cols: dict[str, int] = {}
        for col_name, bucket in BALANCE_COL_MAP.items():
            idx = col_idx(col_name)
            if idx is not None:
                balance_cols[bucket] = idx

        accounts: list[AccountSummary] = []
        for row in rows[header_idx + 1:]:
            cells = [c.value for c in row]
            if not cells or ci_ca >= len(cells):
                continue
            ca = str(cells[ci_ca] or "").strip()
            if not ca or ca.startswith("—") or ca.lower() in ("collect as", ""):
                continue

            status     = str(cells[ci_st]  or "").strip() if ci_st  is not None and ci_st  < len(cells) else ""
            assignment = str(cells[ci_as]  or "").strip() if ci_as  is not None and ci_as  < len(cells) else ""

            acc = AccountSummary(collect_as=ca)
            acc.suggested_status      = status
            acc.collections_assignment = assignment
            # Populate raw_balances from the file
            for bucket, idx in balance_cols.items():
                if idx < len(cells):
                    try:
                        acc.raw_balances[bucket] = float(cells[idx] or 0)
                    except (TypeError, ValueError):
                        pass
            # Mirror raw_balances → balances so total_raw_balance works
            for bucket in acc.raw_balances:
                acc.balances[bucket] = acc.raw_balances[bucket]

            accounts.append(acc)

        logger.info("review XLSX: parsed %d account rows", len(accounts))
        return accounts if accounts else None

    except Exception as e:
        logger.warning("review XLSX parse error: %s", e)
        return None


def _parse_letterstream_csv(raw_bytes: bytes) -> set[str]:
    """
    Parse a LetterStream activity export CSV and return the set of account IDs
    found in the Job Name column.  Line 1 is a title row — skip it.
    """
    import re as _re
    pat = _re.compile(r"([0-9]{2}[A-Z]{2,5}[0-9]{2,5})", _re.I)
    ids: set[str] = set()
    try:
        text = raw_bytes.decode("utf-8-sig", errors="replace")
        lines = text.splitlines()
        if len(lines) > 1:
            reader = csv.DictReader(lines[1:])
            for row in reader:
                job = row.get("Job Name") or row.get(list(row.keys())[0] if row else "") or ""
                m = pat.search(str(job))
                if m:
                    ids.add(m.group(1).upper())
    except Exception as e:
        logger.warning("LetterStream parse error: %s", e)
    return ids


def _run_metrics_job(job_id: str, month: int, year: int,
                     letterstream_ids: set[str],
                     review_accounts: list | None = None) -> None:
    """Background thread: pull NS data + metrics data, compute metrics, store result."""
    try:
        def progress(msg: str):
            _set_job(job_id, message=msg)

        if review_accounts is not None:
            # Use pre-parsed assignments from the uploaded review file
            accounts = review_accounts
            logger.info("Metrics job %s: using %d accounts from uploaded review XLSX",
                        job_id, len(accounts))
        else:
            # Fall back: re-derive assignments from live NetSuite data
            _set_job(job_id, status="running", message="Pulling invoices from NetSuite…")
            invoices, credits = pull_netsuite_data(progress_cb=progress)
            _set_job(job_id, message=f"Processing {len(invoices)} invoices…")
            accounts = process_collections(invoices, credits, month, year)
            logger.info("Metrics job %s: derived %d accounts from live pull (no review file uploaded)",
                        job_id, len(accounts))

            # Diagnostic logs (live pull path only)
            fc_nonzero = [(a.collect_as, a.raw_balances.get("Finance Charges", 0.0))
                          for a in accounts if a.raw_balances.get("Finance Charges", 0.0) > 0.01]
            logger.info("Metrics job %s: %d accounts have non-zero FC; sample: %s",
                        job_id, len(fc_nonzero), fc_nonzero[:5])
            dso_sample = [(a.collect_as[:30], round(a.oldest_dso or 0))
                          for a in accounts if a.oldest_dso is not None][:5]
            logger.info("Metrics job %s: DSO sample (from due date): %s", job_id, dso_sample)

        _set_job(job_id, status="running", message="Pulling collections notes and emails from NetSuite…")
        notes, messages = pull_metrics_data(month, year, progress_cb=progress)

        _set_job(job_id, message="Building metrics report…")
        xlsx_bytes = build_metrics_workbook(
            accounts=accounts,
            notes=notes,
            messages=messages,
            letterstream_ids=letterstream_ids,
            sf_calls={},   # Salesforce integration placeholder
            month=month,
            year=year,
        )
        _set_job(job_id, status="done", xlsx=xlsx_bytes, month=month, year=year)
        logger.info("Metrics job %s complete (%d bytes)", job_id, len(xlsx_bytes))

    except RuntimeError as e:
        logger.error("Metrics job %s RuntimeError: %s", job_id, e)
        _set_job(job_id, status="error", error=str(e))
    except Exception as e:
        logger.exception("Metrics job %s unexpected error", job_id)
        _set_job(job_id, status="error", error=f"{type(e).__name__}: {e}")


@app.route("/metrics/start", methods=["POST"])
def metrics_start():
    """
    Start a background metrics job.
    Accepts multipart/form-data with optional fields:
      month (int), year (int), letterstream_file (CSV upload)
    Returns {job_id}.
    """
    month = int(request.form.get("month", datetime.date.today().month))
    year  = int(request.form.get("year",  datetime.date.today().year))

    review_accounts = None
    rv_file = request.files.get("review_file")
    if rv_file:
        review_accounts = _parse_review_xlsx(rv_file.read())
        if review_accounts is None:
            logger.warning("review XLSX upload could not be parsed — will re-derive from live data")
        else:
            logger.info("review XLSX: %d accounts loaded", len(review_accounts))

    letterstream_ids: set[str] = set()
    ls_file = request.files.get("letterstream_file")
    if ls_file:
        letterstream_ids = _parse_letterstream_csv(ls_file.read())
        logger.info("LetterStream: %d account IDs parsed", len(letterstream_ids))

    job_id = uuid.uuid4().hex
    _set_job(job_id, status="pending", month=month, year=year)
    t = threading.Thread(
        target=_run_metrics_job,
        args=(job_id, month, year, letterstream_ids, review_accounts),
        daemon=True,
    )
    t.start()
    logger.info("Started metrics job %s for %d/%d", job_id, month, year)
    return jsonify({"job_id": job_id})


@app.route("/metrics/status/<job_id>")
def metrics_status(job_id: str):
    """Poll for metrics job status."""
    job = _get_job(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    safe = {k: v for k, v in job.items() if k != "xlsx"}
    return jsonify(safe)


@app.route("/metrics/download/<job_id>")
def metrics_download(job_id: str):
    """Download the metrics XLSX once the job is done."""
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
    m, y = job["month"], job["year"]
    month_name = datetime.date(y, m, 1).strftime("%B_%Y")
    filename = f"Collections_Metrics_{month_name}.xlsx"
    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
