"""
Collections Review Generator — Flask application.
"""
from __future__ import annotations

import datetime
import io
import logging
import os

from flask import Flask, jsonify, render_template, request, send_file

from processor.logic import process_collections
from processor.netsuite import pull_netsuite_data, check_credentials
from processor.parser import parse_invoices, parse_credits
from processor.excel_writer import build_workbook

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB upload limit


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


@app.route("/pull", methods=["POST"])
def pull():
    """Pull from NetSuite via TBA, return JSON preview."""
    data = request.get_json(force=True)
    month = int(data.get("month", datetime.date.today().month))
    year  = int(data.get("year",  datetime.date.today().year))

    try:
        invoices, credits = pull_netsuite_data()
        accounts = process_collections(invoices, credits, month, year)
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logger.exception("Unexpected error pulling from NetSuite")
        return jsonify({"error": f"Unexpected error: {e}"}), 500

    return jsonify(_build_preview(accounts, month, year))


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
