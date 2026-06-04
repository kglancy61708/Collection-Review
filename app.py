"""
Collections Review Generator — Flask application.
"""
from __future__ import annotations

import datetime
import io
import logging
import os
import traceback

from flask import Flask, jsonify, render_template, request, send_file

from processor.logic import process_collections, STATUS_SEVERITY
from processor.netsuite import pull_netsuite_data
from processor.parser import parse_invoices, parse_credits
from processor.excel_writer import build_workbook

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB upload limit


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/pull", methods=["POST"])
def pull():
    """
    Pull data from NetSuite, run processing, return JSON preview.
    Body: {account_id, email, password, month, year}
    """
    data = request.get_json(force=True)
    account_id = data.get("account_id", "").strip()
    email = data.get("email", "").strip()
    password = data.get("password", "")
    month = int(data.get("month", datetime.date.today().month))
    year = int(data.get("year", datetime.date.today().year))

    if not account_id or not email or not password:
        return jsonify({"error": "account_id, email, and password are required."}), 400

    try:
        invoice_csv, credit_csv = pull_netsuite_data(account_id, email, password)
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logger.exception("Unexpected error pulling from NetSuite")
        return jsonify({"error": f"Unexpected error: {e}"}), 500

    try:
        invoices = parse_invoices(invoice_csv)
        credits = parse_credits(credit_csv)
        accounts = process_collections(invoices, credits, month, year)
    except Exception as e:
        logger.exception("Error processing NetSuite data")
        return jsonify({"error": f"Processing error: {e}"}), 500

    preview = _build_preview(accounts, month, year)
    return jsonify(preview)


@app.route("/generate", methods=["POST"])
def generate():
    """
    Generate and return the XLSX file.
    Accepts either:
      (a) multipart form with invoice_file + credits_file + month + year
      (b) JSON with account_id + email + password + month + year
    """
    # Determine input mode
    if request.content_type and "multipart/form-data" in request.content_type:
        return _generate_from_upload()
    else:
        return _generate_from_netsuite()


def _generate_from_upload():
    invoice_file = request.files.get("invoice_file")
    credits_file = request.files.get("credits_file")
    month = int(request.form.get("month", datetime.date.today().month))
    year = int(request.form.get("year", datetime.date.today().year))

    if not invoice_file or not credits_file:
        return jsonify({"error": "Both invoice_file and credits_file are required."}), 400

    try:
        invoice_bytes = invoice_file.read()
        credit_bytes = credits_file.read()
        invoices = parse_invoices(invoice_bytes)
        credits = parse_credits(credit_bytes)
        accounts = process_collections(invoices, credits, month, year)
        xlsx_bytes = build_workbook(accounts, month, year)
    except Exception as e:
        logger.exception("Error generating report from upload")
        return jsonify({"error": str(e)}), 500

    return _xlsx_response(xlsx_bytes, month, year)


def _generate_from_netsuite():
    data = request.get_json(force=True)
    account_id = data.get("account_id", "").strip()
    email = data.get("email", "").strip()
    password = data.get("password", "")
    month = int(data.get("month", datetime.date.today().month))
    year = int(data.get("year", datetime.date.today().year))

    if not account_id or not email or not password:
        return jsonify({"error": "account_id, email, and password are required."}), 400

    try:
        invoice_csv, credit_csv = pull_netsuite_data(account_id, email, password)
        invoices = parse_invoices(invoice_csv)
        credits = parse_credits(credit_csv)
        accounts = process_collections(invoices, credits, month, year)
        xlsx_bytes = build_workbook(accounts, month, year)
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logger.exception("Error generating report from NetSuite")
        return jsonify({"error": f"Error: {e}"}), 500

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
    status_counts = Counter()
    for a in accounts:
        if a.section == "actionable":
            status_counts[a.suggested_status] += 1

    section_counts = {
        "actionable": sum(1 for a in accounts if a.section == "actionable"),
        "current_only": sum(1 for a in accounts if a.section == "current_only"),
        "autopay": sum(1 for a in accounts if a.section == "autopay"),
    }

    return {
        "status": "ok",
        "month": month,
        "year": year,
        "total_accounts": len(accounts),
        "section_counts": section_counts,
        "status_counts": dict(status_counts),
        "errors": [],
    }


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
