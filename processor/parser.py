"""
Parse NetSuite export files (SpreadsheetML XLS or CSV) into lists of dicts.
Column matching is case-insensitive.
"""
import csv
import io
import xml.etree.ElementTree as ET
import logging
import re

logger = logging.getLogger(__name__)

# SpreadsheetML namespace
SS_NS = "urn:schemas-microsoft-com:office:spreadsheet"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def parse_invoices(content: str | bytes) -> list[dict]:
    """Parse invoice export. Returns list of row dicts."""
    rows = _parse_content(content)
    return _normalise_invoices(rows)


def parse_credits(content: str | bytes) -> list[dict]:
    """Parse credit memo export. Returns list of row dicts."""
    rows = _parse_content(content)
    return _normalise_credits(rows)


# ---------------------------------------------------------------------------
# Detect format and parse
# ---------------------------------------------------------------------------

def _parse_content(content: str | bytes) -> list[dict]:
    if isinstance(content, bytes):
        content = content.decode("utf-8", errors="replace")

    stripped = content.lstrip()
    if stripped.startswith("<?xml") or stripped.startswith("<Workbook") or stripped.startswith("<ss:"):
        return _parse_spreadsheetml(content)
    else:
        return _parse_csv(content)


def _parse_csv(text: str) -> list[dict]:
    reader = csv.DictReader(io.StringIO(text))
    rows = list(reader)
    # Normalise header keys: strip whitespace
    normalised = []
    for row in rows:
        normalised.append({k.strip(): v.strip() if isinstance(v, str) else v for k, v in row.items()})
    return normalised


def _parse_spreadsheetml(xml_text: str) -> list[dict]:
    """Parse SpreadsheetML (Office 2003 XML) into list of dicts."""
    try:
        root = ET.fromstring(xml_text.encode("utf-8"))
    except ET.ParseError as e:
        raise ValueError(f"Failed to parse SpreadsheetML XML: {e}")

    ns = {"ss": SS_NS}

    # Find the first worksheet
    worksheet = root.find(".//ss:Worksheet", ns)
    if worksheet is None:
        # Try without namespace prefix
        worksheet = root.find(".//{%s}Worksheet" % SS_NS)
    if worksheet is None:
        raise ValueError("No worksheet found in SpreadsheetML file.")

    table = worksheet.find("ss:Table", ns) or worksheet.find("{%s}Table" % SS_NS)
    if table is None:
        raise ValueError("No table found in worksheet.")

    rows_el = table.findall("ss:Row", ns) or table.findall("{%s}Row" % SS_NS)

    result = []
    headers = []

    for i, row_el in enumerate(rows_el):
        cells_el = row_el.findall("ss:Cell", ns) or row_el.findall("{%s}Cell" % SS_NS)
        cell_values = []
        col_index = 0
        for cell_el in cells_el:
            # Handle ss:Index attribute for sparse rows
            idx_attr = cell_el.get("{%s}Index" % SS_NS) or cell_el.get("ss:Index")
            if idx_attr:
                col_index = int(idx_attr) - 1
            data_el = cell_el.find("ss:Data", ns) or cell_el.find("{%s}Data" % SS_NS)
            value = data_el.text if data_el is not None else ""
            value = value.strip() if value else ""

            # Pad with empty strings for any skipped columns
            while len(cell_values) < col_index:
                cell_values.append("")
            cell_values.append(value)
            col_index += 1

        if i == 0:
            headers = cell_values
        else:
            if not any(cell_values):
                continue  # skip empty rows
            row_dict = {}
            for j, h in enumerate(headers):
                row_dict[h.strip()] = cell_values[j] if j < len(cell_values) else ""
            result.append(row_dict)

    return result


# ---------------------------------------------------------------------------
# Column normalisation helpers
# ---------------------------------------------------------------------------

def _find_col(row: dict, *candidates: str) -> str | None:
    """Case-insensitive search for column value."""
    row_lower = {k.lower(): v for k, v in row.items()}
    for c in candidates:
        if c.lower() in row_lower:
            return row_lower[c.lower()]
    return None


def _col_key(headers: list[str], *candidates: str) -> str | None:
    """Return the actual header key matching one of the candidates."""
    for h in headers:
        for c in candidates:
            if h.strip().lower() == c.lower():
                return h
    return None


def _safe_float(val: str | None) -> float:
    if not val:
        return 0.0
    val = str(val).replace(",", "").replace("$", "").strip()
    try:
        return float(val)
    except ValueError:
        return 0.0


def _normalise_invoices(rows: list[dict]) -> list[dict]:
    result = []
    for row in rows:
        def g(*names):
            return _find_col(row, *names) or ""

        record = {
            "collect_as": g("Collect As", "CollectAs"),
            "date": g("Date", "Invoice Date"),
            "amount_remaining": _safe_float(g("Amount Remaining", "Amount (Remaining)")),
            "business_unit": g("Business Unit", "BusinessUnit"),
            "category": g("Category"),
            "collections_status": g("Collections Status", "Collection Status"),
            "is_finance_charge": g("Is Finance Charge", "Finance Charge"),
            "collection_escalation_status": g(
                "Collection Escalation Status", "Escalation Status"
            ),
            "fortis_autopay_enrollment": g(
                "Fortis Autopay Enrollment", "Autopay Enrollment", "Autopay"
            ),
            "account_restricted": g("Account Restricted", "Restricted"),
        }
        if record["collect_as"]:
            result.append(record)
    return result


def _normalise_credits(rows: list[dict]) -> list[dict]:
    result = []
    for row in rows:
        def g(*names):
            return _find_col(row, *names) or ""

        custom_form = g("Custom Form", "Form")
        # Exclude SRDP credits
        if "srdp" in custom_form.lower():
            continue

        record = {
            "collect_as": g("Collect As", "CollectAs"),
            "amount_remaining": _safe_float(g("Amount Remaining", "Amount (Remaining)")),
            "custom_form": custom_form,
        }
        if record["collect_as"]:
            result.append(record)
    return result
