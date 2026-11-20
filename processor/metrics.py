"""
Collections metrics computation and Excel generation.
Mirrors the logic from collections_review_build_reference_oct2026.py,
adapted for live NetSuite data.
"""
from __future__ import annotations

import datetime
import io
import re
from typing import Optional

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .logic import AccountSummary

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

REPS = ["Rosas, Yoniva", "Quilantan, Maria"]

# (attempts_required, letter_required_or_None)
STATUS_GOALS: dict[str, tuple[int, Optional[str]]] = {
    "Call Customer":         (2, None),
    "Skipped Invoice":       (2, "Skipped Letter"),
    "Short Pay":             (2, None),
    "Short Paid":            (2, None),
    "Collection Letter":     (2, "Collection Letter"),
    "Predemand Letter":      (2, "Predemand Letter"),
    "Cancel NP":             (3, None),
    "Cancel Np":             (3, None),
    "Special Circumstances": (2, None),
    "Friendly Reminder Email": (0, None),
    "No Action":             (0, None),
}

# DSO thresholds — accounts with max DSO below threshold are Paid/Current
DSO_THRESHOLDS: dict[str, int] = {
    "Call Customer":     60,
    "Collection Letter": 90,
    "Predemand Letter":  120,
    "Cancel NP":         150,
    "Cancel Np":         150,
}

ACCT_PATTERN = re.compile(r"([0-9]{2}[A-Z]{2,5}[0-9]{2,5})", re.I)


def _acct_id(s: str) -> Optional[str]:
    m = ACCT_PATTERN.search(str(s))
    return m.group(1).upper() if m else None


def _to_last_first(name: str) -> str:
    """Normalize a name to 'Last, First' format.
    Leaves 'Last, First' unchanged; converts 'First Last' → 'Last, First'.
    """
    name = name.strip()
    if not name:
        return name
    if "," in name:
        return name  # already Last, First
    parts = name.split()
    if len(parts) >= 2:
        return parts[-1] + ", " + " ".join(parts[:-1])
    return name


# ---------------------------------------------------------------------------
# Core metrics computation
# ---------------------------------------------------------------------------

def compute_metrics(
    accounts: list[AccountSummary],
    notes: list[dict],
    messages: list[dict],
    letterstream_ids: set[str],   # account IDs confirmed in LetterStream export
    sf_calls: dict[str, int],     # rep_name -> call count (may be empty)
    month: int,
    year: int,
) -> dict[str, list[dict]]:
    """
    Compute per-account metrics for each rep.
    Returns {rep_name: [row_dict, ...]} for REPS only.
    """
    # Build note indices
    # notes_per_account[rep][acct_id] = count of Phone Call + Collection Note types
    # all_notes_per_account[rep][acct_id] = total note count
    phone_coll_counts: dict[str, dict[str, int]] = {r: {} for r in REPS}
    all_notes_counts:  dict[str, dict[str, int]] = {r: {} for r in REPS}
    letter_note_counts: dict[str, dict[str, int]] = {r: {} for r in REPS}

    for n in notes:
        rep   = _to_last_first(n.get("employee", "").strip())
        cust  = n.get("customer", "").strip()
        ntype = n.get("type", "").strip()
        if rep not in REPS:
            continue
        aid = _acct_id(cust) or cust
        all_notes_counts[rep][aid] = all_notes_counts[rep].get(aid, 0) + 1
        if ntype in ("Phone Call", "Collection Note"):
            phone_coll_counts[rep][aid] = phone_coll_counts[rep].get(aid, 0) + 1
        if ntype == "Letter":
            letter_note_counts[rep][aid] = letter_note_counts[rep].get(aid, 0) + 1

    # Email index: email_counts[rep][acct_id] = count
    email_counts: dict[str, dict[str, int]] = {r: {} for r in REPS}
    for m in messages:
        rep  = _to_last_first(m.get("author", "").strip())
        cust = m.get("customer", "").strip()
        if rep not in REPS:
            continue
        aid = _acct_id(cust) or cust
        email_counts[rep][aid] = email_counts[rep].get(aid, 0) + 1

    result: dict[str, list[dict]] = {}

    for rep in REPS:
        rows: list[dict] = []
        rep_accounts = [a for a in accounts if a.collections_assignment == rep]

        for acc in rep_accounts:
            cust  = acc.collect_as
            aid   = _acct_id(cust)
            status = acc.suggested_status or ""
            status_title = status.title()

            creq, lreq = STATUS_GOALS.get(status, STATUS_GOALS.get(status_title, (2, None)))

            # Attempt counts — look up by account ID
            nc    = phone_coll_counts[rep].get(aid, 0) if aid else 0
            ec    = email_counts[rep].get(aid, 0)      if aid else 0
            total = nc + ec

            # Letter evidence
            ln     = letter_note_counts[rep].get(aid, 0) if aid else 0
            lr     = (aid in letterstream_ids)            if aid else False

            # Regular open balance and DSO (excludes finance charges)
            reg_bal  = sum(
                acc.raw_balances[b] for b in ["30+", "60+", "90+", "120+", "150+", "180+"]
            )
            fc_bal   = acc.raw_balances.get("Finance Charges", 0.0) or 0.0
            total_bal = acc.total_raw_balance or 0.0

            # Max DSO (days since oldest non-FC invoice date)
            max_dso: Optional[float] = acc.oldest_dso  # may be None if not computed

            # Paid/Current determination
            paid = False
            waive_reason = "—"
            norm = status.lower()
            if norm in ("friendly reminder email", "no action"):
                paid = True
                waive_reason = "No collection action required"
            elif reg_bal <= 0.01:
                paid = True
                waive_reason = "Paid — no regular open invoices"
            elif status in DSO_THRESHOLDS and max_dso is not None and max_dso < DSO_THRESHOLDS[status]:
                paid = True
                waive_reason = f"Current — max DSO {int(max_dso)}d (below {DSO_THRESHOLDS[status]}d threshold)"

            # Compliance
            if paid:
                attempt_goal = True
                letter_goal  = True
                compliance   = "Compliant"
            else:
                attempt_goal = total >= creq
                letter_goal  = (lreq is None) or (max(ln, 1 if lr else 0) >= 1)
                compliance   = "Compliant" if (attempt_goal and letter_goal) else "Not Compliant"

            rows.append({
                "Customer":           cust,
                "Account ID":         aid or "",
                "Status":             status,
                "Regular Open Bal":   round(reg_bal, 2),
                "Finance Charge Bal": round(fc_bal, 2),
                "Total Open Bal":     round(total_bal, 2),
                "Attempts Req":       0 if paid else creq,
                "Calls (notes)":      nc,
                "Emails Sent":        ec,
                "Total Attempts":     total,
                "Attempts Left":      0 if paid else max(0, creq - total),
                "Attempt Goal Met":   "Yes" if (attempt_goal or paid) else "No",
                "Total Notes":        all_notes_counts[rep].get(aid, 0) if aid else 0,
                "Letter Req":         "—" if paid else (lreq or "—"),
                "Letter in Notes":    ln,
                "Letter in Report":   "Yes" if lr else "No",
                "Letter Goal Met":    "Yes" if (letter_goal or paid) else "No",
                "Paid / Current":     "Yes" if paid else "No",
                "Waive Reason":       waive_reason,
                "Final Compliance":   compliance,
            })

        result[rep] = rows

    return result


def _summarise(rows: list[dict]) -> tuple[int, int, int, int, int, float]:
    t = len(rows)
    p = sum(1 for r in rows if r["Paid / Current"] == "Yes")
    a = t - p
    c = sum(1 for r in rows if r["Paid / Current"] == "No" and r["Final Compliance"] == "Compliant")
    n = sum(1 for r in rows if r["Paid / Current"] == "No" and r["Final Compliance"] == "Not Compliant")
    pct = round(c / a * 100, 1) if a else 0.0
    return t, p, a, c, n, pct


# ---------------------------------------------------------------------------
# Excel generation
# ---------------------------------------------------------------------------

DARK    = "1F3864"
COL_REP = {"Quilantan, Maria": "145A32", "Rosas, Yoniva": "2E5D9B"}
RED_F   = "FCEBEB"; RED_T   = "791F1F"
GRN_F   = "EAF3DE"; GRN_T   = "27500A"
AMB_F   = "FAEEDA"; AMB_T   = "633806"
ORG_F   = "FFF3E0"; ORG_T   = "6D3B00"
BLU_F   = "E6F1FB"; BLU_T   = "0C447C"
LG      = "F4F6F9"; WH      = "FFFFFF"
BD      = "BFBFBF"

STATUS_FILL = {
    "Call Customer":           "D6EAF8",
    "Skipped Invoice":         "FEF9E7",
    "Collection Letter":       "EDE7F6",
    "Predemand Letter":        "FDEDEC",
    "Cancel NP":               "F2F3F4",
    "Cancel Np":               "F2F3F4",
    "Friendly Reminder Email": "EBF5FB",
}


def _border():
    s = Side(style="thin", color=BD)
    return Border(left=s, right=s, top=s, bottom=s)


def _wc(ws, r, c, v="", bold=False, fg="000000", bg=WH, al="center", wrap=False, sz=9, it=False):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font      = Font(name="Arial", size=sz, bold=bold, color=fg, italic=it)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=al, vertical="center", wrap_text=wrap)
    cell.border    = _border()
    return cell


def _banner(ws, rng, text, fg, bg, h=20, it=True, bold=False, sz=9):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value     = text
    c.font      = Font(name="Arial", size=sz, italic=it, bold=bold, color=fg)
    c.fill      = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[c.row].height = h


def _write_rep_sheet(wb, rep: str, rows: list[dict], month_name: str, sf_calls: int):
    ws = wb.create_sheet(rep.split(",")[0].strip() + " " + rep.split(",")[1].strip()
                         if "," in rep else rep)
    ws.sheet_view.showGridLines = False
    ch = COL_REP.get(rep, DARK)

    t, p, a, c, n, pct = _summarise(rows)

    _banner(ws, "A1:S1",
            f"{month_name} — Collections Metrics: {rep}",
            WH, ch, 28, False, True, 13)

    kpis = [
        ("Total Accounts",       t,                None),
        ("Paid / Current",       p,                GRN_F),
        ("Unpaid Active",        a,                None),
        ("Unpaid Compliant",     c,                GRN_F),
        ("Unpaid Non-Compliant", n,                RED_F),
        ("Compliance %",         f"{pct}%",        GRN_F if pct >= 80 else AMB_F),
        ("Calls (log)",          sf_calls,         None),
        ("Calls (notes)",        sum(r["Calls (notes)"] for r in rows), "E1F5EE"),
        ("Emails Sent",          sum(r["Emails Sent"]   for r in rows), "EEEDFE"),
        ("Regular Open Bal",     f"${sum(r['Regular Open Bal'] for r in rows):,.0f}", None),
        ("Finance Charge Bal",   f"${sum(r['Finance Charge Bal'] for r in rows):,.0f}", ORG_F),
    ]
    for i, (label, val, bg) in enumerate(kpis):
        col = i * 2 + 1
        for rw, fsz, v, bd in [(2, 9, label, False), (3, 14, val, True)]:
            x = ws.cell(row=rw, column=col, value=v)
            x.font      = Font(name="Arial", size=fsz, bold=bd,
                                color="666666" if rw == 2 else ch)
            x.fill      = PatternFill("solid", fgColor=bg or LG)
            x.alignment = Alignment(horizontal="center")
            ws.merge_cells(start_row=rw, start_column=col,
                           end_row=rw, end_column=col + 1)
    ws.row_dimensions[3].height = 26

    _banner(ws, "A4:S4",
            "Paid/Current = Friendly Reminder/No Action status, zero regular open invoices, "
            "or all regular invoices below the DSO threshold. Finance charges excluded. "
            "Attempts = Phone Call + Collection Note types + outgoing emails (account-linked).",
            AMB_T, AMB_F, 30)

    headers = [
        "Customer", "Acct ID", "Status", "Regular Open Bal", "Finance Charge Bal",
        "Total Open Bal\n(ref)", "Attempts Req", "Calls (notes)", "Emails",
        "Total Attempts", "Attempts Left", "Attempt Goal?", "Total Notes\n(ref)",
        "Letter Req", "Ltr Notes", "Ltr Report", "Letter Goal?",
        "Paid / Current", "Waive Reason", "Compliance",
    ]
    for ci, h in enumerate(headers, 1):
        _wc(ws, 5, ci, h, True, WH, DARK, wrap=True)
    ws.row_dimensions[5].height = 30

    r = 6
    for i, x in enumerate(rows):
        base = LG if i % 2 == 0 else WH
        cs   = x["Final Compliance"]
        paid = x["Paid / Current"] == "Yes"

        vals = [
            x["Customer"], x["Account ID"], x["Status"],
            x["Regular Open Bal"], x["Finance Charge Bal"], x["Total Open Bal"],
            x["Attempts Req"], x["Calls (notes)"], x["Emails Sent"],
            x["Total Attempts"], x["Attempts Left"], x["Attempt Goal Met"],
            x["Total Notes"], x["Letter Req"], x["Letter in Notes"],
            x["Letter in Report"], x["Letter Goal Met"],
            x["Paid / Current"], x["Waive Reason"], x["Final Compliance"],
        ]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=ci, value=v)
            cell.border    = _border()
            cell.font      = Font(name="Arial", size=9)
            cell.fill      = PatternFill("solid", fgColor=base)
            cell.alignment = Alignment(
                horizontal="left" if ci in (1, 19) else "center",
                vertical="center"
            )

            if ci == 20:  # Compliance
                cell.fill = PatternFill("solid", fgColor=GRN_F if cs == "Compliant" else RED_F)
                cell.font = Font(name="Arial", size=9, bold=True,
                                  color=GRN_T if cs == "Compliant" else RED_T)
            elif ci == 18 and paid:  # Paid/Current = Yes
                cell.fill = PatternFill("solid", fgColor=GRN_F)
                cell.font = Font(name="Arial", size=9, bold=True, color=GRN_T)
            elif ci == 5 and isinstance(v, float) and v > 0:  # Finance charge
                cell.fill = PatternFill("solid", fgColor=ORG_F)
                cell.font = Font(name="Arial", size=9, italic=True, color=ORG_T)
            elif ci == 6:  # Total Open Bal ref
                cell.fill = PatternFill("solid", fgColor=LG)
                cell.font = Font(name="Arial", size=9, italic=True, color="888888")
            elif ci in (12, 17) and v == "Yes":
                cell.fill = PatternFill("solid", fgColor=GRN_F)
                cell.font = Font(name="Arial", size=9, color=GRN_T)
            elif ci in (12, 17) and v == "No":
                cell.fill = PatternFill("solid", fgColor=RED_F)
                cell.font = Font(name="Arial", size=9, color=RED_T)
            elif ci == 11 and isinstance(v, int) and v > 0:  # Attempts left
                cell.fill = PatternFill("solid", fgColor=RED_F)
                cell.font = Font(name="Arial", size=9, bold=True, color=RED_T)

            if ci in (4, 5, 6):
                cell.number_format = '"$"#,##0.00'
                cell.alignment = Alignment(horizontal="right")
        r += 1

    col_widths = [42, 12, 18, 14, 14, 13, 10, 12, 10, 12, 12, 11, 12, 14, 10, 11, 11, 8, 30, 14]
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = ws["A6"]


def _write_summary_sheet(wb, metrics: dict[str, list[dict]], month_name: str,
                          sf_calls: dict[str, int]):
    ws = wb.create_sheet("Metrics Summary", 0)
    ws.sheet_view.showGridLines = False

    _banner(ws, "A1:L1",
            f"{month_name} — Collections Metrics Summary",
            WH, DARK, 28, False, True, 13)
    _banner(ws, "A2:L2",
            "Paid/Current = No Action/Friendly Reminder, zero balance, or all invoices below "
            "DSO threshold. Finance charges excluded. Attempts = Notes (Phone Call + "
            "Collection Note) + Emails.",
            AMB_T, AMB_F)

    headers = [
        "Associate", "Total", "Paid / Current", "Unpaid Active",
        "Unpaid Compliant", "Unpaid Non-Compliant", "Compliance %",
        "Calls (log)", "Calls (notes)", "Emails Sent",
        "Regular Open Bal", "Finance Charge Bal",
    ]
    for ci, h in enumerate(headers, 1):
        _wc(ws, 4, ci, h, True, WH, "2E5D9B", wrap=True, sz=10)
    ws.row_dimensions[4].height = 28

    name_map = {
        "Rosas, Yoniva":    "Yoniva Rosas",
        "Quilantan, Maria": "Maria Quilantan",
    }
    for ri, rep in enumerate(REPS, 5):
        rows = metrics.get(rep, [])
        t, p, a, c, n, pct = _summarise(rows)
        vals = [
            name_map.get(rep, rep), t, p, a, c, n, f"{pct}%",
            sf_calls.get(rep, 0),
            sum(r["Calls (notes)"] for r in rows),
            sum(r["Emails Sent"]   for r in rows),
            f"${sum(r['Regular Open Bal']   for r in rows):,.2f}",
            f"${sum(r['Finance Charge Bal'] for r in rows):,.2f}",
        ]
        for ci, v in enumerate(vals, 1):
            x = _wc(ws, ri, ci, v, sz=10, bg=LG if ri % 2 == 0 else WH,
                    al="left" if ci == 1 else "center")
            if ci == 3:
                x.fill = PatternFill("solid", fgColor=GRN_F)
                x.font = Font(name="Arial", size=10, bold=True, color=GRN_T)
            if ci == 12:
                x.fill = PatternFill("solid", fgColor=ORG_F)
                x.font = Font(name="Arial", size=10, italic=True, color=ORG_T)

    # Status breakdown
    r = 8
    ws.cell(row=r, column=1, value="Status Breakdown — Combined").font = \
        Font(name="Arial", size=10, bold=True, color=DARK)
    r += 1

    h2 = ["Status", "Total", "Paid/Current", "Unpaid Active",
          "Unpaid Compliant", "Unpaid Non-Compliant", "Compliance %",
          "Calls (notes)", "Emails", "Regular Open Bal", "Finance Charge Bal"]
    for ci, h in enumerate(h2, 1):
        _wc(ws, r, ci, h, True, WH, DARK)
    r += 1

    all_rows = [row for rows in metrics.values() for row in rows]
    for status in [
        "Call Customer", "Collection Letter", "Skipped Invoice",
        "Predemand Letter", "Cancel NP", "Cancel Np", "Friendly Reminder Email",
    ]:
        subset = [row for row in all_rows if row["Status"] in (status, status.title())]
        if not subset:
            continue
        t, p, a, c, n, pct = _summarise(subset)
        fill_bg = STATUS_FILL.get(status, WH)
        for ci, v in enumerate([
            status, t, p, a, c, n, f"{pct}%",
            sum(x["Calls (notes)"] for x in subset),
            sum(x["Emails Sent"]   for x in subset),
            f"${sum(x['Regular Open Bal']   for x in subset):,.2f}",
            f"${sum(x['Finance Charge Bal'] for x in subset):,.2f}",
        ], 1):
            x = _wc(ws, r, ci, v, bg=fill_bg, al="left" if ci == 1 else "center")
            if ci == 3:
                x.fill = PatternFill("solid", fgColor=GRN_F)
            if ci == 11:
                x.fill = PatternFill("solid", fgColor=ORG_F)
        r += 1

    col_widths = [24, 9, 13, 13, 16, 18, 12, 10, 12, 10, 16, 16]
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w


def build_metrics_workbook(
    accounts: list[AccountSummary],
    notes: list[dict],
    messages: list[dict],
    letterstream_ids: set[str],
    sf_calls: dict[str, int],
    month: int,
    year: int,
) -> bytes:
    """Build the metrics Excel workbook and return as bytes."""
    month_name = datetime.date(year, month, 1).strftime("%B %Y")

    metrics = compute_metrics(
        accounts, notes, messages, letterstream_ids, sf_calls, month, year
    )

    wb = Workbook()
    wb.remove(wb.active)

    _write_summary_sheet(wb, metrics, month_name, sf_calls)
    for rep in REPS:
        _write_rep_sheet(wb, rep, metrics.get(rep, []), month_name,
                         sf_calls.get(rep, 0))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
