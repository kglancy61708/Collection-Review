"""
Core business logic for the collections review report.

Given parsed invoice and credit rows plus a reporting month/year,
produces a list of AccountSummary objects ready for Excel output.
"""
from __future__ import annotations

import datetime
import re
from dataclasses import dataclass, field
from typing import Optional

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

BUCKETS = ["Current", "30+", "60+", "90+", "120+", "150+", "180+", "Finance Charges"]
AGED_BUCKETS = ["30+", "60+", "90+", "120+", "150+", "180+"]

STATUS_SEVERITY = {
    "Cancel NP": 1,
    "Predemand Letter": 2,
    "Collection Letter": 3,
    "Call Customer": 4,
    "Friendly Reminder Email": 5,
    "Skipped Invoice": 6,
    "Special Circumstances": 7,
    "No Action": 8,
}

# Special account group patterns (lowercase match)
SALES_FULL = ["decron properties", "bridge property management", "richdale group"]
SALES_PARTIAL = ["dominium management services"]
GREYSTAR = "greystar management"
WITHHELD_CODES = ["10uni100", "10uni101", "10leg300"]


# ---------------------------------------------------------------------------
# Date / bucket helpers
# ---------------------------------------------------------------------------

def _first_business_day(year: int, month: int) -> datetime.date:
    d = datetime.date(year, month, 1)
    while d.weekday() >= 5:  # 5=Sat, 6=Sun
        d += datetime.timedelta(days=1)
    return d


def _last_day(year: int, month: int) -> datetime.date:
    if month == 12:
        return datetime.date(year, 12, 31)
    return datetime.date(year, month + 1, 1) - datetime.timedelta(days=1)


def _month_offset(year: int, month: int, offset: int) -> tuple[int, int]:
    """Return (year, month) shifted by `offset` months."""
    total = month - 1 + offset
    y = year + total // 12
    m = total % 12 + 1
    return y, m


def compute_bucket_ranges(report_year: int, report_month: int) -> dict:
    """
    Returns dict of bucket_name -> (start_date, end_date) inclusive.
    Future -> (end_of_current_month+1, max_date)  — excluded from totals.
    """
    # Current: 2nd of prior month through last day of report month
    py, pm = _month_offset(report_year, report_month, -1)
    current_start = datetime.date(py, pm, 2)
    current_end = _last_day(report_year, report_month)

    ranges = {}
    ranges["Current"] = (current_start, current_end)

    # 30+, 60+, ..., 150+: each is (2nd of N months ago) to (1st of N-1 months ago)
    for i, label in enumerate(["30+", "60+", "90+", "120+", "150+"], start=1):
        start_y, start_m = _month_offset(report_year, report_month, -(i + 1))
        end_y, end_m = _month_offset(report_year, report_month, -i)
        ranges[label] = (datetime.date(start_y, start_m, 2), datetime.date(end_y, end_m, 1))

    # 180+: everything on or before the day before 150+ start
    start_150 = ranges["150+"][0]
    ranges["180+"] = (datetime.date(1900, 1, 1), start_150 - datetime.timedelta(days=1))

    # Future: after end of current month
    ranges["Future"] = (current_end + datetime.timedelta(days=1), datetime.date(9999, 12, 31))

    return ranges


def date_to_bucket(invoice_date: datetime.date, bucket_ranges: dict) -> str:
    """Assign an invoice date to a bucket. Returns bucket name or 'Future'."""
    for label in ["Future", "Current", "30+", "60+", "90+", "120+", "150+", "180+"]:
        start, end = bucket_ranges[label]
        if start <= invoice_date <= end:
            return label
    return "180+"  # fallback for very old dates


def _parse_date(date_str: str) -> Optional[datetime.date]:
    if not date_str:
        return None
    s = date_str.strip()
    # ISO datetime with time component (SpreadsheetML): "2024-05-15T00:00:00.000"
    if "T" in s:
        s = s.split("T")[0]
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%m-%d-%Y", "%m/%d/%y"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------------------
# Account data structure
# ---------------------------------------------------------------------------

@dataclass
class AccountSummary:
    collect_as: str
    business_unit: str = ""
    category: str = ""
    current_collections_status: str = ""
    collection_escalation_status: str = ""
    account_restricted: str = ""
    future_restriction: str = ""
    fortis_autopay_enrollment: str = ""

    # Bucket balances (post-credit-waterfall)
    balances: dict = field(default_factory=lambda: {b: 0.0 for b in BUCKETS})

    # Pre-waterfall balances (for determining original status)
    raw_balances: dict = field(default_factory=lambda: {b: 0.0 for b in BUCKETS})

    total_open_credits: float = 0.0
    credit_adjustment: str = ""  # e.g. "180+, 90+"

    invoice_count: int = 0
    suggested_status: str = ""
    collections_assignment: str = ""

    is_autopay: bool = False
    section: str = "actionable"  # "actionable", "current_only", "autopay"

    @property
    def total_open_balance(self) -> float:
        return sum(self.balances[b] for b in BUCKETS if b != "Future")

    @property
    def total_aged_balance(self) -> float:
        return sum(self.balances[b] for b in AGED_BUCKETS)

    @property
    def oldest_bucket_with_balance(self) -> str:
        for b in reversed(AGED_BUCKETS):
            if self.balances[b] > 0.01:
                return b
        if self.balances["Current"] > 0.01:
            return "Current"
        return ""


# ---------------------------------------------------------------------------
# Exclusion helpers
# ---------------------------------------------------------------------------

def _should_exclude_invoice(inv: dict) -> bool:
    status = inv.get("collections_status", "").strip().lower()
    collect_as = inv.get("collect_as", "").lower()
    if status == "legal":
        return True
    if "collection agency" in status:
        return True
    if "goldspur" in collect_as:
        return True
    if "aquasol companies employee" in collect_as:
        return True
    return False


# ---------------------------------------------------------------------------
# Group detection
# ---------------------------------------------------------------------------

def _in_group(collect_as: str, patterns: list[str]) -> bool:
    ca_lower = collect_as.lower()
    return any(p in ca_lower for p in patterns)


def _is_sales_full(collect_as: str) -> bool:
    return _in_group(collect_as, SALES_FULL)


def _is_sales_dominium(collect_as: str) -> bool:
    return _in_group(collect_as, SALES_PARTIAL)


def _is_greystar(collect_as: str) -> bool:
    return GREYSTAR in collect_as.lower()


def _is_withheld(collect_as: str) -> bool:
    ca_lower = collect_as.lower()
    for code in WITHHELD_CODES:
        if code in ca_lower:
            return True
    return False


# ---------------------------------------------------------------------------
# Status derivation
# ---------------------------------------------------------------------------

def _derive_status(account: AccountSummary) -> str:
    """
    Derive suggested collection status from bucket balances, category, escalation, etc.
    Returns status string.
    """
    b = account.balances
    category = account.category.strip()
    escalation = account.collection_escalation_status.strip()
    orig_status = account.current_collections_status.strip()

    # 1. Autopay — no status
    if account.is_autopay:
        return "No Action"

    # 2. Special Circumstances
    if orig_status.lower() == "special circumstances":
        return "Special Circumstances"

    # 3. Sales full group
    if _is_sales_full(account.collect_as):
        return "Call Customer"

    # Check if there's any aged balance
    oldest = account.oldest_bucket_with_balance
    if not oldest or oldest == "Current":
        return "No Action"

    cat_lower = category.lower()
    is_maintenance = "maintenance" in cat_lower
    is_wm = cat_lower == "wm"
    is_sar = "service and repair" in cat_lower or cat_lower == "service & repair"

    is_poolsure = account.business_unit.strip().lower() == "poolsure"
    is_aquasol = not is_poolsure

    # 3. Maintenance + 90+/120+/150+/180+ → Cancel NP
    if is_maintenance and oldest in ("90+", "120+", "150+", "180+"):
        return "Cancel NP"

    # 4. WM or Maintenance, oldest = 60+, with Current+30++60+ all > 0
    if (is_wm or is_maintenance) and oldest == "60+":
        if b["Current"] > 0.01 and b["30+"] > 0.01 and b["60+"] > 0.01:
            status = "Call Customer"
        else:
            status = "Skipped Invoice"
    else:
        # 5. Standard bucket → status
        if oldest == "30+":
            status = "Friendly Reminder Email"
        elif oldest == "60+":
            status = "Call Customer"
        elif oldest == "90+":
            status = "Collection Letter"
        elif oldest in ("120+", "150+"):
            if is_aquasol:
                status = "Collection Letter"
            else:  # Poolsure
                status = "Predemand Letter"
        elif oldest == "180+":
            if is_aquasol:
                status = "Collection Letter"
            else:  # Poolsure
                aged_180 = b["180+"]
                if aged_180 >= 50.0:
                    status = "Cancel NP"
                else:
                    status = "To be Determined"
        else:
            status = "No Action"

    # 6. Short Leash escalation (advance one tier)
    if escalation.lower() == "short leash":
        status = _escalate_status(status)

    # 7. Service And Repair — advance one tier
    if is_sar:
        status = _escalate_status(status)

    # 8. Small balance rule: if oldest bucket balance ≤ $25, step down one tier
    oldest_bal = b.get(oldest, 0.0) if oldest else 0.0
    if oldest_bal <= 25.0 and oldest_bal > 0.01:
        status = _deescalate_status(status)

    return status


_STATUS_UP = {
    "Friendly Reminder Email": "Call Customer",
    "Call Customer": "Collection Letter",
    "Collection Letter": "Predemand Letter",
    "Predemand Letter": "Cancel NP",
    "Cancel NP": "Cancel NP",
    "Skipped Invoice": "Call Customer",
}

_STATUS_DOWN = {
    "Cancel NP": "Predemand Letter",
    "Predemand Letter": "Collection Letter",
    "Collection Letter": "Call Customer",
    "Call Customer": "Friendly Reminder Email",
    "Friendly Reminder Email": "Friendly Reminder Email",
}


def _escalate_status(status: str) -> str:
    return _STATUS_UP.get(status, status)


def _deescalate_status(status: str) -> str:
    return _STATUS_DOWN.get(status, status)


# ---------------------------------------------------------------------------
# Credit waterfall
# ---------------------------------------------------------------------------

def apply_credit_waterfall(account: AccountSummary, credits: list[float]) -> None:
    """
    Apply credits oldest-to-newest across buckets.
    Modifies account.balances in place and sets credit_adjustment.
    """
    if not credits:
        return

    total_credit = sum(abs(c) for c in credits)
    account.total_open_credits = total_credit

    if total_credit <= 0.01:
        return

    remaining_credit = total_credit
    adjusted_buckets = []

    waterfall_order = ["180+", "150+", "120+", "90+", "60+", "30+", "Current"]
    for bucket in waterfall_order:
        if remaining_credit <= 0.01:
            break
        bal = account.balances[bucket]
        if bal <= 0.01:
            continue
        apply = min(remaining_credit, bal)
        account.balances[bucket] -= apply
        if apply > 0.01:
            adjusted_buckets.append(bucket)
        remaining_credit -= apply

    account.credit_adjustment = ", ".join(adjusted_buckets) if adjusted_buckets else ""


# ---------------------------------------------------------------------------
# Future restriction logic
# ---------------------------------------------------------------------------

def _compute_future_restriction(account: AccountSummary) -> str:
    ca = account.collect_as

    if account.is_autopay:
        return ""

    if _is_withheld(ca) or _is_greystar(ca):
        return "Withheld"

    if _is_sales_full(ca):
        return "Sales"

    if _is_sales_dominium(ca):
        return "Sales"

    escalation = account.collection_escalation_status.strip().lower()
    status = account.suggested_status
    total_aged = account.total_aged_balance

    if escalation == "short leash":
        # Tiered restriction warning for short leash accounts:
        #   30+ DSO: $100+ total aged balance
        #   60+ DSO: $50+  total aged balance
        #   90+/120+/150+/180+ DSO: any balance
        if oldest in ("90+", "120+", "150+", "180+") and total_aged > 0.01:
            return "Short Leash"
        elif oldest == "60+" and total_aged >= 50.0:
            return "Short Leash"
        elif oldest == "30+" and total_aged >= 100.0:
            return "Short Leash"

    cat_lower = account.category.strip().lower()
    is_sar = "service and repair" in cat_lower or cat_lower == "service & repair"
    is_long_leash = escalation == "long leash"
    is_regular = escalation in ("regular", "normal", "")

    if account.credit_adjustment and account.total_aged_balance <= 0.01:
        return ""  # Cleared — credits eliminated all aged balance

    oldest = account.oldest_bucket_with_balance
    aged_buckets_set = set(AGED_BUCKETS)

    if is_sar:
        # Tiered restriction warning for SAR accounts:
        #   30+ DSO: $100+ total aged balance
        #   60+ DSO: $50+  total aged balance
        #   90+/120+/150+/180+ DSO: any balance
        if oldest in ("90+", "120+", "150+", "180+") and total_aged > 0.01:
            return "Warning"
        elif oldest == "60+" and total_aged >= 50.0:
            return "Warning"
        elif oldest == "30+" and total_aged >= 100.0:
            return "Warning"

    if is_long_leash:
        # Tiered restriction warning for long leash accounts:
        #   90+ DSO:       $100+ total aged balance
        #   120+ DSO:      $50+  total aged balance
        #   150+/180+ DSO: any balance
        if oldest in ("150+", "180+") and total_aged > 0.01:
            return "Warning"
        elif oldest == "120+" and total_aged >= 50.0:
            return "Warning"
        elif oldest == "90+" and total_aged >= 100.0:
            return "Warning"

    if is_regular:
        # Tiered restriction warning thresholds:
        #   60+ DSO: $100+ total aged balance
        #   90+ DSO: $50+  total aged balance
        #   120+/150+/180+ DSO: any balance
        if oldest in ("120+", "150+", "180+") and total_aged > 0.01:
            return "Warning"
        elif oldest == "90+" and total_aged >= 50.0:
            return "Warning"
        elif oldest == "60+" and total_aged >= 100.0:
            return "Warning"

    return ""


# ---------------------------------------------------------------------------
# Rep assignment
# ---------------------------------------------------------------------------

def assign_reps(accounts: list[AccountSummary]) -> None:
    """
    Greedy bin-packing: split actionable non-special accounts between
    Rosas and Quilantan keeping account prefixes together.
    """
    actionable_unassigned = [
        a for a in accounts
        if a.section == "actionable"
        and a.collections_assignment == ""
        and a.suggested_status not in ("No Action", "Special Circumstances")
        and not a.is_autopay
        and not _is_greystar(a.collect_as)
        and not _is_sales_full(a.collect_as)
    ]

    # Group by first word of collect_as
    prefix_groups: dict[str, list[AccountSummary]] = {}
    for a in actionable_unassigned:
        prefix = a.collect_as.split()[0].lower() if a.collect_as else "?"
        prefix_groups.setdefault(prefix, []).append(a)

    # Sort prefixes by name for determinism
    sorted_prefixes = sorted(prefix_groups.keys())

    rosas_accounts: list[AccountSummary] = []
    quilantan_accounts: list[AccountSummary] = []

    # Dominium → could go to either; try to keep roughly equal
    for prefix in sorted_prefixes:
        group = prefix_groups[prefix]
        # Check if this prefix includes dominium
        rep_choice = _pick_rep_for_group(group, rosas_accounts, quilantan_accounts)
        for a in group:
            a.collections_assignment = rep_choice
        if rep_choice == "Rosas, Yoniva":
            rosas_accounts.extend(group)
        else:
            quilantan_accounts.extend(group)


def _pick_rep_for_group(
    group: list[AccountSummary],
    rosas: list[AccountSummary],
    quilantan: list[AccountSummary],
) -> str:
    """Pick the rep with the shorter current list."""
    if len(rosas) <= len(quilantan):
        return "Rosas, Yoniva"
    else:
        return "Quilantan, Maria"


# ---------------------------------------------------------------------------
# Main processing function
# ---------------------------------------------------------------------------

def process_collections(
    invoices: list[dict],
    credits: list[dict],
    report_month: int,
    report_year: int,
) -> list[AccountSummary]:
    """
    Full processing pipeline. Returns list of AccountSummary objects.
    """
    bucket_ranges = compute_bucket_ranges(report_year, report_month)

    # --- Aggregate invoices by account ---
    account_map: dict[str, AccountSummary] = {}

    for inv in invoices:
        if _should_exclude_invoice(inv):
            continue

        ca = inv.get("collect_as", "").strip()
        if not ca:
            continue

        if ca not in account_map:
            account_map[ca] = AccountSummary(
                collect_as=ca,
                business_unit=inv.get("business_unit", ""),
                category=inv.get("category", ""),
                current_collections_status=inv.get("collections_status", ""),
                collection_escalation_status=inv.get("collection_escalation_status", ""),
                account_restricted=inv.get("account_restricted", ""),
                fortis_autopay_enrollment=inv.get("fortis_autopay_enrollment", ""),
            )

        acc = account_map[ca]
        acc.invoice_count += 1

        # Set autopay if any invoice has enrollment
        if inv.get("fortis_autopay_enrollment", "").strip():
            acc.is_autopay = True
            acc.fortis_autopay_enrollment = inv["fortis_autopay_enrollment"]

        # Prefer non-empty category / business unit from first available
        if not acc.business_unit and inv.get("business_unit"):
            acc.business_unit = inv["business_unit"]
        if not acc.category and inv.get("category"):
            acc.category = inv["category"]
        if not acc.current_collections_status and inv.get("collections_status"):
            acc.current_collections_status = inv["collections_status"]
        if not acc.collection_escalation_status and inv.get("collection_escalation_status"):
            acc.collection_escalation_status = inv["collection_escalation_status"]
        if inv.get("account_restricted", "").lower() == "yes":
            acc.account_restricted = "Yes"

        amount = inv.get("amount_remaining", 0.0)
        if amount <= 0.0:
            continue  # skip zero/negative invoices

        # Finance charges
        if inv.get("is_finance_charge", "").strip().lower() == "yes":
            acc.balances["Finance Charges"] += amount
            acc.raw_balances["Finance Charges"] += amount
            continue

        # Date-based bucketing
        inv_date = _parse_date(inv.get("date", ""))
        if inv_date is None:
            continue

        bucket = date_to_bucket(inv_date, bucket_ranges)
        if bucket == "Future":
            continue  # Future invoices excluded from balances

        acc.balances[bucket] += amount
        acc.raw_balances[bucket] += amount

    # --- Group credits by account ---
    credit_map: dict[str, list[float]] = {}
    for cr in credits:
        ca = cr.get("collect_as", "").strip()
        if not ca:
            continue
        amt = cr.get("amount_remaining", 0.0)
        credit_map.setdefault(ca, []).append(amt)

    # --- Apply credit waterfall ---
    for ca, acc in account_map.items():
        if ca in credit_map:
            apply_credit_waterfall(acc, credit_map[ca])

    # --- Derive status and section ---
    accounts = list(account_map.values())

    for acc in accounts:
        # Save raw balances (pre-waterfall for reference)
        for b in BUCKETS:
            acc.raw_balances[b] = acc.balances[b]  # at this point balances are post-waterfall

        # Determine suggested status
        acc.suggested_status = _derive_status(acc)

        # Sales override: Call Customer
        if _is_sales_full(acc.collect_as) or _is_sales_dominium(acc.collect_as):
            acc.suggested_status = "Call Customer"

        # Compute future restriction
        acc.future_restriction = _compute_future_restriction(acc)

        # Pre-assign special accounts
        if acc.current_collections_status.strip().lower() == "special circumstances":
            acc.collections_assignment = "Wharton, Nancy"
        elif acc.suggested_status == "Cancel NP":
            # Maintenance Cancel NP → Wharton
            if "maintenance" in acc.category.lower():
                acc.collections_assignment = "Wharton, Nancy"
        elif _is_sales_full(acc.collect_as):
            acc.collections_assignment = ""  # Unassigned (Sales)
        elif _is_greystar(acc.collect_as):
            acc.collections_assignment = ""

        # Determine section
        total_aged = acc.total_aged_balance
        if acc.is_autopay:
            if total_aged > 0.01:
                acc.section = "autopay"
            else:
                acc.section = "current_only"
        elif total_aged <= 0.01 and acc.total_open_balance > 0.01:
            acc.section = "current_only"
        elif acc.total_open_balance <= 0.01 and not acc.total_open_credits:
            acc.section = "current_only"
        else:
            acc.section = "actionable"

    # --- Assign reps ---
    assign_reps(accounts)

    # --- Sort ---
    def sort_key(a: AccountSummary):
        # section order: actionable first, then current_only, then autopay
        section_order = {"actionable": 0, "current_only": 1, "autopay": 2}
        rep_order = {
            "Wharton, Nancy": 0,
            "Rosas, Yoniva": 1,
            "Quilantan, Maria": 2,
            "": 3,
        }
        sev = STATUS_SEVERITY.get(a.suggested_status, 9)
        return (
            section_order.get(a.section, 9),
            rep_order.get(a.collections_assignment, 3),
            sev,
            a.collect_as.lower(),
        )

    accounts.sort(key=sort_key)
    return accounts
