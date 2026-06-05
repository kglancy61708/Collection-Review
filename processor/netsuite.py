"""
NetSuite Token-Based Authentication (OAuth 1.0a) data pull.
Uses SuiteQL REST API with HMAC-SHA256 signed requests.

Credentials are read from environment variables:
  NS_ACCOUNT_ID       — e.g. 3412280
  NS_CONSUMER_KEY     — from Integration record
  NS_CONSUMER_SECRET  — from Integration record
  NS_TOKEN_ID         — from Access Token record
  NS_TOKEN_SECRET     — from Access Token record

NOTE: Custom field internal IDs (custbody_*, custentity_*) below must match
your NetSuite account configuration. If a column returns empty, check the
field ID against Setup → Customization → Transaction Body Fields in NetSuite.
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import logging
import os
import time
import urllib.parse
import uuid

import requests

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
PAGE_SIZE = 1000  # NetSuite max rows per SuiteQL page
# All credentials are read live from os.environ inside _build_auth_header()
# and _suiteql_request() — never cached at import time.

# ---------------------------------------------------------------------------
# Custom field IDs — set via Railway environment variables.
# Find internal IDs in NetSuite: Setup → Customization → Transaction Body
# Fields (for custbody_*) or Entity Fields (for custentity_*).
# If a variable is not set the column will return empty strings '' in SQL.
# ---------------------------------------------------------------------------
def _build_invoice_query() -> str:
    """Build the invoice SuiteQL query using live env var field IDs."""
    def col(env_var: str, alias: str, on_customer: bool = False) -> str:
        fid = os.environ.get(env_var, "")
        if fid:
            if on_customer:
                return f'  BUILTIN.DF(c.{fid})  AS "{alias}"'
            return f'  t.{fid}  AS "{alias}"'
        return f'  \'\'  AS "{alias}"'

    category_fid = os.environ.get("NS_FIELD_CATEGORY", "")
    category_col = f'  BUILTIN.DF(t.{category_fid})  AS "Category"' if category_fid \
                   else '  \'\'  AS "Category"'

    return f"""
SELECT
  c.altname || ' : ' || c.entityid          AS "Collect As",
  TO_CHAR(t.trandate, 'MM/DD/YYYY')          AS "Date",
  t.foreignamountunpaid                       AS "Amount Remaining",
  BUILTIN.DF(t.subsidiary)                   AS "Business Unit",
{category_col},
{col("NS_FIELD_COLLECTIONS_STATUS",       "Collections Status")},
{col("NS_FIELD_IS_FINANCE_CHARGE",        "Is Finance Charge")},
{col("NS_FIELD_COLLECTION_ESCALATION",    "Collection Escalation Status")},
{col("NS_FIELD_FORTIS_AUTOPAY",           "Fortis Autopay Enrollment")},
{col("NS_FIELD_ACCOUNT_RESTRICTED",       "Account Restricted", on_customer=True)}
FROM transaction t
INNER JOIN customer c ON t.entity = c.id
WHERE t.type = 'CustInvc'
  AND COALESCE(t.foreignamountunpaid, 0) > 0
ORDER BY c.entityid, t.trandate
"""


def _build_credit_query() -> str:
    # foreignamountunpaid on credit memos may be negative (amount owed to customer)
    # or positive depending on NS configuration — use ABS to handle both.
    return """
SELECT
  c.altname || ' : ' || c.entityid   AS "Collect As",
  t.foreignamountunpaid               AS "Amount Remaining",
  BUILTIN.DF(t.customform)            AS "Custom Form"
FROM transaction t
INNER JOIN customer c ON t.entity = c.id
WHERE t.type = 'CustCred'
  AND ABS(t.foreignamountunpaid) > 0
ORDER BY c.entityid
"""


# ---------------------------------------------------------------------------
# OAuth 1.0a signing
# ---------------------------------------------------------------------------

def _pct_encode(s: str) -> str:
    """RFC 3986 percent-encode (encodes ! ' ( ) * as well)."""
    return urllib.parse.quote(str(s), safe="")


def _build_auth_header(method: str, url: str, query_params: dict | None = None) -> str:
    """
    Build a signed OAuth 1.0a Authorization header.

    query_params: any URL query string parameters (e.g. {"limit": "1000", "offset": "0"})
    that must be included in the signature base string.
    """
    # Read credentials live from os.environ so Railway vars are always current
    consumer_key    = os.environ.get("NS_CONSUMER_KEY",    "")
    consumer_secret = os.environ.get("NS_CONSUMER_SECRET", "")
    token_id        = os.environ.get("NS_TOKEN_ID",        "")
    token_secret    = os.environ.get("NS_TOKEN_SECRET",    "")
    account_id      = os.environ.get("NS_ACCOUNT_ID",      "3412280")

    oauth_params: dict[str, str] = {
        "oauth_consumer_key":     consumer_key,
        "oauth_nonce":            uuid.uuid4().hex,
        "oauth_signature_method": "HMAC-SHA256",
        "oauth_timestamp":        str(int(time.time())),
        "oauth_token":            token_id,
        "oauth_version":          "1.0",
    }

    # Collect all params for signature: oauth params + query string params
    all_params: dict[str, str] = {**oauth_params}
    if query_params:
        all_params.update({str(k): str(v) for k, v in query_params.items()})

    # Parameter string: sorted by encoded key, then encoded value
    param_pairs = sorted(
        (_pct_encode(k), _pct_encode(v)) for k, v in all_params.items()
    )
    param_string = "&".join(f"{k}={v}" for k, v in param_pairs)

    # Base URL: scheme + host + path only (no query string)
    parsed = urllib.parse.urlparse(url)
    base_url = urllib.parse.urlunparse(
        (parsed.scheme, parsed.netloc, parsed.path, "", "", "")
    )

    # Signature base string
    base_string = "&".join([
        _pct_encode(method.upper()),
        _pct_encode(base_url),
        _pct_encode(param_string),
    ])

    # Signing key
    signing_key = f"{_pct_encode(consumer_secret)}&{_pct_encode(token_secret)}"

    # HMAC-SHA256
    raw_sig = hmac.new(
        signing_key.encode("utf-8"),
        base_string.encode("utf-8"),
        hashlib.sha256,
    ).digest()
    oauth_params["oauth_signature"] = base64.b64encode(raw_sig).decode()

    # Build header value — realm uses uppercase account ID
    realm = account_id.upper().replace("-", "_")
    header_parts = [f'realm="{realm}"'] + [
        f'{k}="{_pct_encode(v)}"'
        for k, v in sorted(oauth_params.items())
    ]
    return "OAuth " + ", ".join(header_parts)


# ---------------------------------------------------------------------------
# HTTP request with retry / backoff
# ---------------------------------------------------------------------------

def _suiteql_request(query: str, offset: int = 0) -> dict:
    """POST one page of a SuiteQL query. Retries on 429 with backoff."""
    account_id = os.environ.get("NS_ACCOUNT_ID", "3412280")
    suiteql_url = (
        f"https://{account_id}.suitetalk.api.netsuite.com"
        f"/services/rest/query/v1/suiteql"
    )
    params = {"limit": str(PAGE_SIZE), "offset": str(offset)}

    for attempt in range(5):
        auth = _build_auth_header("POST", suiteql_url, query_params=params)
        try:
            resp = requests.post(
                suiteql_url,
                params=params,
                json={"q": query.strip()},
                headers={
                    "Authorization": auth,
                    "Content-Type":  "application/json",
                    "prefer":        "transient",
                },
                timeout=(15, 90),  # (connect timeout, read timeout) in seconds
            )
        except requests.RequestException as e:
            raise RuntimeError(f"Network error calling SuiteQL: {e}")

        if resp.status_code == 429:
            wait = 3 * (attempt + 1)
            logger.warning("NS rate limit (429), retrying in %ds…", wait)
            time.sleep(wait)
            continue

        if resp.status_code == 401:
            raise RuntimeError(
                "NetSuite returned 401 Unauthorized. "
                "Check that NS_CONSUMER_KEY, NS_CONSUMER_SECRET, NS_TOKEN_ID, "
                "and NS_TOKEN_SECRET environment variables are set correctly."
            )

        if resp.status_code not in (200, 204):
            body = resp.text[:1000]
            logger.error("SuiteQL HTTP %s — query: %s — response: %s",
                         resp.status_code, query.strip()[:300], body)
            raise RuntimeError(
                f"SuiteQL returned HTTP {resp.status_code}: {body}"
            )

        result = resp.json()
        logger.info("SuiteQL page OK — %d items returned", len(result.get("items", [])))
        return result

    raise RuntimeError("NetSuite SuiteQL failed after 5 retries (rate limited).")


# ---------------------------------------------------------------------------
# Paginated query runner
# ---------------------------------------------------------------------------

def _run_query(query: str, label: str,
               progress_cb=None) -> list[dict]:
    """
    Run a paginated SuiteQL query, return all rows as list of dicts.
    progress_cb(message: str) is called after each page if provided.
    """
    rows: list[dict] = []
    offset = 0
    total = None

    while True:
        logger.info("Fetching %s rows %d–%d…", label, offset, offset + PAGE_SIZE)
        data = _suiteql_request(query, offset=offset)

        items = data.get("items", [])
        rows.extend(items)

        if total is None:
            total = data.get("totalResults", 0)

        if progress_cb and total:
            pct = min(100, int(len(rows) / total * 100))
            progress_cb(f"Downloading {label}: {len(rows):,} / {total:,} rows ({pct}%)…")

        offset += PAGE_SIZE
        if offset >= total or not items:
            break

    logger.info("Fetched %d total %s rows.", len(rows), label)
    return rows


# ---------------------------------------------------------------------------
# Column normalisation helpers
# ---------------------------------------------------------------------------

def _normalise_invoice_rows(rows: list[dict]) -> list[dict]:
    """
    Convert SuiteQL result rows to the normalised invoice dict format
    expected by processor.logic.process_collections().
    """
    result = []
    for row in rows:
        def g(key: str, *fallbacks: str) -> str:
            for k in (key, *fallbacks):
                v = row.get(k) or row.get(k.lower()) or row.get(k.upper()) or ""
                if v is not None:
                    return str(v).strip()
            return ""

        amount_str = g("Amount Remaining", "amountremaining", "amountRemaining")
        try:
            amount = float(amount_str.replace(",", ""))
        except (ValueError, AttributeError):
            amount = 0.0

        record = {
            "collect_as":                    g("Collect As"),
            "date":                          g("Date"),
            "amount_remaining":              amount,
            "business_unit":                 g("Business Unit"),
            "category":                      g("Category"),
            "collections_status":            g("Collections Status"),
            "is_finance_charge":             g("Is Finance Charge"),
            "collection_escalation_status":  g("Collection Escalation Status"),
            "fortis_autopay_enrollment":     g("Fortis Autopay Enrollment"),
            "account_restricted":            g("Account Restricted"),
        }
        if record["collect_as"]:
            result.append(record)
    return result


def _normalise_credit_rows(rows: list[dict]) -> list[dict]:
    """
    Convert SuiteQL result rows to the normalised credit dict format.
    SRDP credits are excluded here.
    """
    result = []
    for row in rows:
        def g(key: str, *fallbacks: str) -> str:
            for k in (key, *fallbacks):
                v = row.get(k) or row.get(k.lower()) or ""
                if v is not None:
                    return str(v).strip()
            return ""

        custom_form = g("Custom Form", "customform")
        if "srdp" in custom_form.lower():
            continue

        amount_str = g("Amount Remaining", "amountremaining")
        try:
            amount = float(amount_str.replace(",", ""))
        except (ValueError, AttributeError):
            amount = 0.0

        record = {
            "collect_as":       g("Collect As"),
            "amount_remaining": amount,
            "custom_form":      custom_form,
        }
        if record["collect_as"]:
            result.append(record)
    return result


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def check_credentials() -> bool:
    """Return True if all required env vars are set. Reads os.environ live."""
    return all([
        os.environ.get("NS_CONSUMER_KEY", ""),
        os.environ.get("NS_CONSUMER_SECRET", ""),
        os.environ.get("NS_TOKEN_ID", ""),
        os.environ.get("NS_TOKEN_SECRET", ""),
    ])


def diagnose() -> dict:
    """
    Run diagnostic SuiteQL queries to help identify why invoices return 0 rows.
    Returns a dict with counts and sample data for review in Railway logs.
    """
    results = {}

    tests = [
        ("total_transactions",   "SELECT COUNT(*) AS cnt FROM transaction"),
        ("invoice_types",        "SELECT type, COUNT(*) AS cnt FROM transaction GROUP BY type ORDER BY cnt DESC"),
        ("custinvc_count",       "SELECT COUNT(*) AS cnt FROM transaction WHERE type = 'CustInvc'"),
        ("custinvc_any_amount",  "SELECT COUNT(*) AS cnt FROM transaction WHERE type = 'CustInvc' AND COALESCE(foreignamountunpaid, 0) > 0"),
        ("custcred_count",       "SELECT COUNT(*) AS cnt FROM transaction WHERE type = 'CustCred'"),
        ("customer_count",       "SELECT COUNT(*) AS cnt FROM customer"),
    ]

    for label, query in tests:
        try:
            data = _suiteql_request(query, offset=0)
            items = data.get("items", [])
            results[label] = items
            logger.info("DIAG %s: %s", label, items)
        except Exception as e:
            results[label] = {"error": str(e)}
            logger.error("DIAG %s FAILED: %s", label, e)

    return results


def pull_netsuite_data(progress_cb=None) -> tuple[list[dict], list[dict]]:
    """
    Pull open invoices and credits from NetSuite via SuiteQL.
    Returns (invoice_rows, credit_rows) as normalised dicts ready for
    processor.logic.process_collections().

    Raises RuntimeError with a descriptive message on failure.
    """
    if not check_credentials():
        raise RuntimeError(
            "NetSuite credentials are not configured. "
            "Set NS_CONSUMER_KEY, NS_CONSUMER_SECRET, NS_TOKEN_ID, and "
            "NS_TOKEN_SECRET as environment variables on your Railway service."
        )

    raw_invoices = _run_query(_build_invoice_query(), "invoices", progress_cb=progress_cb)
    raw_credits  = _run_query(_build_credit_query(),  "credits",  progress_cb=progress_cb)

    invoices = _normalise_invoice_rows(raw_invoices)
    credits  = _normalise_credit_rows(raw_credits)

    if not invoices:
        logger.warning(
            "SuiteQL returned 0 invoice rows. Check role permissions and "
            "custom field IDs (F_COLLECTIONS_STATUS etc.) in processor/netsuite.py."
        )

    return invoices, credits
