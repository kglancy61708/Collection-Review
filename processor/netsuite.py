"""
NetSuite session-based data pull.
Authenticates via the customer login page and downloads saved search CSVs.
"""
import requests
import logging

logger = logging.getLogger(__name__)

INVOICE_SEARCH_ID = "7474"
CREDIT_SEARCH_ID = "6011"
LOGIN_URL = "https://system.netsuite.com/pages/customerlogin.jsp"


def pull_netsuite_data(account_id: str, email: str, password: str) -> tuple[str, str]:
    """
    Authenticate with NetSuite and download both saved searches.
    Returns (invoice_csv_text, credits_csv_text).
    Raises RuntimeError with descriptive message on failure.
    """
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        )
    })

    # Step 1: Authenticate
    login_payload = {
        "email": email,
        "password": password,
        "account": account_id.strip(),
        "redirect": "index.html",
        "machine": "",
        "trusteddevice": "F",
    }

    try:
        login_resp = session.post(LOGIN_URL, data=login_payload, allow_redirects=True, timeout=30)
    except requests.RequestException as e:
        raise RuntimeError(f"Network error during NetSuite login: {e}")

    # Detect login failure by looking for known failure indicators
    body = login_resp.text.lower()
    if (
        "invalid email address or password" in body
        or "your e-mail address or password is incorrect" in body
        or "customerlogin" in login_resp.url.lower()
        and "error" in body
    ):
        raise RuntimeError(
            "NetSuite authentication failed. Check your account ID, email, and password."
        )

    if login_resp.status_code not in (200, 302):
        raise RuntimeError(
            f"NetSuite login returned unexpected status {login_resp.status_code}."
        )

    # Step 2: Download saved searches
    base_url = f"https://{account_id.strip()}.app.netsuite.com"
    invoice_url = (
        f"{base_url}/app/common/search/searchresults.nl"
        f"?searchid={INVOICE_SEARCH_ID}&whence=&csv=T"
    )
    credit_url = (
        f"{base_url}/app/common/search/searchresults.nl"
        f"?searchid={CREDIT_SEARCH_ID}&whence=&csv=T"
    )

    invoice_csv = _download_search(session, invoice_url, "invoices")
    credit_csv = _download_search(session, credit_url, "credits")

    return invoice_csv, credit_csv


def _download_search(session: requests.Session, url: str, label: str) -> str:
    """Download a saved search CSV and return its text content."""
    try:
        resp = session.get(url, timeout=60, allow_redirects=True)
    except requests.RequestException as e:
        raise RuntimeError(f"Network error downloading {label} search: {e}")

    if resp.status_code == 403:
        raise RuntimeError(
            f"Access denied downloading {label} search. "
            "The NetSuite session may have expired or the saved search ID is incorrect."
        )
    if resp.status_code != 200:
        raise RuntimeError(
            f"Failed to download {label} search: HTTP {resp.status_code}"
        )

    # Detect redirect back to login (session not established)
    content_type = resp.headers.get("Content-Type", "").lower()
    if "text/html" in content_type and "<html" in resp.text[:200].lower():
        raise RuntimeError(
            f"NetSuite returned an HTML page instead of {label} CSV data. "
            "Authentication may have failed or timed out."
        )

    return resp.text
