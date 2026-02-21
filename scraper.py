"""Search and cable extraction module for LeakQuest."""

import json
import re
import subprocess
import tempfile
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote_plus

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://wikileaks.org/plusd"
SEARCH_URL = f"{BASE_URL}/"
SPHINX_URL = f"{BASE_URL}/sphinxer_do.php"
CABLE_URL = f"{BASE_URL}/cables"

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}

# Retry config
MAX_RETRIES = 3
BACKOFF_SECONDS = [5, 10, 20]


def _create_session():
    """Create a requests session with default headers."""
    session = requests.Session()
    session.headers.update(DEFAULT_HEADERS)
    return session


def _curl_fetch(url):
    """Fetch a URL using curl to avoid requests' URL encoding issues.

    The WikiLeaks PlusD server uses literal [] in query params and redirects.
    Python's requests library encodes these as %5B%5D, causing infinite redirects.
    curl handles this correctly.
    """
    with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        result = subprocess.run(
            ["curl", "-s", "-S", "-L", "--max-redirs", "5", "-o", tmp_path,
             "-A", DEFAULT_HEADERS["User-Agent"], url],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0:
            stderr = result.stderr.strip()
            raise ConnectionError(f"curl failed (exit {result.returncode}): {stderr or 'unknown error'}")
        return Path(tmp_path).read_text(encoding="utf-8", errors="replace")
    except subprocess.TimeoutExpired:
        raise ConnectionError("curl timed out after 60 seconds")
    except FileNotFoundError:
        raise ConnectionError("curl not found — ensure curl is installed and on PATH")
    finally:
        try:
            Path(tmp_path).unlink()
        except OSError:
            pass


def _request_with_retry(session, method, url, **kwargs):
    """Make an HTTP request with retry logic."""
    for attempt in range(MAX_RETRIES):
        try:
            resp = session.request(method, url, timeout=30, **kwargs)
            resp.raise_for_status()
            return resp
        except requests.RequestException as e:
            if attempt < MAX_RETRIES - 1:
                wait = BACKOFF_SECONDS[attempt]
                time.sleep(wait)
            else:
                raise


def _normalize_date(raw_date):
    """Normalize various WikiLeaks date formats to YYYY-MM-DD.

    Handles:
      - Cable page:    "1987 December 19, 20:12 (Saturday)"
      - Search listing: "Sat, 19 Dec 1987"
      - Already normalized: "1987-12-19"
    Returns empty string if parsing fails.
    """
    if not raw_date:
        return ""

    # Cable page format: "1987 December 19, 20:12 (Saturday)"
    m = re.match(r"(\d{4})\s+(\w+)\s+(\d{1,2})", raw_date)
    if m:
        try:
            dt = datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%Y %B %d")
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass

    # Search listing format: "Sat, 19 Dec 1987"
    m = re.match(r"\w+,\s+(\d{1,2})\s+(\w+)\s+(\d{4})", raw_date)
    if m:
        try:
            dt = datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%d %b %Y")
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass

    # Already YYYY-MM-DD (possibly with trailing content)
    m = re.search(r"\d{4}-\d{2}-\d{2}", raw_date)
    if m:
        return m.group(0)

    return ""


def _parse_result_rows(html):
    """Parse cable result rows from HTML content.

    Returns list of dicts with cable_id, title, date.
    """
    soup = BeautifulSoup(html, "html.parser")
    results = []
    for tr in soup.select("tr[id]"):
        cable_id = tr.get("id", "").strip()
        if not cable_id or not cable_id[0].isdigit():
            continue

        # Extract columns: Classification, Date, Subject, From, To, Length
        tds = tr.find_all("td")
        title = ""
        date = ""
        if len(tds) >= 3:
            date = _normalize_date(tds[1].get_text(strip=True))
            link = tds[2].find("a")
            title = link.get_text(strip=True) if link else tds[2].get_text(strip=True)

        results.append({
            "cable_id": cable_id,
            "title": title,
            "date": date,
        })
    return results


def _js_obj_to_json(js_obj):
    """Convert a simple JavaScript object literal to valid JSON.

    Handles: unquoted keys, single-quoted values, trailing commas.
    Preserves colons inside string values (e.g. URLs).
    """
    result = []
    i = 0
    s = js_obj
    while i < len(s):
        ch = s[i]

        # Skip whitespace
        if ch in " \t\r\n":
            result.append(ch)
            i += 1
            continue

        # Pass through structural chars
        if ch in "{},:":
            result.append(ch)
            i += 1
            continue

        # String literal (single or double quoted)
        if ch in ('"', "'"):
            quote = ch
            j = i + 1
            while j < len(s) and s[j] != quote:
                if s[j] == "\\":
                    j += 1  # skip escaped char
                j += 1
            # Extract string content (without quotes), output as double-quoted
            content = s[i + 1 : j]
            if quote == "'":
                # Unescape \' (not valid in JSON) and escape " (needed in JSON)
                content = content.replace("\\'", "'")
                content = content.replace('"', '\\"')
            result.append('"' + content + '"')
            i = j + 1
            continue

        # Bare word (unquoted key or value like true/false/null/number)
        j = i
        while j < len(s) and s[j] not in " \t\r\n{},:\"'":
            j += 1
        word = s[i:j]
        # Check if this is followed by a colon (it's a key) — quote it
        rest = s[j:].lstrip()
        if rest and rest[0] == ":":
            result.append('"' + word + '"')
        else:
            # It's a value — numbers, true, false, null stay unquoted
            result.append(word)
        i = j
        continue

    json_str = "".join(result)
    # Remove trailing commas before closing braces
    json_str = re.sub(r",\s*}", "}", json_str)
    return json_str


def _extract_page_parameters(html):
    """Extract page_parameters and result_token from the search page JavaScript.

    Returns (page_params_dict, result_token) or raises ValueError.
    """
    # Extract page_parameters object
    match = re.search(r"var\s+page_parameters\s*=\s*(\{.*?\})\s*;", html, re.DOTALL)
    if not match:
        raise ValueError("Could not find page_parameters in search page")

    js_obj = match.group(1)
    json_str = _js_obj_to_json(js_obj)

    try:
        page_params = json.loads(json_str)
    except json.JSONDecodeError:
        raise ValueError(f"Could not parse page_parameters as JSON: {json_str[:200]}")

    # Extract result_token
    token_match = re.search(r'result_token\s*=\s*["\']([^"\']*)["\']', html)
    result_token = token_match.group(1) if token_match else ""

    return page_params, result_token


def search_cables(keyword, date_from=None, date_to=None,
                  projects=None, progress_callback=None):
    """Search for cables matching criteria.

    Args:
        keyword: Search keyword(s)
        date_from: Start date string YYYY-MM-DD
        date_to: End date string YYYY-MM-DD
        projects: List of project codes, default ["cg"] (Cablegate)
        progress_callback: Optional callable(message) for status updates

    Returns:
        List of dicts with cable_id, title, date
    """
    if not keyword:
        raise ValueError("Keywords must be provided")

    if projects is None:
        projects = ["cg"]

    session = _create_session()

    # Build search URL with literal [] (requests would encode them as %5B%5D
    # which causes an infinite 302 redirect loop on the WikiLeaks server)
    query_parts = []
    query_parts.append(f"q={quote_plus(keyword)}")
    if date_from:
        query_parts.append(f"qtfrom={quote_plus(date_from)}")
    if date_to:
        query_parts.append(f"qtto={quote_plus(date_to)}")
    for proj in projects:
        query_parts.append(f"qproject[]={quote_plus(proj)}")
    search_url = SEARCH_URL + "?" + "&".join(query_parts)

    def _log(msg):
        if progress_callback:
            progress_callback(msg)

    _log("Searching WikiLeaks PlusD...")

    # The server returns a 302 redirect with #result appended and literal []
    # in the URL. Python's requests library always URL-encodes [] to %5B%5D,
    # which causes the server to redirect infinitely. We use curl instead,
    # which follows the redirect correctly with literal [].
    html = _curl_fetch(search_url)
    if not html:
        raise ConnectionError("Failed to fetch search results from WikiLeaks PlusD")

    # Parse initial results
    results = _parse_result_rows(html)
    _log(f"Found {len(results)} initial results")

    # Extract pagination parameters
    try:
        page_params, token = _extract_page_parameters(html)
    except ValueError as e:
        _log(f"Pagination info not found ({e}), returning initial results only")
        return results

    if not token:
        _log("No pagination token, returning initial results only")
        return results

    # Paginate to get all results
    qlimit = 500
    page_num = 1

    while True:
        time.sleep(1)  # Rate limit between pagination requests

        sphinx_params = {
            "format": "html",
            "command": "doc_list_from_query",
            "project": page_params.get("project", "all_cables"),
            "subp": page_params.get("subp", "cg"),
            "qcanonical": page_params.get("qcanonical", ""),
            "qcanonical_seal": page_params.get("qcanonical_seal", ""),
            "qsort": "tasc",
            "qlimit": qlimit,
            "token": token,
        }
        if date_from:
            sphinx_params["qtfrom"] = date_from
        if date_to:
            sphinx_params["qtto"] = date_to
        if page_params.get("s"):
            sphinx_params["s"] = page_params["s"]

        try:
            resp = _request_with_retry(session, "GET", SPHINX_URL, params=sphinx_params)
            data = resp.json()
        except (requests.RequestException, json.JSONDecodeError) as e:
            _log(f"Pagination request failed: {e}")
            break

        content = data.get("content", "")
        new_token = data.get("token", "")
        length = data.get("length", 0)

        if content:
            page_results = _parse_result_rows(content)
            results.extend(page_results)
            page_num += 1
            _log(f"Page {page_num}: +{len(page_results)} results (total: {len(results)})")

        if length < qlimit or not new_token:
            break

        token = new_token

    # Deduplicate by cable_id (in case of overlaps)
    seen = set()
    unique_results = []
    for r in results:
        if r["cable_id"] not in seen:
            seen.add(r["cable_id"])
            unique_results.append(r)

    _log(f"Search complete: {len(unique_results)} unique cables found")
    return unique_results


def fetch_cable(cable_id, session=None):
    """Fetch and extract data from a single cable page.

    Args:
        cable_id: The cable identifier (e.g. "1974BONN02212")
        session: Optional requests.Session to reuse

    Returns:
        Dict with cable_id, title, full_text
    """
    if session is None:
        session = _create_session()

    url = f"{CABLE_URL}/{cable_id}.html"
    resp = _request_with_retry(session, "GET", url)
    soup = BeautifulSoup(resp.text, "html.parser")

    # Extract title, date, and origin from synopsis table
    title = ""
    date = ""
    origin = ""
    synopsis = soup.select_one("#synopsis")
    if synopsis:
        # Title is in the first td with colspan
        title_td = synopsis.select_one("td[colspan]")
        if title_td:
            title = title_td.get_text(strip=True)
        else:
            # Fallback: first td in first tr
            first_td = synopsis.select_one("tr td")
            if first_td:
                title = first_td.get_text(strip=True)

        # Extract metadata from synopsis (div.s_key / div.s_val pairs)
        for row in synopsis.select("tr"):
            key_div = row.select_one("div.s_key")
            val_div = row.select_one("div.s_val")
            if not key_div or not val_div:
                continue
            key_text = key_div.get_text(strip=True).lower()
            if "date" in key_text and not date:
                date = _normalize_date(val_div.get_text(strip=True))
            elif key_text.startswith("from"):
                origin = val_div.get_text(strip=True)

    # Extract full text
    full_text = ""
    tagged_text = soup.select_one("#tagged-text")
    if tagged_text:
        full_text = tagged_text.get_text("\n", strip=True)

    result = {
        "cable_id": cable_id,
        "title": title,
        "date": date,
        "full_text": full_text,
    }
    if origin:
        result["origin"] = origin
    return result


def load_checkpoint(checkpoint_file):
    """Load checkpoint data from a JSON file.

    Returns:
        Dict with keyword, cable_ids, completed (dict of cable_id → data)
    """
    path = Path(checkpoint_file)
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"keyword": "", "cable_ids": [], "completed": {}}


def save_checkpoint(checkpoint_file, checkpoint_data):
    """Save checkpoint data to a JSON file."""
    with open(checkpoint_file, "w", encoding="utf-8") as f:
        json.dump(checkpoint_data, f, ensure_ascii=False)


def fetch_all_cables(cable_list, checkpoint_file="checkpoint.json", delay=1.5):
    """Fetch all cables with checkpoint/resume support.

    Args:
        cable_list: List of dicts with cable_id (and optionally title, date)
        checkpoint_file: Path to checkpoint JSON file
        delay: Seconds to wait between requests

    Yields:
        (index, total, cable_data_dict) for each fetched cable
    """
    checkpoint = load_checkpoint(checkpoint_file)
    completed = checkpoint.get("completed", {})

    # Store cable IDs in checkpoint for resume
    checkpoint["cable_ids"] = [c["cable_id"] for c in cable_list]

    session = _create_session()
    total = len(cable_list)

    for i, cable_info in enumerate(cable_list):
        cable_id = cable_info["cable_id"]

        # Skip already completed (but retry previously failed cables)
        if cable_id in completed and "_fetch_error" not in completed[cable_id]:
            yield (i + 1, total, completed[cable_id])
            continue

        # Fetch with retry
        try:
            cable_data = fetch_cable(cable_id, session)
        except requests.RequestException as e:
            cable_data = {
                "cable_id": cable_id,
                "title": cable_info.get("title", ""),
                "full_text": f"[ERROR: Failed to fetch - {e}]",
                "_fetch_error": str(e),
            }

        # Use search-result date as fallback if cable page had no date
        if not cable_data.get("date") and cable_info.get("date"):
            cable_data["date"] = cable_info["date"]

        # Save to checkpoint
        completed[cable_id] = cable_data
        checkpoint["completed"] = completed
        save_checkpoint(checkpoint_file, checkpoint)

        yield (i + 1, total, cable_data)

        # Rate limit
        if i < total - 1:
            time.sleep(delay)
