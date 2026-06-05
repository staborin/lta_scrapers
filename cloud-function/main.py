"""
refresh_entries — Google Cloud Function

Receives a POST request from the web app with either:

  A) Single tournament refresh:
     {"url": "https://competitions.lta.org.uk/sport/event.aspx?id=...&event=N",
      "sheet": "LUKA_TOURNAMENTS",
      "col": 0}

  B) Refresh all tournaments across all 6 sheets:
     {"action": "refreshAllEntries"}

For (A):
  1. Fetches the entries page with requests (no browser needed — plain HTML table)
  2. Parses entries: prefers "Entries (N)" table, falls back to "Online entries (N)"
  3. Clears old entry rows, writes fresh entries (row 10 onwards, correct col)
  4. Enriches watchlist sheets with player name if not present
  5. Triggers masterPopulate via Apps Script webhook

For (B):
  1. Reads all 6 sheets via Apps Script doGet
  2. Parses tournament blocks (3 cols wide, up to 10 blocks per sheet)
  3. Skips blocks where tournament end date is in the past
  4. Fires all tournament refreshes in parallel using ThreadPoolExecutor
  5. Each thread: clear old rows, fetch LTA entries, enrich if watchlist, write
  6. Calls masterPopulate once after all threads complete
"""

import os
import re
import json
import base64
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from urllib.parse import urlparse
from bs4 import BeautifulSoup
import functions_framework

WEBAPP_URL    = os.environ.get("LTA_SHEETS_WEBAPP_URL", "").strip()
WEBAPP_SECRET = os.environ.get("LTA_SHEETS_WEBAPP_SECRET", "").strip()

REFRESH_SHEETS = [
    "LUKA_TOURNAMENTS",
    "DYLAN_TOURNAMENTS",
    "SERGE_TOURNAMENTS",
    "LUKA_WATCHLIST",
    "DYLAN_WATCHLIST",
    "SERGE_WATCHLIST",
]

# Watchlist sheets require the player's name to be preserved in the entry list
# even if they are not yet officially on the LTA entries page.
WATCHLIST_PLAYERS = {
    "LUKA_WATCHLIST":  "Luka Taborin",
    "DYLAN_WATCHLIST": "Dylan Taborin",
    "SERGE_WATCHLIST": "Serge Taborin",
}


def post_to_webapp(payload: dict, timeout: int = 30) -> dict:
    """POST a payload to the Apps Script webapp."""
    url = WEBAPP_URL
    if WEBAPP_SECRET:
        joiner = "&" if "?" in url else "?"
        url = f"{url}{joiner}secret={WEBAPP_SECRET}"
    r = requests.post(url, json=payload, timeout=timeout)
    r.raise_for_status()
    return r.json()


def get_from_webapp(params: dict, timeout: int = 30) -> dict:
    """GET from the Apps Script webapp with query parameters."""
    url = WEBAPP_URL
    if WEBAPP_SECRET:
        params = {**params, "secret": WEBAPP_SECRET}
    r = requests.get(url, params=params, timeout=timeout)
    r.raise_for_status()
    return r.json()


def parse_sheet_date(date_str: str) -> date | None:
    """
    Parse a date string from the sheet.
    Accepts "DD/MM/YYYY", "DD/MM/YYYY to DD/MM/YYYY", "DD/MM/YYYY - DD/MM/YYYY",
    and dates with appended text like "DD/MM/YYYY (1.45pm)".
    Always returns the LAST date in the range (or the single date).
    Returns None if unparseable.
    """
    if not date_str or not date_str.strip():
        return None
    dates = re.findall(r'\b(\d{1,2})/(\d{1,2})/(\d{4})\b', date_str)
    if not dates:
        return None
    # Take the last date found (end of range)
    d, m, y = dates[-1]
    try:
        return date(int(y), int(m), int(d))
    except (ValueError, TypeError):
        return None


def fetch_entries(tournament_url: str) -> tuple[list[dict], bool]:
    """
    Fetch the LTA entries page and parse the entries table.
    Prefers 'Entries (N)' table (returns player names only, no dates).
    Falls back to 'Online entries (N)' table (returns names + dates).
    Returns tuple of (list of {"name": str, "entry_date": str}, used_entries_table: bool).
    """
    headers = {
        "User-Agent":                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept":                    "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language":           "en-GB,en;q=0.5",
        "Accept-Encoding":           "gzip, deflate, br",
        "Connection":                "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    }
    parsed   = urlparse(tournament_url)
    base_url = f"{parsed.scheme}://{parsed.netloc}"
    return_url = f"{parsed.path}?{parsed.query}" if parsed.query else parsed.path

    session = requests.Session()

    # Step 1: GET — establishes ASP.NET session and receives the cookie wall.
    session.get(tournament_url, headers=headers, timeout=30)

    # Step 2: POST to /cookiewall/Save exactly as the browser "Accept" button does.
    session.post(
        f"{base_url}/cookiewall/Save",
        data={
            "ReturnUrl":      return_url,
            "SettingsOpen":   "false",
            "CookiePurposes": ["2", "4", "8", "16"],
        },
        headers={
            **headers,
            "Referer":      tournament_url,
            "Origin":       base_url,
            "Content-Type": "application/x-www-form-urlencoded",
        },
        allow_redirects=False,
        timeout=30,
    )

    # Step 3: GET the tournament page — session now carries the consent cookie.
    resp = session.get(tournament_url, headers=headers, timeout=30)
    resp.raise_for_status()

    soup    = BeautifulSoup(resp.text, "html.parser")
    entries = []

    # ── Prefer "Entries (N)" table ──────────────────────────────────────────
    entries_table = None
    for table in soup.find_all("table", class_="ruler"):
        caption = table.find("caption")
        if caption and re.match(r"^\s*Entries\s*\(\d+\)\s*$", caption.get_text(), re.I):
            entries_table = table
            break

    if entries_table:
        seen = set()
        rows = entries_table.select("tbody tr")

        # PASS 1: Maindraw / Qualification rows — player in td[1]
        for tr in rows:
            tds = tr.find_all(["td", "th"])
            if len(tds) < 2:
                continue
            label = tds[0].get_text(strip=True).replace("\u00a0", " ").lower()
            if label.startswith("reserve"):
                continue
            if not (label.startswith("maindraw") or label.startswith("qualification")):
                continue
            a    = tds[1].find("a")
            name = a.get_text(strip=True) if a else tds[1].get_text(strip=True)
            if name and name not in seen:
                seen.add(name)
                entries.append({"name": name, "entry_date": ""})

        # PASS 2: fallback if pass 1 empty — qualification rows only
        if not entries:
            for tr in rows:
                tds = tr.find_all(["td", "th"])
                if len(tds) < 2:
                    continue
                col0 = tds[0].get_text(strip=True)
                if re.search(r"^\s*reserve\b", col0, re.I):
                    continue
                if not re.search(r"^\s*qualification\b", col0, re.I):
                    continue
                a    = tds[1].find("a")
                name = a.get_text(strip=True) if a else tds[1].get_text(strip=True)
                if name and name not in seen:
                    seen.add(name)
                    entries.append({"name": name, "entry_date": ""})

        return entries, True

    # ── Fallback: "Online entries (N)" table ────────────────────────────────
    online_table = None
    for table in soup.find_all("table", class_="ruler"):
        caption = table.find("caption")
        if caption and re.match(r"^\s*Online entries\s*\(\d+\)\s*$", caption.get_text(), re.I):
            online_table = table
            break

    if online_table:
        for tr in online_table.select("tbody tr"):
            tds = tr.find_all("td")
            if len(tds) < 2:
                continue
            name       = tds[0].get_text(strip=True)
            entry_date = tds[1].get_text(strip=True)
            if name and name.lower() not in ("name", "player"):
                entries.append({"name": name, "entry_date": entry_date})

    return entries, False


def refresh_single_tournament(sheet_name: str, col: int, url: str, draw_size: str = "") -> dict:
    """
    Refresh entries for one tournament block.
    - Clears old entry rows (50 blank rows)
    - Fetches fresh entries from LTA
    - Enriches watchlist sheets with player name if not present
    - Writes fresh entries to sheet
    Returns a log dict.
    """
    start_col = col + 1  # convert to 1-based

    # Clear old entry rows first (prevents ghost entries from previous runs)
    blank_rows = [["", ""] for _ in range(100)]
    post_to_webapp({
        "sheet":      sheet_name,
        "clearFirst": False,
        "startRow":   10,
        "startCol":   start_col,
        "rows":       blank_rows,
    })

    # Fetch fresh entries from LTA
    entries, used_entries_table = fetch_entries(url)

    # Override draw size if Entries table entry count differs from sheet value
    if used_entries_table and entries and draw_size:
        try:
            current_draw = int(draw_size)
        except (ValueError, TypeError):
            current_draw = 0
        if len(entries) != current_draw:
            post_to_webapp({
                "sheet":      sheet_name,
                "clearFirst": False,
                "startRow":   6,
                "startCol":   start_col + 1,
                "rows":       [[str(len(entries))]],
            })

    # Enrich watchlist sheets with player name if not present
    if sheet_name in WATCHLIST_PLAYERS:
        player_name    = WATCHLIST_PLAYERS[sheet_name]
        existing_names = [e["name"].strip().lower() for e in entries]
        if player_name.lower() not in existing_names:
            entries.append({"name": player_name, "entry_date": ""})

    # Write fresh entries
    rows = [[e["name"], e["entry_date"]] for e in entries] if entries else [["", ""]]
    post_to_webapp({
        "sheet":      sheet_name,
        "clearFirst": False,
        "startRow":   10,
        "startCol":   start_col,
        "rows":       rows,
    })

    return {"sheet": sheet_name, "col": col, "entries": len(entries)}


def refresh_all_entries() -> dict:
    """
    Read all 6 sheets, identify non-expired tournament blocks,
    refresh all of them in parallel, then call masterPopulate once.
    """
    today = date.today()
    log   = []
    tasks = []  # list of (sheet_name, col, url) to refresh

    # ── Read all sheets and collect tasks ────────────────────────────────────
    for sheet_name in REFRESH_SHEETS:
        try:
            result = get_from_webapp({"action": "read", "sheet": sheet_name}, timeout=30)
        except Exception as e:
            log.append({"sheet": sheet_name, "skipped": f"read failed: {e}"})
            continue

        if result.get("status") != "ok":
            log.append({"sheet": sheet_name, "skipped": f"read error: {result.get('message')}"})
            continue

        rows = result.get("rows", [])
        if not rows:
            log.append({"sheet": sheet_name, "skipped": "empty sheet"})
            continue

        for block_idx in range(10):
            label_col_idx = block_idx * 3
            value_col_idx = label_col_idx + 1

            if value_col_idx >= len(rows[0]):
                break

            tournament_name = (rows[0][label_col_idx] if len(rows[0]) > label_col_idx else "").strip()
            if not tournament_name:
                break

            date_str = (rows[1][value_col_idx] if len(rows) > 1 and len(rows[1]) > value_col_idx else "").strip()
            end_date = parse_sheet_date(date_str)

            if not end_date:
                log.append({"sheet": sheet_name, "block": block_idx, "skipped": f"unparseable date: {date_str}"})
                continue

            if end_date < today:
                log.append({"sheet": sheet_name, "block": block_idx, "skipped": f"past: {date_str}"})
                continue

            url = (rows[6][value_col_idx] if len(rows) > 6 and len(rows[6]) > value_col_idx else "").strip()
            if not url:
                log.append({"sheet": sheet_name, "block": block_idx, "skipped": "no URL"})
                continue

            tasks.append((sheet_name, block_idx * 3, url,
                          (rows[5][value_col_idx] if len(rows) > 5 and len(rows[5]) > value_col_idx else "").strip()))

    # ── Fire all tasks in parallel ───────────────────────────────────────────
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {
            executor.submit(refresh_single_tournament, sheet_name, col, url, draw_size): (sheet_name, col)
            for sheet_name, col, url, draw_size in tasks
        }
        for future in as_completed(futures):
            sheet_name, col = futures[future]
            try:
                result = future.result()
                log.append(result)
            except Exception as e:
                log.append({"sheet": sheet_name, "col": col, "error": str(e)})

    # ── Trigger masterPopulate once after all threads complete ───────────────
    try:
        post_to_webapp({"action": "masterPopulate"}, timeout=300)
        log.append({"action": "masterPopulate", "status": "ok"})
    except Exception as e:
        log.append({"action": "masterPopulate", "error": str(e)})

    return {"refreshed": len([x for x in log if "entries" in x]), "log": log}


@functions_framework.http
def refresh_entries(request):
    """HTTP Cloud Function entry point."""

    # ── CORS headers ─────────────────────────────────────────────────────────
    cors_headers = {
        "Access-Control-Allow-Origin":  "*",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type",
    }

    if request.method == "OPTIONS":
        return ("", 204, cors_headers)

    # ── Serve .ics calendar file from base64-encoded query param ─────────────
    if request.method == "GET":
        ics_b64 = request.args.get("ics", "")
        filename = request.args.get("filename", "tournament.ics")
        if not ics_b64:
            return ("Missing 'ics' parameter", 400, cors_headers)
        try:
            padded = ics_b64 + "=" * (-len(ics_b64) % 4)
            ics_content = base64.urlsafe_b64decode(padded).decode("utf-8")
        except Exception as e:
            return (f"Invalid ics encoding: {e}", 400, cors_headers)
        safe_filename = re.sub(r'[^A-Za-z0-9_.-]', '_', filename)[:100]
        return (ics_content, 200, {
            **cors_headers,
            "Content-Type": "text/calendar; charset=utf-8",
            "Content-Disposition": f'inline; filename="{safe_filename}"',
        })

    if request.method != "POST":
        return (json.dumps({"status": "error", "message": "POST required"}),
                405, {**cors_headers, "Content-Type": "application/json"})

    if not WEBAPP_URL:
        return (json.dumps({"status": "error", "message": "WEBAPP_URL not configured"}),
                500, {**cors_headers, "Content-Type": "application/json"})

    # ── Parse request ────────────────────────────────────────────────────────
    # Cloud Scheduler may send unquoted JSON e.g. {action:refreshAllEntries}
    # due to shell escaping issues. We handle both valid JSON and this format.
    try:
        body = request.get_json(force=True)
        if body is None:
            raise ValueError("null body")
    except Exception:
        try:
            raw = request.get_data(as_text=True).strip()
            # Convert unquoted {key:value} to {"key":"value"}
            raw = re.sub(r'(\{|,)\s*([A-Za-z0-9_]+)\s*:', r'\1"\2":', raw)
            raw = re.sub(r':\s*([A-Za-z0-9_]+)\s*([,}])', r':"\1"\2', raw)
            body = json.loads(raw)
        except Exception as e:
            return (json.dumps({"status": "error", "message": f"Bad request: {e}"}),
                    400, {**cors_headers, "Content-Type": "application/json"})

    # ── Refresh all entries across all 6 sheets (parallel) ───────────────────
    if body.get("action") == "refreshAllEntries":
        try:
            result = refresh_all_entries()
            return (json.dumps({"status": "ok", **result}),
                    200, {**cors_headers, "Content-Type": "application/json"})
        except Exception as e:
            return (json.dumps({"status": "error", "message": f"refreshAllEntries failed: {e}"}),
                    502, {**cors_headers, "Content-Type": "application/json"})

    # ── Mark a tournament as added to calendar ───────────────────────────────
    if body.get("action") == "markCalendarAdded":
        identifier = (body.get("identifier") or "").strip()
        if not identifier:
            return (json.dumps({"status": "error", "message": "Missing 'identifier'"}),
                    400, {**cors_headers, "Content-Type": "application/json"})
        try:
            result = post_to_webapp({"action": "markCalendarAdded", "identifier": identifier})
            return (json.dumps(result),
                    200, {**cors_headers, "Content-Type": "application/json"})
        except Exception as e:
            return (json.dumps({"status": "error", "message": f"markCalendarAdded failed: {e}"}),
                    502, {**cors_headers, "Content-Type": "application/json"})

    # ── Single tournament refresh ────────────────────────────────────────────
    try:
        tournament_url = body["url"]
        sheet_name     = body["sheet"]
        col            = int(body["col"])   # 0-based block index (0, 3, 6, ...)
    except Exception as e:
        return (json.dumps({"status": "error", "message": f"Bad request: {e}"}),
                400, {**cors_headers, "Content-Type": "application/json"})

    try:
        result = refresh_single_tournament(sheet_name, col, tournament_url)
    except Exception as e:
        return (json.dumps({"status": "error", "message": f"Refresh failed: {e}"}),
                502, {**cors_headers, "Content-Type": "application/json"})

    # Trigger masterPopulate
    try:
        post_to_webapp({"action": "masterPopulate"})
    except Exception as e:
        return (json.dumps({
            "status":  "partial",
            "message": f"Entries written but masterPopulate failed: {e}",
            "entries": result["entries"],
        }), 200, {**cors_headers, "Content-Type": "application/json"})

    return (json.dumps({
        "status":  "ok",
        "entries": result["entries"],
        "sheet":   sheet_name,
        "col":     col,
    }), 200, {**cors_headers, "Content-Type": "application/json"})