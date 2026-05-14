"""
LTA Watchlist Scraper
Scrapes tournament data for watchlist tournaments (not yet entered).

Reads tournament URLs and event types from the WATCHLIST Google Sheet:
  Luka:  col B = Event, col C = URL  (player name in row 4, data from row 6)
  Dylan: col F = Event, col G = URL
  Serge: col J = Event, col K = URL

For each tournament:
  - Visits the tournament page (no login required)
  - Dismisses cookie banner if present
  - Scrapes Fact Sheet: name, date, closing deadline, grade, draw size
  - Scrapes entry list from Events tab
  - Checks if player (e.g. "Luka Taborin") is already in entries
  - If not, appends player name to bottom of entry list
  - Writes results to LUKA_WATCHLIST / DYLAN_WATCHLIST / SERGE_WATCHLIST
    in the same 3-col stride format as LUKA_TOURNAMENTS etc.

Output row layout (mirrors existing tournament sheets):
  Row 1: "Tournament: <name>"
  Row 2: "Date:"         | date string
  Row 3: "Event:"        | event type
  Row 4: "Closing Date:" | closing deadline string (replaces Status)
  Row 5: "Grade:"        | grade
  Row 6: "Draw Size:"    | draw size
  Row 7: blank
  Row 8: "Entry Name"    | (no entry date needed)
  Row 9+: player names
"""

import os
import re
import logging
import requests
from datetime import datetime

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

WEBAPP_URL    = os.environ.get("LTA_SHEETS_WEBAPP_URL", "").strip()
WEBAPP_SECRET = os.environ.get("LTA_SHEETS_WEBAPP_SECRET", "").strip()

# ── Player config ─────────────────────────────────────────────────────────────
PLAYERS = [
    {
        "name":        "Luka Taborin",
        "key":         "luka",
        "event_col":   2,   # col B (1-based)
        "url_col":     3,   # col C
        "name_col":    4,   # col D
        "entered_sheet": "LUKA_TOURNAMENTS",
        "output_sheet": "LUKA_WATCHLIST",
    },
    {
        "name":        "Dylan Taborin",
        "key":         "dylan",
        "event_col":   6,   # col F
        "url_col":     7,   # col G
        "name_col":    8,   # col H
        "entered_sheet": "DYLAN_TOURNAMENTS",
        "output_sheet": "DYLAN_WATCHLIST",
    },
    {
        "name":        "Serge Taborin",
        "key":         "serge",
        "event_col":   10,  # col J
        "url_col":     11,  # col K
        "name_col":    12,  # col L
        "entered_sheet": "SERGE_TOURNAMENTS",
        "output_sheet": "SERGE_WATCHLIST",
    },
]

WATCHLIST_SHEET = "WATCHLIST"
DATA_START_ROW  = 6   # row 6 onwards (1-based)
DATA_END_ROW    = 15  # row 15 (1-based) — max 10 watchlist slots


# ── Logger ────────────────────────────────────────────────────────────────────
def setup_logger() -> logging.Logger:
    logger = logging.getLogger("lta_watchlist")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
        ch.setFormatter(fmt)
        logger.addHandler(ch)
    return logger


# ── Read watchlist from Google Sheets ────────────────────────────────────────
def read_watchlist_from_sheet(player: dict, logger: logging.Logger) -> list[dict]:
    """
    Reads event/URL pairs for one player from the WATCHLIST sheet via webapp.
    Returns list of dicts: {event, url}
    """
    if not WEBAPP_URL:
        raise RuntimeError("Missing env var LTA_SHEETS_WEBAPP_URL")

    url = WEBAPP_URL
    if WEBAPP_SECRET:
        joiner = "&" if "?" in url else "?"
        url = f"{url}{joiner}secret={WEBAPP_SECRET}"

    # Use GET with sheet + range params to read data
    params = {
        "action": "read",
        "sheet":  WATCHLIST_SHEET,
    }

    try:
        r = requests.get(url, params=params, timeout=30)
        r.raise_for_status()
        response = r.json()
        if response.get("status") != "ok":
            logger.error(f"Webapp error: {response.get('message')}")
            return []
        data = response["rows"]  # extract the actual 2D array
    except Exception as e:
        logger.error(f"Failed to read WATCHLIST sheet: {e}")
        return []

    entries = []
    event_col = player["event_col"] - 1   # convert to 0-based
    url_col   = player["url_col"] - 1

    
    for row_idx, row in enumerate(data):
        
        # Skip rows before data start (rows are 0-based here, DATA_START_ROW is 1-based)
        if row_idx < DATA_START_ROW - 1:
            continue

        # Safely get values
        event = str(row[event_col]).strip() if len(row) > event_col else ""
        link  = str(row[url_col]).strip()   if len(row) > url_col   else ""

        if not event or not link:
            continue

        entries.append({"event": event, "url": link, "sheet_row": row_idx + 1})  # 1-based row

    logger.info(f"[{player['key']}] Found {len(entries)} watchlist entries")
    return entries


# ── Deduplication: remove entered tournaments from watchlist ──────────────────

def extract_uuid(url_str: str) -> str | None:
    """Extract a UUID from any LTA URL format. Returns lowercase or None."""
    if not url_str:
        return None
    m = re.search(r"[A-Fa-f0-9]{8}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{12}", url_str)
    return m.group(0).lower() if m else None


def read_entered_tournament_ids(player: dict, logger: logging.Logger) -> set[str]:
    """
    Read one player's entered tournament sheet and return a set of tournament UUIDs.
    URLs live in row 7 (0-based index 6), at col+1 of each 3-column block.
    """
    if not WEBAPP_URL:
        return set()

    url = WEBAPP_URL
    if WEBAPP_SECRET:
        joiner = "&" if "?" in url else "?"
        url = f"{url}{joiner}secret={WEBAPP_SECRET}"

    sheet_name = player["entered_sheet"]
    try:
        r = requests.get(url, params={"action": "read", "sheet": sheet_name}, timeout=30)
        r.raise_for_status()
        response = r.json()
        if response.get("status") != "ok":
            logger.error(f"Error reading {sheet_name}: {response.get('message')}")
            return set()
        data = response["rows"]
    except Exception as e:
        logger.error(f"Failed to read {sheet_name}: {e}")
        return set()

    ids = set()
    if len(data) < 7:
        return ids

    url_row = data[6]  # row 7 (0-based index 6)
    # URL values sit at col 1, 4, 7, 10, ... (col+1 of each 3-col block)
    for col_idx in range(1, len(url_row), 3):
        uid = extract_uuid(str(url_row[col_idx]))
        if uid:
            ids.add(uid)

    logger.info(f"[{player['key']}] Found {len(ids)} entered tournament ID(s) in {sheet_name}")
    return ids


def remove_entered_from_watchlist(logger: logging.Logger) -> int:
    """
    Read WATCHLIST rows 6-15. For each player, if a watchlist URL's UUID matches
    an entered tournament, blank the event/url/name cells. Write rows 6-15 back.
    Returns total number of entries removed.
    """
    if not WEBAPP_URL:
        return 0

    url = WEBAPP_URL
    if WEBAPP_SECRET:
        joiner = "&" if "?" in url else "?"
        url = f"{url}{joiner}secret={WEBAPP_SECRET}"

    # Read the full WATCHLIST sheet
    try:
        r = requests.get(url, params={"action": "read", "sheet": WATCHLIST_SHEET}, timeout=30)
        r.raise_for_status()
        response = r.json()
        if response.get("status") != "ok":
            logger.error(f"Error reading WATCHLIST: {response.get('message')}")
            return 0
        all_data = response["rows"]
    except Exception as e:
        logger.error(f"Failed to read WATCHLIST: {e}")
        return 0

    # Extract rows 6-15 (0-based indices 5-14)
    start_idx = DATA_START_ROW - 1   # 5
    end_idx   = DATA_END_ROW         # 15 (exclusive in slice = rows 6-15)
    if len(all_data) < end_idx:
        # Pad with empty rows if sheet is shorter than 15 rows
        max_cols = max((len(row) for row in all_data), default=12)
        while len(all_data) < end_idx:
            all_data.append([""] * max_cols)

    data_rows = [list(row) for row in all_data[start_idx:end_idx]]

    # Ensure each row has enough columns (at least 12 for col L)
    min_cols = 12
    for row in data_rows:
        while len(row) < min_cols:
            row.append("")

    total_removed = 0

    for player in PLAYERS:
        entered_ids = read_entered_tournament_ids(player, logger)
        if not entered_ids:
            continue

        event_ci = player["event_col"] - 1  # 0-based
        url_ci   = player["url_col"] - 1
        name_ci  = player["name_col"] - 1

        for row in data_rows:
            wl_uuid = extract_uuid(str(row[url_ci]))
            if wl_uuid and wl_uuid in entered_ids:
                tournament_name = str(row[name_ci]).strip() or str(row[url_ci]).strip()
                logger.info(f"[{player['key']}] Removing entered tournament from watchlist: {tournament_name}")
                row[event_ci] = ""
                row[url_ci]   = ""
                row[name_ci]  = ""
                total_removed += 1

    if total_removed > 0:
        # Write only rows 6-15 back (clearFirst=false, startRow=6)
        payload = {
            "sheet":      WATCHLIST_SHEET,
            "clearFirst": False,
            "startRow":   DATA_START_ROW,
            "rows":       data_rows,
        }
        try:
            r = requests.post(url, json=payload, timeout=30)
            r.raise_for_status()
            logger.info(f"WATCHLIST updated: {total_removed} entered tournament(s) removed")
        except Exception as e:
            logger.error(f"Failed to write updated WATCHLIST: {e}")
    else:
        logger.info("No entered tournaments found on watchlist — nothing to remove")

    return total_removed


# ── Cookie banner ─────────────────────────────────────────────────────────────
def dismiss_cookie_banner(page):
    try:
        btn = page.query_selector('button:has-text("ACCEPT"), button:has-text("Accept")')
        if btn:
            btn.click()
            page.wait_for_timeout(800)
            print("  ✓ Cookie banner dismissed")
    except Exception as e:
        print(f"  (Cookie banner: {e})")


# ── Fact sheet extraction (reused from existing scrapers) ─────────────────────
def extract_fact_sheet_value(page, event_text, label_text):
    event_text = (event_text or "").strip()
    label_text = (label_text or "").strip()

    # A) Accordion path
    try:
        toggle = page.locator("dt.list__label button").filter(
            has=page.locator("span.text--bold",
                has_text=re.compile(rf"^\s*{re.escape(event_text)}\s*$", re.I))
        ).first

        if toggle.count() > 0:
            if toggle.get_attribute("aria-expanded") != "true":
                toggle.click()
                page.wait_for_timeout(500)

            dt_node = toggle.locator("xpath=ancestor::dt[1]").first
            panel   = dt_node.locator("xpath=following-sibling::dd[1]").first
            dt_label = panel.locator(
                "dt.list__label",
                has_text=re.compile(rf"^\s*{re.escape(label_text)}\s*$", re.I)
            ).first
            if dt_label.count() > 0:
                dd_val = dt_label.locator("xpath=following-sibling::dd[1]").first
                val = dd_val.inner_text().strip()
                return val or None
    except Exception:
        pass

    # B) Fallback path
    try:
        container = page.locator("dl.list--flex").first
        if container.count() == 0:
            return None
        dt_label = container.locator(
            "dt.list__label",
            has_text=re.compile(rf"^\s*{re.escape(label_text)}\s*$", re.I)
        ).first
        if dt_label.count() == 0:
            return None
        dd_val = dt_label.locator("xpath=following-sibling::dd[1]").first
        val = dd_val.inner_text().strip()
        return val or None
    except Exception:
        return None


def parse_closing_deadline_str(closing_str: str) -> str:
    """
    Returns a clean closing deadline string "dd/mm/yyyy HH:MM" or "" if unparseable.
    """
    if not closing_str:
        return ""
    s = closing_str.replace("\u00a0", " ").strip()
    m = re.search(r"\b(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2})(?::\d{2})?\b", s)
    if not m:
        return closing_str  # return raw if we can't parse cleanly
    return f"{m.group(1)} {m.group(2)}"


def extract_start_time(timings_str: str | None) -> str | None:
    """
    Extract a start time from a Timings string if present.
    Matches patterns like: "1.45pm start", "10am", "2:30pm", "10.00am start"
    Returns the time string (e.g. "1.45pm") or None.
    """
    if not timings_str:
        return None
    m = re.search(r'\b(\d{1,2}(?:[.:]\d{2})?\s*[ap]m)\b', timings_str, re.IGNORECASE)
    if m:
        return m.group(1).strip().lower()
    return None


def parse_closing_datetime(closing_str: str) -> datetime | None:
    """
    Parse a closing deadline string into a datetime for comparison.
    Accepts "dd/mm/yyyy HH:MM" or raw strings with embedded date/time.
    Returns datetime or None if unparseable.
    """
    if not closing_str:
        return None
    s = closing_str.replace("\u00a0", " ").strip()
    m = re.search(r"\b(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2})", s)
    if not m:
        m2 = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", s)
        if m2:
            try:
                return datetime.strptime(m2.group(1), "%d/%m/%Y")
            except ValueError:
                return None
        return None
    try:
        return datetime.strptime(f"{m.group(1)} {m.group(2)}", "%d/%m/%Y %H:%M")
    except ValueError:
        return None


def clear_expired_watchlist_rows(player: dict, expired_rows: list[int], logger: logging.Logger):
    """
    Clear event + URL + name cells on the WATCHLIST sheet for expired rows.
    expired_rows: list of 1-based row numbers.
    """
    if not expired_rows or not WEBAPP_URL:
        return

    url = WEBAPP_URL
    if WEBAPP_SECRET:
        joiner = "&" if "?" in url else "?"
        url = f"{url}{joiner}secret={WEBAPP_SECRET}"

    for row in expired_rows:
        # Clear event + URL columns
        payload = {
            "sheet":      WATCHLIST_SHEET,
            "clearFirst": False,
            "startRow":   row,
            "startCol":   player["event_col"],
            "rows":       [["", ""]],
        }
        try:
            r = requests.post(url, json=payload, timeout=30)
            r.raise_for_status()
        except Exception as e:
            logger.error(f"[{player['key']}] Failed to clear event/url row {row}: {e}")

        # Clear name column
        payload_name = {
            "sheet":      WATCHLIST_SHEET,
            "clearFirst": False,
            "startRow":   row,
            "startCol":   player["name_col"],
            "rows":       [[""]],
        }
        try:
            r = requests.post(url, json=payload_name, timeout=30)
            r.raise_for_status()
        except Exception as e:
            logger.error(f"[{player['key']}] Failed to clear name row {row}: {e}")

        logger.info(f"[{player['key']}] Cleared expired watchlist row {row}")


def write_tournament_names(player: dict, name_entries: list[dict], logger: logging.Logger):
    """
    Write tournament names back to the WATCHLIST sheet.
    name_entries: list of {"sheet_row": int, "name": str}
    """
    if not name_entries or not WEBAPP_URL:
        return

    url = WEBAPP_URL
    if WEBAPP_SECRET:
        joiner = "&" if "?" in url else "?"
        url = f"{url}{joiner}secret={WEBAPP_SECRET}"

    for entry in name_entries:
        payload = {
            "sheet":      WATCHLIST_SHEET,
            "clearFirst": False,
            "startRow":   entry["sheet_row"],
            "startCol":   player["name_col"],
            "rows":       [[entry["name"]]],
        }
        try:
            r = requests.post(url, json=payload, timeout=30)
            r.raise_for_status()
        except Exception as e:
            logger.error(f"[{player['key']}] Failed to write name at row {entry['sheet_row']}: {e}")

    logger.info(f"[{player['key']}] Wrote {len(name_entries)} tournament name(s) to WATCHLIST sheet")


# ── Scrape one watchlist tournament ──────────────────────────────────────────
def scrape_watchlist_tournament(page, url: str, event: str, player_name: str) -> dict | None:
    """
    Visits tournament page, scrapes fact sheet + entries for the given event.
    Appends player_name to entries if not already present.
    Returns tournament dict or None on failure.
    """
    print(f"\n{'='*60}")
    print(f"Scraping watchlist tournament: {url}")
    print(f"Event: {event} | Player: {player_name}")
    print(f"{'='*60}")

    tournament = {
        "name":         "",
        "date":         "",
        "start_time":   "",
        "event":        event,
        "closing_date": "",
        "grade":        "",
        "draw_size":    0,
        "entries":      [],
        "url":          url,
    }

    try:
        page.goto(url, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(1000)
        dismiss_cookie_banner(page)

        # ── Entry open date (from Overview page timeline, before navigating away)
        try:
            entry_open_el = page.locator("li.list_item.is-entry-open time").first
            if entry_open_el.count() > 0:
                iso_str = entry_open_el.get_attribute("datetime") or ""
                m_iso = re.search(r"(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})", iso_str)
                if m_iso:
                    tournament["entry_open_date"] = f"{m_iso.group(3)}/{m_iso.group(2)}/{m_iso.group(1)} {m_iso.group(4)}:{m_iso.group(5)}"
                    print(f"  ✓ Entry open date: {tournament['entry_open_date']}")
        except Exception as e:
            print(f"  (Entry open date not found: {e})")
        # ── Fact Sheet ────────────────────────────────────────────────────
        try:
            fact_tab = page.get_by_role("link", name=re.compile(r"^\s*Fact Sheet\s*$", re.I))
            if fact_tab.count() == 0:
                fact_tab = page.locator('a:has-text("Fact Sheet")')
            fact_tab.first.click(timeout=15000)
            page.wait_for_load_state("domcontentloaded")
            page.wait_for_timeout(800)
        except Exception as e:
            print(f"  ✗ Could not navigate to Fact Sheet: {e}")
            return None

        # ── Find and click event button ────────────────────────────────────
        page.wait_for_selector("text=/Event Information/i", timeout=10000)
        page.wait_for_timeout(600)


        try:
            raw_name = extract_fact_sheet_value(page, "", "Tournament name")
            if raw_name:
                tournament["name"] = raw_name.replace("London & South East Tour - ", "").strip()
                print(f"  ✓ Name: {tournament['name']}")
            else:
                print("  ✗ Tournament name not found")
        except Exception as e:
            print(f"  ✗ Could not extract name: {e}")


        all_buttons = page.query_selector_all("button")
        event_button = None
        for btn in all_buttons:
            try:
                if btn.inner_text().strip() == event:
                    event_button = btn
                    break
            except Exception:
                continue

        if not event_button:
            print(f"  ✗ Could not find event button for '{event}'")
            return None

        event_button.click()
        page.wait_for_timeout(1500)

        # ── Start date ────────────────────────────────────────────────────
        try:
            raw_start = extract_fact_sheet_value(page, event, "Start date")
            if raw_start:
                m = re.search(r"\b\d{2}/\d{2}/\d{4}\b", raw_start)
                tournament["date"] = m.group(0) if m else raw_start
                print(f"  ✓ Start date: {tournament['date']}")
        except Exception as e:
            print(f"  ✗ Start date error: {e}")

        # ── End date (append to date string if available) ─────────────────
        try:
            raw_end = extract_fact_sheet_value(page, event, "End date")
            if raw_end:
                m = re.search(r"\b\d{2}/\d{2}/\d{4}\b", raw_end)
                end_str = m.group(0) if m else raw_end
                if end_str and end_str != tournament["date"]:
                    tournament["date"] = f"{tournament['date']} to {end_str}"
                    print(f"  ✓ Date range: {tournament['date']}")
        except Exception:
            pass

        # ── Start time from Timings (stored separately) ──────────────────
        try:
            timings_raw = extract_fact_sheet_value(page, event, "Timings")
            if timings_raw:
                start_time = extract_start_time(timings_raw)
                if start_time:
                    tournament["start_time"] = start_time
                    print(f"  ✓ Start time: {start_time}")
        except Exception:
            pass

        # ── Closing deadline ───────────────────────────────────────────────
        try:
            raw_closing = extract_fact_sheet_value(page, event, "Closing deadline")
            tournament["closing_date"] = parse_closing_deadline_str(raw_closing or "")
            print(f"  ✓ Closing date: {tournament['closing_date']}")
        except Exception as e:
            print(f"  ✗ Closing date error: {e}")

        # ── Grade ──────────────────────────────────────────────────────────
        try:
            raw_grade = extract_fact_sheet_value(page, event, "Grade")
            if raw_grade:
                m = re.search(r"Grade\s*\d+", raw_grade, re.I)
                tournament["grade"] = m.group(0) if m else raw_grade
                print(f"  ✓ Grade: {tournament['grade']}")
        except Exception as e:
            print(f"  ✗ Grade error: {e}")

        # ── Draw size ──────────────────────────────────────────────────────
        try:
            raw_draw = extract_fact_sheet_value(page, event, "Proposed Draw Size")
            if raw_draw:
                m = re.search(r"\d+", raw_draw)
                tournament["draw_size"] = int(m.group(0)) if m else 16
            else:
                tournament["draw_size"] = 16
            print(f"  ✓ Draw size: {tournament['draw_size']}")
        except Exception as e:
            print(f"  ✗ Draw size error: {e}")
            tournament["draw_size"] = 16

    except Exception as e:
        print(f"  ERROR scraping fact sheet: {e}")
        return None

    # ── Entries ────────────────────────────────────────────────────────────
    scrape_watchlist_entries(page, tournament)

    # ── Add player if not already in list ─────────────────────────────────
    existing_names = [e["name"].strip().lower() for e in tournament["entries"]]
    if player_name.lower() not in existing_names:
        tournament["entries"].append({"name": player_name})
        print(f"  ✓ Added '{player_name}' to entry list (was not present)")
    else:
        print(f"  ℹ  '{player_name}' already in entry list")

    return tournament


def scrape_watchlist_entries(page, tournament: dict):
    """
    Scrapes the entry list for the tournament's event.
    Entry date is not needed — names only.
    Reuses same logic as existing scrapers.
    """
    event_text = (tournament.get("event") or "").strip()
    print(f"\n  Scraping entries for event: {event_text}")

    try:
        # Navigate to Events tab
        events_tab = page.locator('a', has_text=re.compile(r'^\s*Events\s*$', re.I)).first
        if events_tab.count() == 0:
            events_tab = page.locator('a[href*="/events"]').first
        if events_tab.count() == 0:
            print("  ✗ Could not find Events tab")
            tournament["entries"] = []
            return

        events_tab.click()
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(800)
        # Wait for events table to render (up to 5s)
        try:
            page.locator("table.ruler.admintournamentevents").first.wait_for(timeout=5000)
        except:
            pass  # proceed anyway — table might have a different class on some pages

        # Click event link
        event_link = page.locator(
            "a", has_text=re.compile(rf"^\s*{re.escape(event_text)}\s*$", re.I)
        ).first
        if event_link.count() == 0:
            event_link = page.locator(f'a:has-text("{event_text}")').first
        if event_link.count() == 0:
            print(f"  ✗ Could not find event link for '{event_text}'")
            tournament["entries"] = []
            return

        event_link.click()
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(1200)
        tournament["url"] = page.url  # replace overview URL with direct entries URL

        entries = []
        seen    = set()

        # ── Prefer "Entries (N)" table ─────────────────────────────────────
        entries_table = page.locator("table.ruler").filter(
            has=page.locator("caption",
                has_text=re.compile(r"^\s*Entries\s*\(\d+\)\s*$", re.I))
        ).first

        if entries_table.count() > 0:
            print("  ✓ Found 'Entries' table")
            rows = entries_table.locator("tbody tr")
            n    = rows.count()

            for i in range(n):
                tr  = rows.nth(i)
                tds = tr.locator("td")
                if tds.count() < 2:
                    continue

                raw_label = tds.nth(0).inner_text() or ""
                label = raw_label.replace("\u00a0", " ").strip().lower()

                if label.startswith("reserve"):
                    continue
                if not (label.startswith("maindraw") or label.startswith("qualification")):
                    continue

                player_td   = tds.nth(1)
                player_link = player_td.locator("a").first
                player_name = (
                    player_link.inner_text().strip()
                    if player_link.count() > 0
                    else player_td.inner_text().strip()
                )

                if player_name and player_name not in seen:
                    seen.add(player_name)
                    entries.append({"name": player_name})

            # Pass 2: qualification rows only if pass 1 returned nothing
            if not entries:
                for i in range(n):
                    tr  = rows.nth(i)
                    tds = tr.locator("td")
                    if tds.count() < 2:
                        continue
                    col0 = tds.nth(0).inner_text().strip()
                    if re.search(r"^\s*reserve\b", col0, re.I):
                        continue
                    if not re.search(r"^\s*qualification\b", col0, re.I):
                        continue
                    player_td   = tds.nth(1)
                    player_link = player_td.locator("a").first
                    player_name = (
                        player_link.inner_text().strip()
                        if player_link.count() > 0
                        else player_td.inner_text().strip()
                    )
                    if player_name and player_name not in seen:
                        seen.add(player_name)
                        entries.append({"name": player_name})

            tournament["entries"] = entries
            print(f"  ✓ {len(entries)} entries extracted")
            return

        # ── Fallback: "Online entries (N)" table ───────────────────────────
        online_table = page.locator("table.ruler").filter(
            has=page.locator("caption",
                has_text=re.compile(r"^\s*Online entries\s*\(\d+\)\s*$", re.I))
        ).first

        if online_table.count() == 0:
            print("  ✗ No entries table found")
            tournament["entries"] = []
            return

        print("  ✓ Found 'Online entries' table")
        rows = online_table.locator("tbody tr")
        n    = rows.count()

        for i in range(n):
            tr  = rows.nth(i)
            tds = tr.locator("td")
            if tds.count() < 1:
                continue
            name = tds.nth(0).inner_text().strip()
            if name and name.lower() not in ("name", "player"):
                entries.append({"name": name})

        tournament["entries"] = entries
        print(f"  ✓ {len(entries)} entries extracted")

    except Exception as e:
        print(f"  ✗ ERROR scraping entries: {e}")
        tournament["entries"] = []


# ── Write results to Google Sheets ───────────────────────────────────────────
def write_watchlist_to_sheet(tournaments: list, sheet_tab: str, logger: logging.Logger):
    """
    Writes watchlist tournaments to Google Sheets.
    Same 3-col stride format as existing tournament sheets.
    Row layout:
      1: Tournament name
      2: Date
      3: Event
      4: Closing Date   ← replaces Status
      5: Grade
      6: Draw Size
      7: blank
      8: Entry Name header
      9+: entries (name only)
    """
    if not WEBAPP_URL:
        raise RuntimeError("Missing env var LTA_SHEETS_WEBAPP_URL")

    if not tournaments:
        logger.info(f"No watchlist tournaments to write for {sheet_tab} — clearing sheet")
        # Still send a clear so stale data from previous runs is removed
        url = WEBAPP_URL
        if WEBAPP_SECRET:
            joiner = "&" if "?" in url else "?"
            url = f"{url}{joiner}secret={WEBAPP_SECRET}"
        payload = {"sheet": sheet_tab, "clearFirst": True, "rows": [[""]]}
        requests.post(url, json=payload, timeout=30)
        return

    num_tournaments = len(tournaments)
    total_cols      = num_tournaments * 3
    max_entries     = max((len(t.get("entries") or []) for t in tournaments), default=0)
    total_rows      = 9 + max_entries  # row 7 now holds URL

    grid = [[None] * total_cols for _ in range(total_rows)]

    for t_idx, t in enumerate(tournaments):
        col = t_idx * 3

        grid[0][col]     = f"Tournament: {t.get('name', '')}"
        grid[1][col]     = "Date:"
        grid[1][col + 1] = t.get("date", "")
        grid[1][col + 2] = t.get("start_time", "")
        grid[2][col]     = "Event:"
        grid[2][col + 1] = t.get("event", "")
        grid[3][col]     = "Closing Date:"
        grid[3][col + 1] = t.get("closing_date", "")
        grid[3][col + 2] = t.get("entry_open_date", "")
        grid[4][col]     = "Grade:"
        grid[4][col + 1] = t.get("grade", "")
        grid[5][col]     = "Draw Size:"
        grid[5][col + 1] = t.get("draw_size", "")
        grid[6][col]     = "URL:"
        grid[6][col + 1] = t.get("url", "")        # row 7: URL (matches tournament scrapers)
        grid[8][col]     = "Entry Name"             # row 9: header (was row 8)

        for e_idx, entry in enumerate(t.get("entries") or []):
            grid[9 + e_idx][col] = entry.get("name", "")  # entries start row 10 (was 9)

    # Replace None with ""
    grid = [["" if cell is None else cell for cell in row] for row in grid]

    url = WEBAPP_URL
    if WEBAPP_SECRET:
        joiner = "&" if "?" in url else "?"
        url = f"{url}{joiner}secret={WEBAPP_SECRET}"

    payload = {
        "sheet":      sheet_tab,
        "clearFirst": True,
        "rows":       grid,
    }

    logger.info(f"Writing {len(tournaments)} watchlist tournament(s) to '{sheet_tab}'")
    r = requests.post(url, json=payload, timeout=120)
    r.raise_for_status()
    logger.info(f"Response: {r.text.strip()}")

    # ── Timestamp ──────────────────────────────────────────────────────────
    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    ts_payload = {
        "sheet":      "WATCHLIST_DASHBOARD",
        "clearFirst": False,
        "startRow":   100,
        "rows":       [[timestamp]],
    }
    ts_r = requests.post(url, json=ts_payload, timeout=30)
    ts_r.raise_for_status()
    logger.info(f"Timestamp written to WATCHLIST_DASHBOARD A100")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("=== LTA Watchlist Scraper ===\n")
    logger = setup_logger()

    # Remove any watchlist entries that the player has already entered
    remove_entered_from_watchlist(logger)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page    = context.new_page()

        for player in PLAYERS:
            print(f"\n{'#'*60}")
            print(f"# Processing player: {player['name']}")
            print(f"{'#'*60}")

            # Read watchlist entries for this player
            watchlist = read_watchlist_from_sheet(player, logger)

            if not watchlist:
                print(f"  No watchlist entries for {player['name']} — clearing sheet and skipping")
                write_watchlist_to_sheet(
                    tournaments=[],
                    sheet_tab=player["output_sheet"],
                    logger=logger,
                )
                continue

            tournaments = []
            expired_rows = []
            name_entries = []
            now = datetime.now()

            for item in watchlist:
                result = scrape_watchlist_tournament(
                    page,
                    url=item["url"],
                    event=item["event"],
                    player_name=player["name"],
                )
                if result:
                    # Track name for writing back to WATCHLIST sheet
                    if result.get("name"):
                        name_entries.append({"sheet_row": item["sheet_row"], "name": result["name"]})

                    # Check if closing date has passed
                    closing_dt = parse_closing_datetime(result.get("closing_date", ""))
                    if closing_dt and closing_dt < now:
                        print(f"  ✗ Expired (closing date passed): {result.get('name', '')} [{result.get('closing_date', '')}]")
                        expired_rows.append(item["sheet_row"])
                    else:
                        tournaments.append(result)
                else:
                    print(f"  ✗ Skipping failed tournament: {item['url']}")

            # Write tournament names to WATCHLIST source sheet (active ones only)
            active_names = [n for n in name_entries if n["sheet_row"] not in expired_rows]
            if active_names:
                write_tournament_names(player, active_names, logger)

            # Clear expired rows from WATCHLIST source sheet
            if expired_rows:
                print(f"\n  Removing {len(expired_rows)} expired tournament(s) from WATCHLIST sheet")
                clear_expired_watchlist_rows(player, expired_rows, logger)

            print(f"\n  {len(tournaments)} active tournament(s) for {player['name']}"
                  f" ({len(expired_rows)} expired, removed from WATCHLIST)")

            write_watchlist_to_sheet(
                tournaments=tournaments,
                sheet_tab=player["output_sheet"],
                logger=logger,
            )

        browser.close()

    print("\n✅ Watchlist scraping complete!")


if __name__ == "__main__":
    main()