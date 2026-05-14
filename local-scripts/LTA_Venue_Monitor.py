"""
LTA Venue Monitor
─────────────────
Monitors LTA tournament search results for specific players and flags
new tournaments or status changes (e.g. cancellations).

Flow:
  1. For each player/search, fetch all search result pages (dates auto-calculated)
  2. Parse tournament cards, filter to allowed age groups only
  3. Read existing data from "Tournament_Venues" sheet
  4. Compare: flag new tournaments and newly cancelled ones with "Y"
  5. Write merged list back to sheet

Designed to run locally or as a Cloud Function on a weekly schedule.
"""

import os
import re
import requests
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

WEBAPP_URL    = os.environ.get("LTA_SHEETS_WEBAPP_URL", "").strip()
WEBAPP_SECRET = os.environ.get("LTA_SHEETS_WEBAPP_SECRET", "").strip()

SHEET_NAME = "Tournament_Venues"

# ── Search config ─────────────────────────────────────────────────────────────
# Each search has a player name, allowed age groups (only these will be kept),
# a description (for logging), and a base search URL.
# StartDate and EndDate in URLs are replaced dynamically with today and today+90.
SEARCHES = [
    # ── Serge ──
    {
        "player": "Serge",
        "allowed_age_groups": {"Open", "35+", "40+", "45+", "50+"},
        "description": "NTC, MS, Open/35+/40+/45+/50+",
        "url": "https://competitions.lta.org.uk/find?DateFilterType=0&StartDate=2026-04-25&EndDate=2026-07-29&LocationFilterType=0&PostalCode=SW15+5FG&StatusFilterID=4&EventGameTypeIDList%5B0%5D=1&EventGameTypeIDList%5B1%5D=false&EventGameTypeIDList%5B2%5D=false&EventGameTypeIDList%5B3%5D=false&EventGameTypeIDList%5B4%5D=false&EventGameTypeIDList%5B5%5D=false&EventGameTypeIDList%5B6%5D=false&AgeGroupIDList%5B0%5D=false&AgeGroupIDList%5B1%5D=false&AgeGroupIDList%5B2%5D=false&AgeGroupIDList%5B3%5D=false&AgeGroupIDList%5B4%5D=false&AgeGroupIDList%5B5%5D=false&AgeGroupIDList%5B6%5D=false&AgeGroupIDList%5B7%5D=false&AgeGroupIDList%5B8%5D=99&AgeGroupIDList%5B9%5D=false&AgeGroupIDList%5B10%5D=35&AgeGroupIDList%5B11%5D=40&AgeGroupIDList%5B12%5D=45&AgeGroupIDList%5B13%5D=50&AgeGroupIDList%5B14%5D=false&AgeGroupIDList%5B15%5D=false&AgeGroupIDList%5B16%5D=false&AgeGroupIDList%5B17%5D=false&AgeGroupIDList%5B18%5D=false&AgeGroupIDList%5B19%5D=false&AgeGroupIDList%5B20%5D=false&AgeGroupIDList%5B21%5D=false&page=1&Q=National+Tennis+Centre",
    },
    {
        "player": "Serge",
        "allowed_age_groups": {"35+", "40+", "45+", "50+"},
        "description": "10mi, MS, 35+/40+/45+/50+",
        "url": "https://competitions.lta.org.uk/find?DateFilterType=0&StartDate=2026-04-25&EndDate=2026-07-29&LocationFilterType=0&PostalCode=SW15+5FG&StatusFilterID=4&page=1&EventGameTypeIDList%5B0%5D=1&EventGameTypeIDList%5B1%5D=false&EventGameTypeIDList%5B2%5D=false&EventGameTypeIDList%5B3%5D=false&EventGameTypeIDList%5B4%5D=false&EventGameTypeIDList%5B5%5D=false&EventGameTypeIDList%5B6%5D=false&Distance=10&AgeGroupIDList%5B0%5D=false&AgeGroupIDList%5B1%5D=false&AgeGroupIDList%5B2%5D=false&AgeGroupIDList%5B3%5D=false&AgeGroupIDList%5B4%5D=false&AgeGroupIDList%5B5%5D=false&AgeGroupIDList%5B6%5D=false&AgeGroupIDList%5B7%5D=false&AgeGroupIDList%5B8%5D=false&AgeGroupIDList%5B9%5D=false&AgeGroupIDList%5B10%5D=35&AgeGroupIDList%5B11%5D=40&AgeGroupIDList%5B12%5D=45&AgeGroupIDList%5B13%5D=50&AgeGroupIDList%5B14%5D=false&AgeGroupIDList%5B15%5D=false&AgeGroupIDList%5B16%5D=false&AgeGroupIDList%5B17%5D=false&AgeGroupIDList%5B18%5D=false&AgeGroupIDList%5B19%5D=false&AgeGroupIDList%5B20%5D=false&AgeGroupIDList%5B21%5D=false",
    },
    # ── Luka ──
    {
        "player": "Luka",
        "allowed_age_groups": {"14U"},
        "description": "20mi, BS, 14U, G4+G5",
        "url": "https://competitions.lta.org.uk/find?DateFilterType=0&StartDate=2026-04-25&EndDate=2026-07-29&LocationFilterType=0&PostalCode=SW15+5FG&StatusFilterID=4&page=1&EventGameTypeIDList%5B0%5D=1&EventGameTypeIDList%5B1%5D=false&EventGameTypeIDList%5B2%5D=false&EventGameTypeIDList%5B3%5D=false&EventGameTypeIDList%5B4%5D=false&EventGameTypeIDList%5B5%5D=false&EventGameTypeIDList%5B6%5D=false&Distance=20&GradingIDList%5B0%5D=false&GradingIDList%5B1%5D=false&GradingIDList%5B2%5D=false&GradingIDList%5B3%5D=4&GradingIDList%5B4%5D=5&GradingIDList%5B5%5D=false&GradingIDList%5B6%5D=false&GradingIDList%5B7%5D=false&AgeGroupIDList%5B0%5D=false&AgeGroupIDList%5B1%5D=false&AgeGroupIDList%5B2%5D=false&AgeGroupIDList%5B3%5D=false&AgeGroupIDList%5B4%5D=false&AgeGroupIDList%5B5%5D=14&AgeGroupIDList%5B6%5D=false&AgeGroupIDList%5B7%5D=false&AgeGroupIDList%5B8%5D=false&AgeGroupIDList%5B9%5D=false&AgeGroupIDList%5B10%5D=false&AgeGroupIDList%5B11%5D=false&AgeGroupIDList%5B12%5D=false&AgeGroupIDList%5B13%5D=false&AgeGroupIDList%5B14%5D=false&AgeGroupIDList%5B15%5D=false&AgeGroupIDList%5B16%5D=false&AgeGroupIDList%5B17%5D=false&AgeGroupIDList%5B18%5D=false&AgeGroupIDList%5B19%5D=false&AgeGroupIDList%5B20%5D=false&AgeGroupIDList%5B21%5D=false",
    },
    # ── Dylan ──
    {
        "player": "Dylan",
        "allowed_age_groups": {"9U"},
        "description": "20mi, BS, 9U, G1+G2",
        "url": "https://competitions.lta.org.uk/find?DateFilterType=0&StartDate=2026-04-25&EndDate=2026-07-29&LocationFilterType=0&PostalCode=SW15+5FG&StatusFilterID=4&page=1&EventGameTypeIDList%5B0%5D=1&EventGameTypeIDList%5B1%5D=false&EventGameTypeIDList%5B2%5D=false&EventGameTypeIDList%5B3%5D=false&EventGameTypeIDList%5B4%5D=false&EventGameTypeIDList%5B5%5D=false&EventGameTypeIDList%5B6%5D=false&Distance=20&GradingIDList%5B0%5D=1&GradingIDList%5B1%5D=2&GradingIDList%5B2%5D=false&GradingIDList%5B3%5D=false&GradingIDList%5B4%5D=false&GradingIDList%5B5%5D=false&GradingIDList%5B6%5D=false&GradingIDList%5B7%5D=false&AgeGroupIDList%5B0%5D=false&AgeGroupIDList%5B1%5D=9&AgeGroupIDList%5B2%5D=false&AgeGroupIDList%5B3%5D=false&AgeGroupIDList%5B4%5D=false&AgeGroupIDList%5B5%5D=false&AgeGroupIDList%5B6%5D=false&AgeGroupIDList%5B7%5D=false&AgeGroupIDList%5B8%5D=false&AgeGroupIDList%5B9%5D=false&AgeGroupIDList%5B10%5D=false&AgeGroupIDList%5B11%5D=false&AgeGroupIDList%5B12%5D=false&AgeGroupIDList%5B13%5D=false&AgeGroupIDList%5B14%5D=false&AgeGroupIDList%5B15%5D=false&AgeGroupIDList%5B16%5D=false&AgeGroupIDList%5B17%5D=false&AgeGroupIDList%5B18%5D=false&AgeGroupIDList%5B19%5D=false&AgeGroupIDList%5B20%5D=false&AgeGroupIDList%5B21%5D=false",
    },
    {
        "player": "Dylan",
        "allowed_age_groups": {"10U"},
        "description": "15mi, 10U BS, G4",
        "url": "https://competitions.lta.org.uk/find?DateFilterType=0&StartDate=2026-04-25&EndDate=2026-07-29&LocationFilterType=0&PostalCode=SW15+5FG&StatusFilterID=4&page=1&EventGameTypeIDList%5B0%5D=1&EventGameTypeIDList%5B1%5D=false&EventGameTypeIDList%5B2%5D=false&EventGameTypeIDList%5B3%5D=false&EventGameTypeIDList%5B4%5D=false&EventGameTypeIDList%5B5%5D=false&EventGameTypeIDList%5B6%5D=false&Distance=15&GradingIDList%5B0%5D=false&GradingIDList%5B1%5D=false&GradingIDList%5B2%5D=false&GradingIDList%5B3%5D=4&GradingIDList%5B4%5D=false&GradingIDList%5B5%5D=false&GradingIDList%5B6%5D=false&GradingIDList%5B7%5D=false&AgeGroupIDList%5B0%5D=false&AgeGroupIDList%5B1%5D=false&AgeGroupIDList%5B2%5D=10&AgeGroupIDList%5B3%5D=false&AgeGroupIDList%5B4%5D=false&AgeGroupIDList%5B5%5D=false&AgeGroupIDList%5B6%5D=false&AgeGroupIDList%5B7%5D=false&AgeGroupIDList%5B8%5D=false&AgeGroupIDList%5B9%5D=false&AgeGroupIDList%5B10%5D=false&AgeGroupIDList%5B11%5D=false&AgeGroupIDList%5B12%5D=false&AgeGroupIDList%5B13%5D=false&AgeGroupIDList%5B14%5D=false&AgeGroupIDList%5B15%5D=false&AgeGroupIDList%5B16%5D=false&AgeGroupIDList%5B17%5D=false&AgeGroupIDList%5B18%5D=false&AgeGroupIDList%5B19%5D=false&AgeGroupIDList%5B20%5D=false&AgeGroupIDList%5B21%5D=false",
    },
    {
        "player": "Dylan",
        "allowed_age_groups": {"10U"},
        "description": "50mi, 10U BS, G3",
        "url": "https://competitions.lta.org.uk/find?DateFilterType=0&StartDate=2026-04-25&EndDate=2026-07-29&LocationFilterType=0&PostalCode=SW15+5FG&StatusFilterID=4&page=1&EventGameTypeIDList%5B0%5D=1&EventGameTypeIDList%5B1%5D=false&EventGameTypeIDList%5B2%5D=false&EventGameTypeIDList%5B3%5D=false&EventGameTypeIDList%5B4%5D=false&EventGameTypeIDList%5B5%5D=false&EventGameTypeIDList%5B6%5D=false&Distance=50&AgeGroupIDList%5B0%5D=false&AgeGroupIDList%5B1%5D=false&AgeGroupIDList%5B2%5D=10&AgeGroupIDList%5B3%5D=false&AgeGroupIDList%5B4%5D=false&AgeGroupIDList%5B5%5D=false&AgeGroupIDList%5B6%5D=false&AgeGroupIDList%5B7%5D=false&AgeGroupIDList%5B8%5D=false&AgeGroupIDList%5B9%5D=false&AgeGroupIDList%5B10%5D=false&AgeGroupIDList%5B11%5D=false&AgeGroupIDList%5B12%5D=false&AgeGroupIDList%5B13%5D=false&AgeGroupIDList%5B14%5D=false&AgeGroupIDList%5B15%5D=false&AgeGroupIDList%5B16%5D=false&AgeGroupIDList%5B17%5D=false&AgeGroupIDList%5B18%5D=false&AgeGroupIDList%5B19%5D=false&AgeGroupIDList%5B20%5D=false&AgeGroupIDList%5B21%5D=false&GradingIDList%5B0%5D=false&GradingIDList%5B1%5D=false&GradingIDList%5B2%5D=3&GradingIDList%5B3%5D=false&GradingIDList%5B4%5D=false&GradingIDList%5B5%5D=false&GradingIDList%5B6%5D=false&GradingIDList%5B7%5D=false",
    },
    {
        "player": "Dylan",
        "allowed_age_groups": {"10U"},
        "description": "BS, 10U, G1+G2",
        "url": "https://competitions.lta.org.uk/find?DateFilterType=0&StartDate=2026-04-25&EndDate=2026-07-29&LocationFilterType=0&PostalCode=SW15+5FG&StatusFilterID=4&page=1&AgeGroupIDList%5B0%5D=false&AgeGroupIDList%5B1%5D=false&AgeGroupIDList%5B2%5D=10&AgeGroupIDList%5B3%5D=false&AgeGroupIDList%5B4%5D=false&AgeGroupIDList%5B5%5D=false&AgeGroupIDList%5B6%5D=false&AgeGroupIDList%5B7%5D=false&AgeGroupIDList%5B8%5D=false&AgeGroupIDList%5B9%5D=false&AgeGroupIDList%5B10%5D=false&AgeGroupIDList%5B11%5D=false&AgeGroupIDList%5B12%5D=false&AgeGroupIDList%5B13%5D=false&AgeGroupIDList%5B14%5D=false&AgeGroupIDList%5B15%5D=false&AgeGroupIDList%5B16%5D=false&AgeGroupIDList%5B17%5D=false&AgeGroupIDList%5B18%5D=false&AgeGroupIDList%5B19%5D=false&AgeGroupIDList%5B20%5D=false&AgeGroupIDList%5B21%5D=false&EventGameTypeIDList%5B0%5D=1&EventGameTypeIDList%5B1%5D=false&EventGameTypeIDList%5B2%5D=false&EventGameTypeIDList%5B3%5D=false&EventGameTypeIDList%5B4%5D=false&EventGameTypeIDList%5B5%5D=false&EventGameTypeIDList%5B6%5D=false&GradingIDList%5B0%5D=1&GradingIDList%5B1%5D=2&GradingIDList%5B2%5D=false&GradingIDList%5B3%5D=false&GradingIDList%5B4%5D=false&GradingIDList%5B5%5D=false&GradingIDList%5B6%5D=false&GradingIDList%5B7%5D=false",
    },
]

# All known age group tags (for parsing)
AGE_GROUP_TAGS = {"8U", "9U", "10U", "11U", "12U", "14U", "16U", "18U", "Open",
                  "30+", "35+", "40+", "45+", "50+", "55+", "60+", "65+", "70+", "75+", "80+", "85+", "90+"}

# Grade tags
GRADE_TAGS = {"Grade 1", "Grade 2", "Grade 3", "Grade 4", "Grade 5", "Grade 6", "Grade 7"}


def clean_tournament_name(name: str) -> str:
    """Strip common prefixes from tournament names."""
    return name.replace("London & South East Tour - ", "").strip()


def parse_date(date_str: str) -> str:
    """
    Parse date from search results. Handles single dates and ranges.
    Returns "dd/mm/yyyy" or "dd/mm/yyyy to dd/mm/yyyy".
    """
    if not date_str:
        return ""
    dates = re.findall(r"\b(\d{1,2}/\d{1,2}/\d{4})\b", date_str)
    if len(dates) >= 2:
        return f"{dates[0]} to {dates[1]}"
    elif len(dates) == 1:
        return dates[0]
    return date_str.strip()


def update_search_dates(url: str) -> str:
    """Replace StartDate and EndDate in a URL with today and today+90 days."""
    today = datetime.now().strftime("%Y-%m-%d")
    end = (datetime.now() + timedelta(days=90)).strftime("%Y-%m-%d")
    url = re.sub(r"StartDate=[^&]+", f"StartDate={today}", url)
    url = re.sub(r"EndDate=[^&]+", f"EndDate={end}", url)
    return url


def dismiss_cookie_banner(page):
    """Dismiss cookie consent banner if present."""
    try:
        btn = page.query_selector('button:has-text("ACCEPT"), button:has-text("Accept")')
        if btn:
            btn.click()
            page.wait_for_timeout(800)
            print("  ✓ Cookie banner dismissed")
    except Exception as e:
        print(f"  (Cookie banner: {e})")


def accept_cookies(page):
    """
    Navigate to a simple LTA page first to accept cookies.
    The cookie persists in the browser context for subsequent requests.
    """
    print("  ⓘ  Accepting cookies...")
    page.goto("https://competitions.lta.org.uk/tournaments", wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(2000)

    try:
        btn = page.query_selector('button:has-text("ACCEPT"), a:has-text("ACCEPT")')
        if not btn:
            btn = page.query_selector('button:has-text("Accept"), a:has-text("Accept")')
        if btn:
            btn.click()
            page.wait_for_load_state("domcontentloaded", timeout=15000)
            page.wait_for_timeout(2000)
            print("  ✓ Cookies accepted")
        else:
            print("  ⓘ  No cookie wall found — may already be accepted")
    except Exception as e:
        print(f"  ⚠️  Cookie acceptance error: {e}")


def fetch_search_page(page, url: str) -> BeautifulSoup | None:
    """Navigate to a search results page and return parsed HTML."""
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(3000)

        # Wait for results to render
        try:
            page.locator("#searchResultArea > li.list__item").first.wait_for(timeout=5000)
        except:
            pass  # may genuinely be zero results

        html = page.content()
        return BeautifulSoup(html, "html.parser")
    except Exception as e:
        print(f"  ✗ Error fetching page: {e}")
        return None


def parse_tournaments_from_page(soup: BeautifulSoup, player: str, allowed_age_groups: set) -> list[dict]:
    """
    Parse tournament cards from a search results page.
    Returns a list of dicts, one per tournament per allowed age group.
    Filters out age groups not in allowed_age_groups.
    """
    results = []

    cards = soup.select("#searchResultArea > li.list__item")
    if not cards:
        return results

    for card in cards:
        # Skip ad cards
        if card.select_one("div.ad"):
            continue

        # ── Tournament name + URL ──
        title_link = card.select_one("h4.media__title a")
        if not title_link:
            continue
        raw_name = title_link.get_text(strip=True)
        name = clean_tournament_name(raw_name)
        href = title_link.get("href", "")
        if href and not href.startswith("http"):
            href = "https://competitions.lta.org.uk" + href

        # ── Venue + Distance ──
        venue_el = card.select_one("small.media__subheading .nav-link__value")
        venue_raw = venue_el.get_text(strip=True) if venue_el else ""
        # Extract distance (e.g. "5.6 miles" from "Sutton Tennis Academy | Surrey (5.6 miles)")
        distance = ""
        dist_m = re.search(r"\((\d+(?:\.\d+)?)\s*miles?\)", venue_raw)
        if dist_m:
            distance = dist_m.group(1)
        # Clean venue — take just the venue name before the pipe
        venue = venue_raw
        if "|" in venue:
            venue = venue.split("|")[0].strip()

        # ── Date ──
        date_el = card.select_one("small.media__subheading--muted")
        date_str = parse_date(date_el.get_text(strip=True) if date_el else "")

        # ── Status (check for Cancelled badge) ──
        status = ""
        cancelled_el = card.find(string=re.compile(r"cancelled", re.I))
        if cancelled_el:
            status = "Cancelled"

        # ── Tags: parse grade + age group pairs ──
        # Each ul.list--inline may contain a grade + age groups
        tag_uls = card.select("ul.list--inline")

        grade_age_pairs = []
        for ul in tag_uls:
            tags = [t.get_text(strip=True) for t in ul.select("span.tag")]
            grade = ""
            age_groups = []
            for tag in tags:
                if tag in GRADE_TAGS:
                    grade = tag
                elif tag in AGE_GROUP_TAGS:
                    age_groups.append(tag)
            if grade and age_groups:
                for ag in age_groups:
                    grade_age_pairs.append((grade, ag))
            elif not grade and age_groups:
                for ag in age_groups:
                    grade_age_pairs.append(("", ag))

        # Fallback: flat tag parsing
        if not grade_age_pairs:
            all_tags = [t.get_text(strip=True) for t in card.select("span.tag")]
            grade = ""
            age_groups = []
            for tag in all_tags:
                if tag in GRADE_TAGS:
                    grade = tag
                elif tag in AGE_GROUP_TAGS:
                    age_groups.append(tag)
            if not age_groups:
                age_groups = ["Other"]
            for ag in age_groups:
                grade_age_pairs.append((grade, ag))

        # ── Filter to allowed age groups and create rows ──
        for grade, ag in grade_age_pairs:
            if ag not in allowed_age_groups:
                continue
            if not grade or grade in ("Grade 6", "Grade 7"):
                continue
            results.append({
                "player":     player,
                "venue":      venue,
                "distance":   distance,
                "tournament": name,
                "date":       date_str,
                "age_group":  ag,
                "grade":      grade,
                "status":     status,
                "url":        href,
            })

    return results


def fetch_all_pages(page, search: dict) -> list[dict]:
    """
    Fetch all pages of search results for a search config.
    Follows pagination by incrementing &page=N.
    """
    all_tournaments = []
    base_url = update_search_dates(search["url"])
    player = search["player"]
    allowed = search["allowed_age_groups"]
    page_num = 1

    # Ensure URL has a page parameter
    if "page=" not in base_url:
        base_url += "&page=1"

    while True:
        url = re.sub(r"[&?]page=\d+", f"&page={page_num}", base_url)
        print(f"  Fetching page {page_num}...")

        soup = fetch_search_page(page, url)
        if not soup:
            break

        tournaments = parse_tournaments_from_page(soup, player, allowed)
        if not tournaments:
            if page_num == 1:
                raw_count = len(soup.select("#searchResultArea > li.list__item"))
                if raw_count > 0:
                    print(f"  ℹ️  {raw_count} results found but none matched allowed age groups")
                else:
                    print(f"  ℹ️  No results found")
            break

        all_tournaments.extend(tournaments)
        print(f"  ✓ Page {page_num}: {len(tournaments)} rows (filtered)")

        next_link = soup.select_one(f'a[href*="page={page_num + 1}"]')
        if not next_link:
            break

        page_num += 1

    return all_tournaments


def make_key(row: dict) -> str:
    """
    Composite key for deduplication:
    player + tournament + venue + date + age_group + grade (all lowercased).
    """
    return "|".join([
        row.get("player", "").lower().strip(),
        row.get("tournament", "").lower().strip(),
        row.get("venue", "").lower().strip(),
        row.get("date", "").lower().strip(),
        row.get("age_group", "").lower().strip(),
        row.get("grade", "").lower().strip(),
    ])


def read_existing_data() -> list[dict]:
    """Read existing tournament data from the Tournament_Venues sheet."""
    if not WEBAPP_URL:
        print("  ⚠️  WEBAPP_URL not set — cannot read existing data")
        return []

    url = WEBAPP_URL
    if WEBAPP_SECRET:
        joiner = "&" if "?" in url else "?"
        url = f"{url}{joiner}secret={WEBAPP_SECRET}"

    joiner = "&" if "?" in url else "?"
    read_url = f"{url}{joiner}action=read&sheet={SHEET_NAME}"

    try:
        resp = requests.get(read_url, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        rows = data.get("rows", [])
    except Exception as e:
        print(f"  ⚠️  Error reading sheet: {e}")
        return []

    if len(rows) < 2:
        return []

    existing = []
    for row in rows[1:]:
        if len(row) < 11:
            row = row + [""] * (11 - len(row))
        existing.append({
            "player":     row[0].strip(),
            "venue":      row[1].strip(),
            "distance":   row[2].strip(),
            "tournament": row[3].strip(),
            "date":       row[4].strip(),
            "age_group":  row[5].strip(),
            "grade":      row[6].strip(),
            "status":     row[7].strip(),
            "url":        row[8].strip(),
            "first_seen": row[9].strip(),
            "new":        row[10].strip(),
        })

    return existing


def write_to_sheet(tournaments: list[dict]):
    """Write the full tournament list to the Tournament_Venues sheet."""
    if not WEBAPP_URL:
        print("  ⚠️  WEBAPP_URL not set — cannot write to sheet")
        return

    url = WEBAPP_URL
    if WEBAPP_SECRET:
        joiner = "&" if "?" in url else "?"
        url = f"{url}{joiner}secret={WEBAPP_SECRET}"

    header = ["Player", "Venue", "Distance", "Tournament", "Date", "Age Group", "Grade", "Status", "URL", "First Seen", "New"]

    grid = [header]
    for t in tournaments:
        grid.append([
            t.get("player", ""),
            t.get("venue", ""),
            t.get("distance", ""),
            t.get("tournament", ""),
            t.get("date", ""),
            t.get("age_group", ""),
            t.get("grade", ""),
            t.get("status", ""),
            t.get("url", ""),
            t.get("first_seen", ""),
            t.get("new", ""),
        ])

    payload = {
        "sheet":      SHEET_NAME,
        "clearFirst": True,
        "rows":       grid,
    }

    try:
        resp = requests.post(url, json=payload, timeout=60)
        resp.raise_for_status()
        print(f"  ✓ Wrote {len(grid) - 1} rows to '{SHEET_NAME}'")
    except Exception as e:
        print(f"  ✗ Error writing to sheet: {e}")


def merge_and_flag(new_tournaments: list[dict], existing: list[dict]) -> list[dict]:
    """
    Compare new scrape results against existing data.
    - New tournaments get New="Y" and First Seen=today
    - Existing tournaments now cancelled get New="Y"
    - All others keep First Seen and clear New
    """
    today = datetime.now().strftime("%d/%m/%Y")

    existing_map = {}
    for row in existing:
        key = make_key(row)
        existing_map[key] = row

    merged = []
    for t in new_tournaments:
        key = make_key(t)
        old = existing_map.get(key)

        if old is None:
            t["first_seen"] = today
            t["new"] = "Y"
            merged.append(t)
        else:
            t["first_seen"] = old.get("first_seen", today)
            if t.get("status") == "Cancelled" and old.get("status") != "Cancelled":
                t["new"] = "Y"
            else:
                t["new"] = ""
            merged.append(t)

    return merged


def deduplicate(tournaments: list[dict]) -> list[dict]:
    """Remove duplicate rows across multiple searches for the same player."""
    seen = set()
    unique = []
    for t in tournaments:
        key = make_key(t)
        if key not in seen:
            seen.add(key)
            unique.append(t)
    return unique


def main():
    print("=" * 60)
    print("LTA Venue Monitor")
    print(f"Run: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"Date range: {datetime.now().strftime('%Y-%m-%d')} to {(datetime.now() + timedelta(days=90)).strftime('%Y-%m-%d')}")
    print("=" * 60)

    # ── Step 1: Scrape all searches ──
    all_new = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        bpage = context.new_page()

        accept_cookies(bpage)

        for search in SEARCHES:
            print(f"\n👤 {search['player']} — {search['description']}")
            tournaments = fetch_all_pages(bpage, search)
            all_new.extend(tournaments)
            print(f"  Total: {len(tournaments)} rows")

        bpage.close()
        context.close()
        browser.close()

    # Deduplicate across searches
    all_new = deduplicate(all_new)

    print(f"\n{'─'*60}")
    print(f"Total scraped: {len(all_new)} unique rows across {len(SEARCHES)} search(es)")

    # ── Step 2: Read existing data ──
    print(f"\nReading existing data from '{SHEET_NAME}'...")
    existing = read_existing_data()
    print(f"  Existing rows: {len(existing)}")

    # ── Step 3: Compare and flag ──
    merged = merge_and_flag(all_new, existing)
    new_count = sum(1 for t in merged if t.get("new") == "Y")
    print(f"\n  New/changed: {new_count}")
    if new_count:
        for t in merged:
            if t.get("new") == "Y":
                status_note = f" [{t['status']}]" if t.get("status") else ""
                print(f"    🆕 [{t['player']}] {t['tournament']} | {t['date']} | {t['age_group']} | {t['grade']}{status_note}")

    # ── Step 4: Write back ──
    print(f"\nWriting {len(merged)} rows to '{SHEET_NAME}'...")
    write_to_sheet(merged)

    # ── Step 5: Write timestamp to cell N1 ──
    if WEBAPP_URL:
        ts_url = WEBAPP_URL
        if WEBAPP_SECRET:
            joiner = "&" if "?" in ts_url else "?"
            ts_url = f"{ts_url}{joiner}secret={WEBAPP_SECRET}"
        try:
            requests.post(ts_url, json={
                "sheet":      SHEET_NAME,
                "clearFirst": False,
                "startRow":   1,
                "startCol":   14,
                "rows":       [[f"Last run: {datetime.now().strftime('%d/%m/%Y %H:%M')}"]],
            }, timeout=30)
            print(f"  ✓ Timestamp written to N1")
        except Exception as e:
            print(f"  ⚠️  Failed to write timestamp: {e}")

    print(f"\n✅ Venue monitor complete!")


if __name__ == "__main__":
    main()
    os._exit(0)
