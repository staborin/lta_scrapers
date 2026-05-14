"""
LTA Tournament Scraper - Master Runner
Runs Dylan, Luka, Serge, and Watchlist scrapers sequentially.
Each script runs in its own subprocess and must fully exit before the next begins.
"""

import subprocess
import sys
import logging
from datetime import datetime
from pathlib import Path


# ── Paths to the individual scrapers ─────────────────────────────────────────
# Update these paths to match where your scripts actually live.
SCRIPT_DIR = Path(__file__).resolve().parent

SCRAPERS = [
    ("DYLAN",      SCRIPT_DIR / "LTA_Entered_Tournaments_scraper_DYLAN.py"),
    ("LUKA",       SCRIPT_DIR / "LTA_Entered_Tournaments_scraper_LUKA.py"),
    ("SERGE",      SCRIPT_DIR / "LTA_Entered_Tournaments_scraper_SERGE.py"),
    ("WATCHLIST",  SCRIPT_DIR / "LTA_Watchlist_Tournaments.py"),
]


# ── Logger setup ──────────────────────────────────────────────────────────────
def setup_logger() -> logging.Logger:
    logger = logging.getLogger("lta_master_runner")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    fh = logging.FileHandler(f"lta_master_runner_{ts}.log", encoding="utf-8")
    fh.setFormatter(fmt)

    sh = logging.StreamHandler()
    sh.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(sh)
    return logger


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    logger = setup_logger()
    logger.info("=== LTA Master Tournament Runner starting ===")

    results = {}

    for name, script_path in SCRAPERS:
        if not script_path.exists():
            logger.error(f"[{name}] Script not found: {script_path}")
            results[name] = "SKIPPED (file not found)"
            continue

        logger.info(f"[{name}] Starting scraper: {script_path.name}")

        try:
            result = subprocess.run(
                [sys.executable, str(script_path)],
                check=False,          # don't raise on non-zero exit; we handle it below
            )

            if result.returncode == 0:
                logger.info(f"[{name}] ✅ Finished successfully (exit code 0)")
                results[name] = "OK"
            else:
                logger.error(f"[{name}] ❌ Exited with code {result.returncode}")
                results[name] = f"FAILED (exit code {result.returncode})"

        except Exception as e:
            logger.exception(f"[{name}] ❌ Unexpected error launching script: {e}")
            results[name] = f"ERROR: {e}"

    # ── Summary ───────────────────────────────────────────────────────────────
    logger.info("=== Run complete. Summary ===")
    for name, status in results.items():
        logger.info(f"  {name}: {status}")

    any_failed = any(s != "OK" for s in results.values())
    if any_failed:
        logger.warning("One or more scrapers did not complete successfully. Check logs above.")
        sys.exit(1)
    else:
        logger.info("All scrapers completed successfully.")
        sys.exit(0)


if __name__ == "__main__":
    main()