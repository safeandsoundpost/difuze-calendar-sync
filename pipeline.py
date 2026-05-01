#!/usr/bin/env python3
"""Run scraper + sync in sequence. Called by launchd on a schedule.

Usage: python3 pipeline.py [--scrape-only] [--sync-only]
"""

import os
import sys
import subprocess
from pathlib import Path

ROOT = Path(__file__).parent
SCRAPER = ROOT / "scraper" / "run.py"
SYNC = ROOT / "sync_to_google.py"


def run(script, label):
    print(f"\n{'='*50}\n  {label}\n{'='*50}")
    r = subprocess.run(
        [sys.executable, str(script)],
        cwd=str(ROOT),
        env={**os.environ, "PATH": os.environ["PATH"]},
        capture_output=False,
    )
    if r.returncode != 0:
        print(f"[ERROR] {label} failed (exit {r.returncode})", file=sys.stderr)
        return False
    return True


def main():
    scrape_only = "--scrape-only" in sys.argv
    sync_only = "--sync-only" in sys.argv

    success = True
    if not sync_only:
        success &= run(SCRAPER, "Scraping Outlook → Worker")

    if not scrape_only:
        if success or sync_only:
            success &= run(SYNC, "Syncing Worker → Google Calendar")

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
