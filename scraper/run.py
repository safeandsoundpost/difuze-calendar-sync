"""Scrape and POST to the Cloudflare Worker ingest endpoint."""
import os
import sys
from pathlib import Path

import httpx
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent / ".env")
from scrape import scrape  # noqa: E402

INGEST_URL = os.environ["INGEST_URL"]
INGEST_TOKEN = os.environ["INGEST_TOKEN"]


def main():
    result = scrape()
    if result["count"] == 0:
        print("No events captured — the calendar may need a fresh login.", file=sys.stderr)
        print("Run: python scraper/login.py", file=sys.stderr)
        sys.exit(2)

    r = httpx.post(
        INGEST_URL,
        json=result,
        headers={"Authorization": f"Bearer {INGEST_TOKEN}"},
        timeout=30,
    )
    r.raise_for_status()
    print(f"Pushed {result['count']} events. Server: {r.json()}")


if __name__ == "__main__":
    main()
