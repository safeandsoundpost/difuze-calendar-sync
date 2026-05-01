"""Scrape events from Outlook calendar via direct OWA API calls.

Uses a persistent browser profile for silent MSAL token refresh (headless),
then makes direct OWA GetCalendarView API calls for each week in the lookahead
window. No sidebar interaction needed — the CalendarId is stored once.
"""

import json
import os
import urllib.parse
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any

import httpx
from playwright.sync_api import sync_playwright

PROFILE_DIR = Path(__file__).parent / "browser_profile"
CONFIG_PATH = Path(__file__).parent / "automation_config.json"
CALENDAR_NAME = os.environ.get("OUTLOOK_CALENDAR_NAME", "Toronto Post")
LOOKAHEAD_DAYS = int(os.environ.get("SCRAPER_LOOKAHEAD_DAYS", "60"))

OWA_URL = "https://outlook.cloud.microsoft/owa/service.svc"
TIMEZONE_ID = "Eastern Standard Time"


def _load_calendar_id() -> str | None:
    if not CONFIG_PATH.exists():
        return None
    with open(CONFIG_PATH) as f:
        config = json.load(f)
    # Toronto Post = AQMk prefix
    for cid in config.get("calendar_ids", {}):
        if cid.startswith("AQMk"):
            return cid
    return None


def _get_fresh_auth() -> dict[str, str]:
    """Open persistent profile, let MSAL refresh tokens silently, capture auth."""
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    auth = {}

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=True,
            executable_path=os.environ.get("CHROMIUM_PATH"),
            args=[
                "--headless=new",
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
            ],
            viewport={"width": 1920, "height": 1080},
        )
        page = context.new_page()

        def handle_route(route):
            url = route.request.url
            if "GetCalendarView" in url and not auth:
                hdrs = {k: v for k, v in route.request.headers.items()}
                auth["token"] = hdrs.get("authorization", "")
                auth["session_id"] = hdrs.get("x-owa-sessionid", "")
                auth["anchor"] = hdrs.get("x-anchormailbox", "")
            route.continue_()

        page.route("**/*", handle_route)
        page.goto(
            "https://outlook.office.com/calendar/view/week",
            wait_until="load",
            timeout=60000,
        )
        page.wait_for_timeout(8000)  # wait for MSAL silent refresh

        # Click Next if no GetCalendarView fired yet
        if not auth:
            next_btn = page.locator('button[aria-label*="next week" i]')
            if next_btn.count() > 0:
                next_btn.first.click()
                page.wait_for_timeout(3000)

        page.close()
        context.close()

    return auth


def _normalize(raw: dict[str, Any]) -> dict[str, Any] | None:
    subject = raw.get("Subject") or raw.get("subject") or raw.get("NormalizedSubject")
    if not subject:
        return None
    start = raw.get("Start") or raw.get("start") or {}
    end = raw.get("End") or raw.get("end") or {}
    if isinstance(start, str):
        start_dt = start
    else:
        start_dt = start.get("DateTime") or start.get("dateTime") or raw.get("StartTime")
    if isinstance(end, str):
        end_dt = end
    else:
        end_dt = end.get("DateTime") or end.get("dateTime") or raw.get("EndTime")
    if not start_dt or not end_dt:
        return None
    location = raw.get("Location") or raw.get("location") or {}
    if isinstance(location, dict):
        location = location.get("DisplayName") or location.get("displayName") or ""
    uid = (
        raw.get("ItemId", {}).get("Id")
        if isinstance(raw.get("ItemId"), dict)
        else raw.get("Id") or raw.get("id") or raw.get("iCalUId") or raw.get("UID")
    )
    organizer = ""
    org = raw.get("Organizer") or raw.get("organizer") or {}
    if isinstance(org, dict):
        eb = org.get("EmailAddress") or org.get("Mailbox") or {}
        if isinstance(eb, dict):
            organizer = eb.get("Name", "")
    return {
        "uid": uid,
        "title": subject,
        "start": start_dt,
        "end": end_dt,
        "location": location or "",
        "all_day": bool(raw.get("IsAllDayEvent") or raw.get("isAllDay")),
        "organizer": organizer,
    }


def _extract_events(payload: Any) -> list[dict[str, Any]]:
    out = []
    def walk(node):
        if isinstance(node, dict):
            if node.get("__type") and "CalendarItem" in str(node.get("__type", "")):
                ev = _normalize(node)
                if ev:
                    out.append(ev)
                    return
            if (node.get("Subject") or node.get("subject")) and (
                node.get("Start") or node.get("start") or node.get("StartTime")
            ):
                ev = _normalize(node)
                if ev:
                    out.append(ev)
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for item in node:
                walk(item)
    walk(payload)
    return out


def scrape() -> dict[str, Any]:
    cal_id = _load_calendar_id()
    if not cal_id:
        raise SystemExit(
            "No Toronto Post CalendarId found. Run python scraper/login.py once, "
            "then the capture script to save it."
        )

    print("Getting fresh auth token…")
    auth = _get_fresh_auth()
    token = auth.get("token", "")
    if not token:
        raise SystemExit("Failed to get auth token. Run python scraper/login.py to re-login.")

    print(f"CalendarId: {cal_id[:50]}...")

    headers = {
        "Authorization": token,
        "Content-Type": "application/json; charset=utf-8",
        "Action": "GetCalendarView",
    }

    all_events: list[dict[str, Any]] = []
    seen_uids: set[str] = set()
    client = httpx.Client(timeout=30)

    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    start_date = today - timedelta(weeks=2)  # backfill for spanning events
    weeks = (LOOKAHEAD_DAYS // 7) + 4

    for w in range(weeks):
        range_start = (start_date + timedelta(weeks=w)).strftime("%Y-%m-%dT00:00:00.000")
        range_end = (start_date + timedelta(weeks=w + 1)).strftime("%Y-%m-%dT00:00:00.000")

        body = {
            "__type": "GetCalendarViewJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2018_01_08",
                "TimeZoneContext": {
                    "__type": "TimeZoneContext:#Exchange",
                    "TimeZoneDefinition": {
                        "__type": "TimeZoneDefinitionType:#Exchange",
                        "Id": TIMEZONE_ID,
                    },
                },
            },
            "Body": {
                "__type": "GetCalendarViewRequest:#Exchange",
                "CalendarId": {
                    "__type": "TargetFolderId:#Exchange",
                    "BaseFolderId": {
                        "__type": "FolderId:#Exchange",
                        "Id": cal_id,
                    },
                },
                "RangeStart": range_start,
                "RangeEnd": range_end,
                "ClientSupportsIrm": True,
                "OptimizeExtendedPropertyLoading": True,
            },
        }

        body_json = json.dumps(body)
        headers["X-OWA-UrlPostData"] = urllib.parse.quote(body_json, safe="")

        try:
            r = client.post(
                OWA_URL,
                params={"action": "GetCalendarView", "app": "Calendar", "n": str(12 + w)},
                content=body_json,
                headers=headers,
            )
            if r.status_code == 200:
                data = r.json()
                events = _extract_events(data)
                for ev in events:
                    uid = ev.get("uid") or f"{ev['title']}|{ev['start']}"
                    if uid not in seen_uids:
                        seen_uids.add(uid)
                        all_events.append(ev)
                if events:
                    print(f"  Week {range_start[:10]}: {len(events)} events (total: {len(all_events)})")
            else:
                print(f"  Week {range_start[:10]}: HTTP {r.status_code}")
        except Exception as e:
            print(f"  Week {range_start[:10]}: Error: {e}")

    client.close()

    return {
        "scraped_at": datetime.now(timezone.utc).isoformat(),
        "calendar": CALENDAR_NAME,
        "count": len(all_events),
        "events": all_events,
    }


if __name__ == "__main__":
    result = scrape()
    print(json.dumps(result, indent=2, default=str))
