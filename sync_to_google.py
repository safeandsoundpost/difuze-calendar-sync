"""Pull events from Cloudflare Worker and sync to Google Calendar."""
import os
import sys
import pickle
from datetime import datetime, timedelta

import httpx
import pytz
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from dotenv import load_dotenv

load_dotenv()

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_PATH = os.path.join(SCRIPT_DIR, "token.pickle")
CREDS_PATH = os.path.join(SCRIPT_DIR, "google_credentials.json")
SCOPES = ["https://www.googleapis.com/auth/calendar"]
TZ = pytz.timezone("America/Toronto")
CALENDAR_NAME = "DIFUZE - STUDIO CALENDAR"
EVENTS_URL = "https://difuze-calendar-sync.safeandsoundpost.workers.dev/api/events"


def get_credentials():
    creds = None
    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, "rb") as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "wb") as f:
            pickle.dump(creds, f)
    return creds


def find_calendar(service, name):
    for cal in service.calendarList().list().execute().get("items", []):
        if cal["summary"] == name:
            return cal["id"]
    raise SystemExit(f"Calendar '{name}' not found.")


def sync():
    print("Fetching events from Worker…")
    r = httpx.get(EVENTS_URL, timeout=30)
    r.raise_for_status()
    data = r.json()
    events = data.get("events", [])
    if not events:
        print("No events to sync.")
        return

    print(f"Got {len(events)} events. Authenticating with Google…")
    creds = get_credentials()
    service = build("calendar", "v3", credentials=creds)
    cal_id = find_calendar(service, CALENDAR_NAME)
    print(f"Found calendar: {CALENDAR_NAME}")

    # Determine date range from scraped events
    starts = []
    for e in events:
        try:
            s = e["start"]
            dt = datetime.fromisoformat(s) if isinstance(s, str) else datetime.fromisoformat(s.get("dateTime", s.get("DateTime", "")))
            starts.append(TZ.localize(dt) if dt.tzinfo is None else dt)
        except Exception:
            pass

    if not starts:
        print("No valid dates in events.")
        return
    first, last = min(starts), max(starts)
    print(f"Date range: {first.date()} → {last.date()}")

    # Clear existing (paginate to handle 250-event page limit)
    print("Clearing old events…")
    deleted = 0
    page_token = None
    while True:
        existing = service.events().list(
            calendarId=cal_id,
            timeMin=first.isoformat(),
            timeMax=(last + timedelta(days=1)).isoformat(),
            singleEvents=True,
            pageToken=page_token,
        ).execute()
        for ev in existing.get("items", []):
            try:
                service.events().delete(calendarId=cal_id, eventId=ev["id"]).execute()
                deleted += 1
            except Exception as ex:
                print(f"  Delete error: {ex}")
        page_token = existing.get("nextPageToken")
        if not page_token:
            break
    print(f"Deleted {deleted} old events.")

    # Import
    imported = 0
    errors = []
    for e in events:
        try:
            s = e["start"]
            end = e["end"]
            start_dt = datetime.fromisoformat(s) if isinstance(s, str) else datetime.fromisoformat(s.get("dateTime", s.get("DateTime", "")))
            end_dt = datetime.fromisoformat(end) if isinstance(end, str) else datetime.fromisoformat(end.get("dateTime", end.get("DateTime", "")))
            if start_dt.tzinfo is None:
                start_dt = TZ.localize(start_dt)
            if end_dt.tzinfo is None:
                end_dt = TZ.localize(end_dt)

            body = {
                "summary": e["title"],
                "start": {"dateTime": start_dt.isoformat(), "timeZone": "America/Toronto"},
                "end": {"dateTime": end_dt.isoformat(), "timeZone": "America/Toronto"},
            }
            if e.get("location"):
                body["location"] = e["location"]

            service.events().insert(calendarId=cal_id, body=body).execute()
            imported += 1
        except Exception as ex:
            errors.append(f"{e.get('title', '?')}: {ex}")

    print(f"Imported {imported} events.")
    if errors:
        print(f"Errors ({len(errors)}):")
        for err in errors:
            print(f"  - {err}")


if __name__ == "__main__":
    sync()
