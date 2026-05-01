# Handoff for claude-ds

**From:** Claude Code session (deepseek-v4-pro)
**To:** SAS scheduler claude-ds session
**Date:** 2026-04-30
**Project directory:** `/Users/thom/Downloads/CALENDAR SYNC SCRIPT`
**User:** safeandsoundpost@gmail.com

---

## Project overview

Calendar sync pipeline: Outlook Web (shared calendar "Toronto Post") → Cloudflare Worker (KV) → Google Calendar ("DIFUZE - STUDIO CALENDAR"). Currently running on a Raspberry Pi at `thom@192.168.0.21`.

## Architecture

```
Outlook Web                  Cloudflare Worker                 Google Calendar
(Toronto Post)  ──→  KV (events.json)  ──→  DIFUZE - STUDIO CALENDAR
     ↑                    ↑                        ↑
  scraper/run.py     difuze-calendar-sync      sync_to_google.py
  (direct OWA API)   .safeandsoundpost         (Google API)
                         .workers.dev
                              │
                    ┌─────────┴─────────┐
                    │  /api/ingest (POST)│
                    │  /api/events (GET) │
                    │  /api/health (GET) │
                    │  /api/events/stream│
                    └────────────────────┘
```

## Key files

| File | Purpose |
|------|---------|
| `scraper/scrape.py` | Direct OWA API calls using CalendarId from `automation_config.json`. Opens headless Chromium briefly for MSAL token refresh, then calls OWA `GetCalendarView` for each week. No sidebar interaction needed. |
| `scraper/run.py` | Wraps scraper, pushes to Worker |
| `scraper/login.py` | One-time interactive login. Saves to `browser_profile/`. Needs display (X forwarding on Pi). |
| `scraper/automation_config.json` | Contains Toronto Post CalendarId (`AQMk...`) and MSAL refresh token. Captured once, used forever. |
| `scraper/browser_profile/` | Persistent Chromium profile with MSAL tokens. Copied from Mac to Pi. |
| `sync_to_google.py` | Fetches from Worker, syncs to Google Calendar. Uses `token.pickle` + `google_credentials.json` for OAuth. |
| `worker/index.js` | Cloudflare Worker. KV-backed. Endpoints: `/api/health`, `/api/events`, `/api/ingest` (bearer auth), `/api/events/stream` (SSE). |
| `pipeline.sh` | Runs `scraper/run.py && sync_to_google.py` with `CHROMIUM_PATH=/snap/bin/chromium` |
| `.env` | `INGEST_URL`, `INGEST_TOKEN`, `OUTLOOK_CALENDAR_NAME`, `SCRAPER_LOOKAHEAD_DAYS`, `TIMEZONE` |

## Raspberry Pi (192.168.0.21)

- **OS:** Ubuntu 24.04.3 LTS, ARM64, 4GB RAM
- **User:** thom
- **SSH:** Key-based from Mac (`~/.ssh/id_ed25519`)
- **Project path:** `~/calendar-sync/`
- **Venv:** `.venv/` (Python 3.12.3)
- **Chromium:** Snap at `/snap/bin/chromium` (147.0.7727.116)
- **Playwright:** Installed without bundled browsers (`PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1`), uses system Chromium via `executable_path`
- **Scheduler:** Cron — runs 4x daily (8:15, 12:15, 4:15, 8:15): `pipeline.sh >> pipeline.log`
- **Auth:** MSAL tokens from Mac profile copied via rsync. Works. May expire ~90 days; re-run `login.py` with `ssh -X thom@192.168.0.21` if scraper fails with "Failed to get auth token".

## Cloudflare Worker

- **Name:** difuze-calendar-sync
- **URL:** https://difuze-calendar-sync.safeandsoundpost.workers.dev
- **KV namespace:** EVENTS_KV (0a232bde77c54289bdfda64c36e892b2)
- **Secret:** INGEST_TOKEN set via `wrangler secret put`
- **Wrangler config:** `worker/wrangler.toml`

## What was just requested (next task)

**Create a "Sync Now" button on a Cloudflare Pages site (behind Zero Trust) that triggers the pipeline.** The approach discussed:

1. Add `/api/trigger` (POST) and `/api/trigger/status` (GET) to the Worker — stores a trigger flag in KV
2. The Pi runs a lightweight polling script (every 30s) checking `/api/trigger/status`
3. When "pending", runs `pipeline.sh` and calls back to clear the trigger
4. A simple HTML page (served by Cloudflare Pages or the Worker itself) with a button that POSTs to `/api/trigger`
5. The HTML page can be behind Zero Trust for security

The Worker at `worker/index.js` needs updating, and a new polling script (`poll-trigger.sh` or similar) needs to be created on the Pi.

## Important notes

- The scraper uses the OWA `service.svc?action=GetCalendarView` endpoint with a pre-captured CalendarId. It does NOT need the shared calendar to be visible in the sidebar.
- Google OAuth (`token.pickle`) uses `google_credentials.json` — a Desktop OAuth client. Do not commit these.
- The `browser_profile/` directory is a full Chromium user data directory with MSAL auth. Treat it like a secret.
- Mac launchd job was removed — Pi handles everything now.
- All sensitive files are in `.gitignore` but this is not a git repo.
