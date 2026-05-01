#!/usr/bin/env python3
"""Poll the Worker for sync triggers — run as a daemon on the Pi.

Usage: python3 poll-trigger.py
       (run via systemd or screen for persistence)

Checks GET /api/trigger/status every 30s. When a pending trigger is found:
  1. Claims it (POST /api/trigger/result {status: "running"})
  2. Runs pipeline.py
  3. Reports back (POST /api/trigger/result {status: "completed"|"failed", result})
"""

import os
import sys
import json
import time
import subprocess
from pathlib import Path

import httpx
from dotenv import load_dotenv

ROOT = Path(__file__).parent
load_dotenv(ROOT / ".env")

WORKER_BASE = os.environ.get(
    "WORKER_URL",
    "https://difuze-calendar-sync.safeandsoundpost.workers.dev",
)
INGEST_TOKEN = os.environ["INGEST_TOKEN"]
TRIGGER_STATUS_URL = f"{WORKER_BASE}/api/trigger/status"
TRIGGER_RESULT_URL = f"{WORKER_BASE}/api/trigger/result"
PIPELINE = ROOT / "pipeline.py"
POLL_INTERVAL = 30

HEADERS = {"Authorization": f"Bearer {INGEST_TOKEN}"}


def poll():
    """GET /api/trigger/status — return trigger dict or None."""
    try:
        r = httpx.get(TRIGGER_STATUS_URL, timeout=10)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"[poll-trigger] Failed to check trigger status: {e}")
        return None


def claim(trigger):
    """POST /api/trigger/result with status=running to claim the trigger."""
    try:
        r = httpx.post(
            TRIGGER_RESULT_URL,
            json={"status": "running"},
            headers=HEADERS,
            timeout=10,
        )
        r.raise_for_status()
        return True
    except Exception as e:
        print(f"[poll-trigger] Failed to claim trigger: {e}")
        return False


def report(status, result=None):
    """POST /api/trigger/result with final status and result."""
    try:
        body = {"status": status}
        if result:
            body["result"] = result
        r = httpx.post(
            TRIGGER_RESULT_URL,
            json=body,
            headers=HEADERS,
            timeout=10,
        )
        r.raise_for_status()
    except Exception as e:
        print(f"[poll-trigger] Failed to report result: {e}")


def run_pipeline():
    """Run pipeline.py and capture structured output."""
    try:
        r = subprocess.run(
            [sys.executable, str(PIPELINE)],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            timeout=300,  # 5 min max
        )
        output = (r.stdout + r.stderr)[-2000:]  # last 2000 chars
        return {
            "exit_code": r.returncode,
            "ok": r.returncode == 0,
            "output": output,
        }
    except subprocess.TimeoutExpired:
        return {"exit_code": -1, "ok": False, "output": "Pipeline timed out after 5 min"}
    except Exception as e:
        return {"exit_code": -1, "ok": False, "output": str(e)}


def main():
    print(f"[poll-trigger] Starting. Polling {TRIGGER_STATUS_URL} every {POLL_INTERVAL}s")
    while True:
        trigger = poll()
        if trigger and trigger.get("status") == "pending":
            trigger_id = trigger.get("id", "?")
            print(f"[poll-trigger] Trigger {trigger_id} is pending. Claiming…")
            if not claim(trigger):
                print("[poll-trigger] Claim failed, will retry next poll")
                time.sleep(POLL_INTERVAL)
                continue

            print("[poll-trigger] Running pipeline…")
            result = run_pipeline()
            status = "completed" if result["ok"] else "failed"
            print(f"[poll-trigger] Pipeline {status} (exit {result['exit_code']})")
            report(status, result)
            print(f"[poll-trigger] Result reported. Back to polling.")

        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()
