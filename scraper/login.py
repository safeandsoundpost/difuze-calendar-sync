"""Interactive login — saves to persistent browser profile for reuse by scrape.py.

Run this once to sign in. The profile is shared with the headless scraper
so calendars you add/select here remain available there.
"""

from pathlib import Path
from playwright.sync_api import sync_playwright

PROFILE_DIR = Path(__file__).parent / "browser_profile"
OUTLOOK_URL = "https://outlook.office.com/calendar/view/week"


def main():
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1920, "height": 1080},
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0.0.0 Safari/537.36"
            ),
        )
        page = context.new_page()
        page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', { get: () => false });
            window.chrome = { runtime: {} };
        """)
        page.goto(OUTLOOK_URL)

        print("Sign in with your Difuze account in the browser window.")
        print("IMPORTANT: Make sure 'Toronto Post' is visible in the sidebar.")
        print("If it's not selected, click it now.")
        print("Once the calendar fully loads, return here and press Enter.")
        input()

        page.close()
        context.close()
        print(f"Profile saved to {PROFILE_DIR}")


if __name__ == "__main__":
    main()
