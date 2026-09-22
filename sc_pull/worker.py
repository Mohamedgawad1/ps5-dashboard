import os
import sys
import time

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "profile")
SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)


def main():
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            PROFILE_DIR,
            channel="chrome",
            headless=False,
            args=[
                "--start-maximized",
                "--remote-debugging-port=9222",
                "--no-first-run",
                "--no-default-browser-check",
            ],
            viewport=None,
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=90000)
        print("WORKER_READY url=%s" % page.url, flush=True)
        while True:
            time.sleep(3600)


if __name__ == "__main__":
    sys.exit(main())