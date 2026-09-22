import json
import os
import sys
import time

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
PROFILE_DIR = os.path.join(BASE_DIR, "profile")

SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)
EDIT_URL = "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/EditForm.aspx?ID=1369&v=vExports"

STEPS = ["Start", "Criteria", "Location & Scheduling", "Export"]


def count_products(page):
    try:
        return page.evaluate(
            """() => (typeof switchboardConfiguration!=="undefined" && switchboardConfiguration.Products
            && switchboardConfiguration.Products.length) ? switchboardConfiguration.Products.length : -1"""
        )
    except Exception:
        return -1


def wait_products(page, timeout_sec):
    start = time.time()
    while time.time() - start < timeout_sec:
        if count_products(page) >= 0:
            return True
        try:
            page.wait_for_timeout(2000)
        except Exception:
            pass
    return False


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            PROFILE_DIR,
            channel="chrome",
            headless=False,
            args=["--start-maximized"],
            viewport=None,
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
        if not wait_products(page, timeout_sec=120):
            print("NO SESSION (will poll longer)", flush=True)
            for i in range(60):
                if count_products(page) >= 0:
                    break
                page.wait_for_timeout(5000)
        if count_products(page) < 0:
            page.bring_to_front()
            if not wait_products(page, timeout_sec=900):
                ctx.close()
                return 1

        tab = ctx.new_page()
        try:
            tab.goto(EDIT_URL, wait_until="commit", timeout=45000)
        except Exception as e:
            print("goto note:", str(e)[:100], flush=True)
        for i in range(30):
            tab.wait_for_timeout(5000)
            try:
                bl = tab.evaluate("() => document.body.innerText.length")
                if bl > 500:
                    break
            except Exception:
                pass

        out = {}
        for step in STEPS:
            print("== step:", step, flush=True)
            try:
                tab.evaluate(
                    """(txt) => {
                        const els = Array.from(document.querySelectorAll('a, span, div, li, td, button'));
                        for (const el of els) {
                            const t = (el.innerText || '').trim();
                            if (t === txt) { el.click(); return true; }
                        }
                        return false;
                    }""",
                    step,
                )
            except Exception as e:
                print("  click err:", str(e)[:100], flush=True)
            tab.wait_for_timeout(6000)
            try:
                body = tab.evaluate("() => document.body.innerText")
            except Exception as e:
                body = ""
            out[step] = body[:5500]
            with open(os.path.join(BASE_DIR, "step_%s.txt" % step.replace(" ", "_").replace("&", "_")), "w", encoding="utf-8") as f:
                f.write(body)
            print("   len:", len(body), flush=True)
            for kw in ["Task State", "Discipline", "PS5", "CPP", "AGI", "System", "Closed", "Open", "In Progress", "Criteria", "Project", "Company", "Phase", "Work Package", "Responsible"]:
                k = body.find(kw)
                if k >= 0:
                    print("   [%s] %s" % (kw, body[max(0,k-110):k+220].replace("\n", " | ")), flush=True)
        print("saved step files", flush=True)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())