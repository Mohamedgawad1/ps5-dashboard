import json
import os
import sys

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "profile")
OUT = os.path.join(BASE_DIR, "fields.json")

SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)
LOGIN_URL = "https://wly04-sc.intergraphsmartcloud.com/Login.aspx"

FOCUS_VIEWS = [
    "vAssets_TestForms",
    "vTasks_TestsPlanned",
    "vPunchlist",
    "vInspectionNonCompliances",
    "vProcessBreakdownsSubsyst",
    "vProcessBreakdownsSyst",
    "vAssets",
]


def count_products(page):
    try:
        return page.evaluate(
            """() => (typeof switchboardConfiguration!=="undefined" && switchboardConfiguration.Products
            && switchboardConfiguration.Products.length) ? switchboardConfiguration.Products.length : -1"""
        )
    except Exception:
        return -1


def wait_products(page, timeout_sec):
    import time
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
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            PROFILE_DIR, channel="chrome", headless=True, viewport={"width": 1400, "height": 900}
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
        if not wait_products(page, timeout_sec=120):
            ctx.close()
            return 1

        out = {}
        for view in FOCUS_VIEWS:
            js = """(async () => {
                const r = {view: VIEWNAME, ok: false};
                try {
                    const ctl = SmartCompletions.data.getViewController('VIEWNAME');
                    if (!ctl) { r.err = 'no controller'; return r; }
                    const B = ctl.Behaviors || ctl.behaviors || {};
                    r.innerKeys = Object.keys(ctl);
                    r.behaviorKeys = Object.keys(B);
                    r.behaviors = {};
                    for (const k of Object.keys(B)) {
                        const val = B[k];
                        if (typeof val === 'function') { r.behaviors[k] = {__t:'fn'}; continue; }
                        try {
                            r.behaviors[k] = JSON.parse(JSON.stringify(val));
                        } catch(e) { r.behaviors[k] = {__t:'nonjson', k: Object.keys(val||{}).slice(0,50)}; }
                    }
                    r.ok = true;
                } catch(e) { r.err = String(e); }
                return r;
            })()""".replace("VIEWNAME", view)
            res = page.evaluate(js)
            out[view] = res
            print("dumped fields:", view)

        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        for view, res in out.items():
            bhv = res.get("behaviors", {})
            names = bhv.get("Names") or {}
            print("\n=== %s ===" % view)
            print("behaviorKeys:", res.get("behaviorKeys"))
            print("Names:", json.dumps(names, ensure_ascii=False)[:1500])
        print("\nsaved ->", OUT)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())