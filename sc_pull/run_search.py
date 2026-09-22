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
APP_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/tools/vTasks_TestsPlanned/index.htm"
)

MAX_ROWS_TO_SAVE = 400000


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


def click_search(app):
    r = app.evaluate(
        """() => {
            let clicked = false;
            for (const el of Array.from(document.querySelectorAll('button, a, span, div, input'))) {
                const t = (el.innerText || el.value || '').trim();
                if (t === 'SEARCH' && (el.offsetParent !== null)) {
                    el.click();
                    clicked = true;
                    break;
                }
            }
            return clicked;
        }"""
    )
    return r


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
        if not wait_products(page, timeout_sec=30):
            print("NO SESSION", flush=True)
            page.bring_to_front()
            if not wait_products(page, timeout_sec=900):
                ctx.close()
                return 1

        app = ctx.new_page()
        try:
            app.goto(APP_URL, wait_until="commit", timeout=45000)
        except Exception as e:
            print("goto note:", str(e)[:100], flush=True)
        for i in range(24):
            app.wait_for_timeout(5000)
            try:
                t = app.evaluate(
                    """() => ({t: document.title, resp: document.body.innerText.indexOf("Responsible"), grids: document.querySelectorAll(".k-grid").length})"""
                )
                if t["resp"] >= 0 and t["grids"] > 0:
                    break
            except Exception:
                pass
        app.bring_to_front()

        clicked = click_search(app)
        print("search clicked:", clicked, flush=True)

        total = -1
        for i in range(40):
            app.wait_for_timeout(5000)
            try:
                total = app.evaluate(
                    """() => {
                        const g = document.querySelector("#vTasks_TestsPlanned_PrimaryList_grid");
                        if (!g) return -2;
                        try {
                            const kg = $(g).data('kendoGrid');
                            return kg ? kg.dataSource.total() : -3;
                        } catch (e) { return -4; }
                    }"""
                )
                print("poll%d total:" % i, total, flush=True)
                if total and total > 0:
                    break
            except Exception as e:
                print("poll err:", str(e)[:80], flush=True)

        if not total or total <= 0:
            print("no data after search", flush=True)
            ctx.close()
            return 1

        data = app.evaluate(
            """() => {
                const g = document.querySelector("#vTasks_TestsPlanned_PrimaryList_grid");
                const kg = $(g).data('kendoGrid');
                const ds = kg.dataSource;
                const all = ds.data();
                const rows = [];
                for (const r of all) {
                    const o = {};
                    for (const k of Object.keys(r)) o[k] = r[k];
                    rows.push(o);
                }
                let pageCount = null, pageSize = null;
                try { pageSize = ds.options.pageSize; } catch(e){}
                try { pageCount = ds.totalPages ? ds.totalPages() : null; } catch(e){}
                return {total: ds.total(), fetched: rows.length, pageSize, pageCount, rows: rows};
            }"""
        )
        print("fetched rows:", data["fetched"], "total:", data["total"], flush=True)
        with open(os.path.join(DATA_DIR, "tasks_raw.json"), "w", encoding="utf-8") as f:
            json.dump({"total": data["total"], "rows": data["rows"]}, f, ensure_ascii=False)
        print("saved data/tasks_raw.json (rows=%d)" % len(data["rows"]), flush=True)

        if data["rows"]:
            first = data["rows"][0]
            print("cols:", list(first.keys()), flush=True)
            for r in data["rows"][:5]:
                keep = {k: (str(v)[:80] if v is not None else None) for k, v in r.items()}
                print("row:", json.dumps(keep, ensure_ascii=False), flush=True)

        ctx.close()


if __name__ == "__main__":
    sys.exit(main())