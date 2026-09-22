import json
import os
import sys
import time

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "profile")

SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)
APP_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/tools/vTasks_TestsPlanned/index.htm"
)


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
                print("st%d:" % i, json.dumps(t, ensure_ascii=False), flush=True)
                if t["resp"] >= 0 and t["grids"] > 0:
                    break
            except Exception:
                pass
        app.bring_to_front()

        # dump reports control candidates
        rep = app.evaluate(
            """() => {
                const out = {items: []};
                const seed = [];
                for (const el of Array.from(document.querySelectorAll('div, span, li, a, button, td, select'))) {
                    const t = (el.innerText || '').trim();
                    if (t === 'Reports' || /Responsible Company/i.test(t) || /Summary/i.test(t)) {
                        out.items.push({tag: el.tagName, id: el.id, cls: (el.className||'').toString().slice(0,80), text: t.slice(0,140)});
                    }
                }
                out.reportsClickables = Array.from(document.querySelectorAll('[class*=dropdown], [class*=combo], [class*=menu], [class*=popup], [id*=report], [id*=Report]')).map(x => ({
                    tag: x.tagName, id: x.id, cls: (x.className||'').toString().slice(0,80), txt: (x.innerText||'').slice(0,200),
                })).slice(0,25);
                return out;
            }"""
        )
        print("REPORT CANDIDATES:", json.dumps(rep, ensure_ascii=False, indent=2), flush=True)

        # inspect grids + their dataSources
        ginfo = app.evaluate(
            """() => {
                const out = [];
                for (const g of Array.from(document.querySelectorAll('.k-grid'))) {
                    const entry = {id: g.id, cls: (g.className||'').toString().slice(0,60)};
                    try {
                        const kg = $(g).data('kendoGrid');
                        if (kg) {
                            const ds = kg.dataSource;
                            entry.hasKendoGrid = true;
                            entry.pageSize = ds.options && ds.options.pageSize;
                            entry.total = ds.total && ds.total();
                            try { entry.serverPaging = ds.options.serverPaging; } catch(e){}
                            try {
                                const rd = ds.options.transport && ds.options.transport.read && ds.options.transport.read.data;
                                entry.readData = (typeof rd === 'function') ? String(rd).slice(0, 900) : rd;
                            } catch(e){}
                            try { entry.view = ds.options.data && ds.options.data.view; } catch(e){}
                            const url = ds.options.transport && ds.options.transport.read && ds.options.transport.read.url;
                            if (typeof url === 'string') entry.readUrl = url.slice(0, 300);
                        }
                    } catch(e) { entry.gridErr = String(e).slice(0,150); }
                    out.push(entry);
                }
                return out;
            }"""
        )
        print("GRID INFO:", json.dumps(ginfo, ensure_ascii=False, indent=2), flush=True)
        with open(os.path.join(BASE_DIR, "grid_info.json"), "w", encoding="utf-8") as f:
            json.dump({"reports": rep, "grids": ginfo}, f, ensure_ascii=False, indent=2)

        ctx.close()


if __name__ == "__main__":
    sys.exit(main())