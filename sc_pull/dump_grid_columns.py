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
        ok = wait_products(page, timeout_sec=60)
        print("session:", ok, flush=True)
        if not ok:
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
            app.wait_for_timeout(6000)
            try:
                t = app.evaluate(
                    """() => ({t: document.title, grids: document.querySelectorAll(".k-grid").length})"""
                )
                if t["grids"] >= 2:
                    break
            except Exception:
                pass
        app.bring_to_front()

        info = app.evaluate(
            """() => {
                const g = document.querySelector("#vTasks_TestsPlanned_PrimaryList_grid");
                const kg = $(g).data("kendoGrid");
                const out = {};
                try {
                    out.columns = (kg.options.columns || []).map(c => {
                        const o = {};
                        for (const k of Object.keys(c)) {
                            let v = c[k];
                            if (typeof v === 'function') v = '[' + String(v).slice(0,120) + ']';
                            else if (typeof v === 'object' && v !== null) { try { v = JSON.stringify(v).slice(0,200); } catch(e){} }
                            o[k] = v;
                        }
                        return o;
                    });
                } catch(e) { out.columnsErr = String(e).slice(0,300); }
                const ds = kg.dataSource;
                try {
                    const opts = ds.options;
                    out.dsKeys = Object.keys(opts);
                    out.pageSize = opts.pageSize;
                    out.serverPaging = opts.serverPaging;
                    const t = opts.transport || {};
                    out.transportKeys = Object.keys(t);
                    for (const k of Object.keys(t)) {
                        let v = t[k];
                        if (typeof v === 'function') v = '[' + String(v).slice(0,2000) + ']';
                        else if (typeof v === 'object') { try { v = JSON.stringify(v).slice(0,2000); } catch(e){} }
                        out["transport_" + k] = v;
                    }
                    try { out.sort = JSON.stringify(opts.sort); } catch(e){}
                    try { out.filter = JSON.stringify(opts.filter); } catch(e){}
                    try { out.schema = JSON.stringify(opts.schema); } catch(e){}
                } catch(e) { out.dsErr = String(e).slice(0,300); }
                return out;
            }"""
        )
        print(json.dumps(info, ensure_ascii=False, indent=2), flush=True)
        with open(os.path.join(DATA_DIR, "grid_columns.json"), "w", encoding="utf-8") as f:
            json.dump(info, f, ensure_ascii=False, indent=2)
        print("[saved data/grid_columns.json]", flush=True)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())