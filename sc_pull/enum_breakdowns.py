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

GET_JS = """((P) => {
    return new Promise((res) => {
        dalc.Get(P.v, P.f, [], {Distinct:false, FirstRecordOrdinal:0, NumberOfRecords:P.n}, "Flat",
            (ok) => {
                let cols = [], rows = [], total = null;
                const vals = (ok && ok.Values) || [];
                if (vals.length) cols = Object.keys(vals[0]).filter(k=>!k.startsWith('__'));
                rows = JSON.parse(JSON.stringify(vals || []));
                if (ok && ok.Request) {
                    const tr = ok.Request.TotalRecords;
                    if (typeof tr !== "undefined" && tr !== null) total = tr;
                }
                res({cols: cols, rows: rows, total: total, n: rows.length});
            },
            (err) => res({err: String((err && err.message) || err)}), false);
    });
})"""

VIEWS = {
    "vProcessBreakdownsProjects": ["ID", "Summary", "ProcessBreakdownUp1Summary", "ProcessBreakdown", "ProcessBreakdownUp1ID", "ProjectID"],
    "vProcessBreakdownsSyst": ["ID", "Summary", "ProcessBreakdownUp1Summary", "ProcessBreakdown", "ProcessBreakdownUp1ID"],
    "vProcessBreakdownsSubsyst": ["ID", "Summary", "ProcessBreakdownUp1Summary", "ProcessBreakdown", "ProcessBreakdownUp1ID", "SystemID", "SubsystemID"],
    "vProcessBreakdowns": ["ID", "Summary", "ProcessBreakdownUp1Summary", "ProcessBreakdown", "ProcessBreakdownUp1ID"],
}


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


def get(page, view, fields, restrictions=None, n=2000, off=0):
    return page.evaluate("(%s)(%s)" % (
        GET_JS,
        json.dumps({"v": view, "f": fields, "r": restrictions or [], "n": n}),
    ))


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
        if not wait_products(page, timeout_sec=40):
            print("NO SESSION", flush=True)
            page.bring_to_front()
            if not wait_products(page, timeout_sec=900):
                ctx.close()
                return 1

        out = {}
        for view, fields in VIEWS.items():
            print("== %s ==" % view, flush=True)
            r = get(page, view, fields)
            cols = r.get("cols", [])
            rows = r.get("rows") or []
            print("   cols:", cols, "n:", r.get("n"), "total:", r.get("total"), "err:", r.get("err"), flush=True)
            out[view] = {"cols": cols, "rows": rows, "total": r.get("total")}
            if not rows:
                continue
            unique = {}
            for row in rows:
                key = str(row.get("ID")) + "|" + str(row.get("Summary"))
                unique[key] = row
            for row in unique.values():
                compact = {k: str(v)[:70] for k, v in row.items()}
                s = json.dumps(compact, ensure_ascii=False)
                print("   " + s[:220], flush=True)
                if any(x in s.upper() for x in ("CPP", "AGI", "PUMP", "PS5")):
                    out.setdefault("matches", []).append(compact)

        with open(os.path.join(DATA_DIR, "breakdowns.json"), "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        print("saved ->", os.path.join(DATA_DIR, "breakdowns.json"), flush=True)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())