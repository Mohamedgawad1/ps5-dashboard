import json
import os
import sys

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "profile")
JS_DIR = os.path.join(BASE_DIR, "js")
OUT = os.path.join(BASE_DIR, "probe_session.json")

SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)

FOCUS_VIEWS = [
    "vAssets_TestForms",
    "vTasks_TestsPlanned",
    "vPunchlist",
    "vInspectionNonCompliances",
    "vProcessBreakdownsSubsyst",
    "vProcessBreakdownsSyst",
    "vAssets",
]

JS_FILES = {
    "data": "/ISC/Scripts/SmartCompletions.Data3.js",
    "view_rendering": "/ISC/Scripts/SmartCompletions.viewPageRendering2.js",
    "utilities": "/ISC/Scripts/SmartCompletions.Utilities3.js",
    "reporting": "/ISC/Scripts/SmartCompletions.Reporting3.js",
    "appconstants": "/ISC/Scripts/AppConstants.js",
}


def wait_products(page, timeout_sec=90):
    import time
    start = time.time()
    while time.time() - start < timeout_sec:
        try:
            n = page.evaluate(
                """() => (typeof switchboardConfiguration!=="undefined" && switchboardConfiguration.Products
                && switchboardConfiguration.Products.length) ? switchboardConfiguration.Products.length : -1"""
            )
            if n >= 0:
                return True
        except Exception:
            pass
        try:
            page.wait_for_timeout(2000)
        except Exception:
            pass
    return False


def main():
    os.makedirs(JS_DIR, exist_ok=True)
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            PROFILE_DIR, channel="chrome", headless=True, viewport={"width": 1400, "height": 900}
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
        if not wait_products(page, timeout_sec=90):
            print("NO SESSION / products not loaded")
            ctx.close()
            return 1

        out = {}
        out["user"] = page.evaluate(
            """() => (typeof dalc!=="undefined" && dalc.user) ? {ID: dalc.user.ID, Resource: dalc.user.Resource, loggedIn: dalc.user.loggedIn} : null"""
        )
        out["dalcKeys"] = page.evaluate("Object.keys(dalc)")
        out["dataKeys"] = page.evaluate("Object.keys(SmartCompletions.data||{})")

        out["views"] = {}
        for view in FOCUS_VIEWS:
            info = page.evaluate(
                """((v) => {
                    const r = {view: v};
                    try {
                        const ctl = SmartCompletions.data.getViewController(v);
                        r.hasController = !!ctl;
                        if (ctl) {
                            r.controllerKeys = Object.keys(ctl);
                            for (const k of ["fields","settings","columns","grids","primaryList","restrictions","options","sortings","filterpanel","toolbar"]) {
                                if (k in ctl) {
                                    const val = ctl[k];
                                    if (typeof val === "object" && val !== null) r["ctl_"+k+"_keys"] = Object.keys(val);
                                    else r["ctl_"+k] = val;
                                }
                            }
                        }
                    } catch(e) { r.controllerError = String(e); }
                    return r;
                })""" % json.dumps(view)
            )
            out["views"][view] = info

        js_src = {}
        for name, rel in JS_FILES.items():
            src = page.evaluate(
                "(async () => { try { const r = await fetch(%s, {headers: {'X-Auth-Token': localStorage.getItem('token') || '', 'Authorization': localStorage.getItem('token') || ''}}); if (!r.ok) return {status: r.status}; return await r.text(); } catch(e) { return {error: String(e)}; } })()"
                % json.dumps(rel)
            )
            if isinstance(src, str):
                with open(os.path.join(JS_DIR, name + ".js"), "w", encoding="utf-8") as f:
                    f.write(src)
                js_src[name] = {"saved": True, "size": len(src)}
            else:
                js_src[name] = src

        # try dalc.Get with a guessed small set of common fields for test views
        out["getTests"] = {}
        test_fields = {
            "vAssets_TestForms": ["ID", "AssetTag", "Description", "TaskCategoryID", "TaskSubCategoryID", "TaskStatus", "Closed", "ClosedDate", "DueDate"],
            "vTasks_TestsPlanned": ["ID", "AssetTag", "Description", "TaskCategoryID", "TaskSubCategoryID", "TaskStatus", "Closed", "ClosedDate", "DueDate"],
            "vPunchlist": ["ID", "AssetTag", "Description", "PunchItemNumber", "Category", "Status", "StatusID", "DateClosed"],
            "vInspectionNonCompliances": ["ID", "AssetTag", "Description", "Number", "Status", "Closed", "DateClosed"],
        }
        for view, fields in test_fields.items():
            ejs = """((PARAMS) => {
                return new Promise((res) => {
                    dalc.Get(PARAMS.v, PARAMS.f, [], {Distinct:false, FirstRecordOrdinal:0, NumberOfRecords:3}, "Flat",
                        (ok) => {
                            let cols = [], rows = 0, total = null, sample = null;
                            if (ok && ok.Values && ok.Values.length) {
                                cols = Object.keys(ok.Values[0]).filter(k=>!k.startsWith('__'));
                                rows = ok.Values.length;
                                sample = JSON.parse(JSON.stringify(ok.Values.slice(0,1)[0] || null));
                            }
                            if (ok && ok.Request) {
                                const tr = ok.Request.TotalRecords;
                                if (typeof tr !== "undefined" && tr !== null) total = tr;
                            }
                            res({ok:true, cols, rows, total, sample});
                        },
                        (err) => res({ok:false, err: String((err && err.message) || err)}), false);
                });
            })""".replace(
                "PARAMS", "ARGS"
            )
            ejs = ejs.replace(
                "(ARGS)", "(%s)" % json.dumps({"v": view, "f": fields})
            )
            res = page.evaluate(ejs)
            if isinstance(res, dict) and "cols" in res:
                res.pop("sample", None)
            out["getTests"][view] = res

        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

        print("USER:", out["user"])
        print("dalcKeys:", out["dalcKeys"])
        print("dataKeys:", out["dataKeys"])
        for v, info in out["views"].items():
            print(f"\n[{v}]", json.dumps(info, ensure_ascii=False))
        for v, res in out["getTests"].items():
            print(f"\n getTest[{v}]:", json.dumps(res, ensure_ascii=False)[:1500])
        print("\nSaved to", OUT)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())