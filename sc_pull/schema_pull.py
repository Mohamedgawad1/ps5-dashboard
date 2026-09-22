import json
import os
import sys

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "profile")
JS_DIR = os.path.join(BASE_DIR, "js")
OUT = os.path.join(BASE_DIR, "schema.json")

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

JS_FILES = {
    "data": "/ISC/Scripts/SmartCompletions.Data3.js",
    "view_rendering": "/ISC/Scripts/SmartCompletions.viewPageRendering2.js",
    "utilities": "/ISC/Scripts/SmartCompletions.Utilities3.js",
    "reporting": "/ISC/Scripts/SmartCompletions.Reporting3.js",
    "appconstants": "/ISC/Scripts/AppConstants.js",
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


def dump_view_info(page, view):
    js = """((TOPVIEW) => {
            const r = {view: TOPVIEW};
            try {
                const ctl = SmartCompletions.data.getViewController(TOPVIEW);
                r.hasController = !!ctl;
                if (ctl) {
                    r.controllerKeys = Object.keys(ctl);
                    for (const k of Object.keys(ctl)) {
                        let val = ctl[k];
                        if (val === null || val === undefined) continue;
                        if (typeof val === 'object') {
                            try {
                                if (Array.isArray(val)) {
                                    r['ctl_'+k] = {__arrLen: val.length};
                                    if (val.length && typeof val[0] === 'object') {
                                        r['ctl_'+k].__firstKeys = Object.keys(val[0]).slice(0, 60);
                                        r['ctl_'+k].__first = (() => { const o = {}; for (const kk of Object.keys(val[0]).slice(0,15)) { const vv = val[0][kk]; if (typeof vv !== 'object') o[kk]=vv; } return o; })();
                                    }
                                } else if ((typeof val.maps !== 'undefined') || (typeof val.queries !== 'undefined')) {
                                    r['ctl_'+k] = {__kind: 'catalog', keys: Object.keys(val).slice(0,100)};
                                } else {
                                    r['ctl_'+k] = {__objKeys: Object.keys(val).slice(0,100)};
                                }
                            } catch(e) { r['ctl_'+k] = {__err: String(e)}; }
                        } else {
                            r['ctl_'+k] = val;
                        }
                    }
                }
            } catch(e) { r.controllerError = String(e); }
            return r;
        })(<<VIEW>>)""".replace("<<VIEW>>", json.dumps(view))
    return page.evaluate(js)


def main():
    os.makedirs(JS_DIR, exist_ok=True)
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            PROFILE_DIR,
            channel="chrome",
            headless=False,
            args=["--start-maximized"],
            viewport=None,
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
        if not wait_products(page, timeout_sec=25):
            print(">>> الجلسة مش موجودة؛ بفتح صفحة الدخول...")
            page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
            page.bring_to_front()

            creds = {}
            if os.path.exists(os.path.join(BASE_DIR, "creds.env")):
                with open(os.path.join(BASE_DIR, "creds.env"), encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if "=" in line and not line.startswith("#"):
                            k, v = line.split("=", 1)
                            creds[k.strip()] = v.strip().strip('"').strip("'")

            if creds.get("SC_USER") and creds.get("SC_PASS"):
                print(">>> هنملأ البيانات تلقائيًا (محاولة واحدة فقط)...")
                try:
                    page.fill("#txtUserName", creds["SC_USER"])
                    page.fill("#txtUserPass", creds["SC_PASS"])
                    page.click("#btnLogin", timeout=15000)
                except Exception as e:
                    print(">>> تعذر إدخال البيانات:", str(e)[:200])
            else:
                print("=" * 60)
                print(">>> اكتب اليوزر والباسورد واضغط LOGIN (بعد فتح النافذة)")
                print("=" * 60)

            if not wait_products(page, timeout_sec=180):
                print("=" * 60)
                print(">>> لو ظهرت صفحة الدخول (Sign In / Username) سجل بنفسك")
                print(">>> في صفحة التحقق/التأكيد (verification) كمّل زي ما بتعمل عادة")
                print(">>> وارجع للمنصة — هنستنى لحد ما تدخل.")
                print("=" * 60)
                if not wait_products(page, timeout_sec=900):
                    print(">>> انتهت المهلة بدون دخول.")
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
            out["views"][view] = dump_view_info(page, view)
            print("dumped controller:", view)

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
        print("JS saved:", {k: v.get("size", v) for k, v in js_src.items()})

        test_fields = {
            "vAssets_TestForms": ["ID", "AssetTag", "Description", "TaskCategoryID", "TaskSubCategoryID", "TaskStatus", "DateClosed", "ClosedID"],
            "vTasks_TestsPlanned": ["ID", "AssetTag", "Description", "TaskCategoryID", "TaskSubCategoryID", "TaskStatus", "DateClosed", "ClosedID"],
            "vPunchlist": ["ID", "AssetTag", "Description", "PunchItemNumber", "Category", "Status", "StatusID", "DateClosed"],
        }
        out["getTests"] = {}
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        for view, fields in test_fields.items():
            ejs = """((P) => {
                return new Promise((res) => {
                    dalc.Get(P.v, P.f, [], {Distinct:false, FirstRecordOrdinal:0, NumberOfRecords:3}, "Flat",
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
            })(%s)""" % json.dumps({"v": view, "f": fields})
            res = page.evaluate(ejs)
            out["getTests"][view] = res
            print(f"getTest[{view}]:", json.dumps(res, ensure_ascii=False)[:400])

        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        print("\nUSER:", out["user"])
        print("saved schema ->", OUT)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())