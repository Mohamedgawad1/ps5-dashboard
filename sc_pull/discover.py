import argparse
import json
import os
import sys

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "profile")
JS_DIR = os.path.join(BASE_DIR, "js")
OUT_JSON = os.path.join(BASE_DIR, "discovery.json")

LOGIN_URL = "https://wly04-sc.intergraphsmartcloud.com/Login.aspx"
SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)

JS_FILES = {
    "data": "/ISC/Scripts/SmartCompletions.Data3.js",
    "view_rendering": "/ISC/Scripts/SmartCompletions.viewPageRendering2.js",
    "utilities": "/ISC/Scripts/SmartCompletions.Utilities3.js",
    "reporting": "/ISC/Scripts/SmartCompletions.Reporting3.js",
}


def parse_args():
    p = argparse.ArgumentParser(description="Discover Smart Completions app list and data layer.")
    p.add_argument("--user", default=os.environ.get("SC_USER", ""), help="Username (or set SC_USER env)")
    p.add_argument("--pass", dest="password", default=os.environ.get("SC_PASS", ""), help="Password (or set SC_PASS env)")
    p.add_argument("--headless", action="store_true", help="Run headless (session must already exist)")
    return p.parse_args()


def wait_products(page, timeout_sec=900):
    import time
    start = time.time()
    while time.time() - start < timeout_sec:
        try:
            n = page.evaluate(
                """() => (typeof switchboardConfiguration!=="undefined" && switchboardConfiguration.Products
                && switchboardConfiguration.Products.length) ? switchboardConfiguration.Products.length : -1"""
            )
            logged = page.evaluate(
                """() => (typeof dalc!=="undefined" && dalc.user && dalc.user.loggedIn===true)"""
            )
            if n >= 0:
                return True
            url = page.url
            if "401" in url or "Login." in url:
                pass
        except Exception:
            pass
        try:
            page.wait_for_timeout(2000)
        except Exception:
            pass
    return False


def do_login(page, user, password):
    login_ok = wait_products(page, timeout_sec=20)
    if login_ok:
        print(">>> الجلسة موجودة بالفعل في البروفايل؛ متابعة بدون تسجيل دخول.")
        return

    print(">>> جارٍ فتح صفحة تسجيل الدخول...")
    try:
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
    except Exception:
        pass

    if user and password:
        try:
            page.fill("#txtUserName", user)
            page.fill("#txtUserPass", password)
            page.click("#btnLogin")
            print(">>> تم ملء البيانات؛ انتظرنا حتى يكتمل تسجيل الدخول (محاولة واحدة فقط).")
        except Exception as e:
            print(">>> تعذر إدخال البيانات تلقائيًا، سجل بنفسك:", e)
    else:
        print(">>> أدخل بياناتك في المتصفح المفتوح ثم اضغط LOGIN (سجل يدويًا حتى لا يقفل الحساب)...")

    if not wait_products(page, timeout_sec=900):
        raise RuntimeError("Login did not complete in time.")


def safe_js(page, expr, default=None):
    try:
        return page.evaluate(expr)
    except Exception as e:
        return {"__error__": str(e)} if default is None else default


def main():
    args = parse_args()
    os.makedirs(JS_DIR, exist_ok=True)

    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            PROFILE_DIR,
            channel="chrome",
            headless=args.headless,
            args=["--start-maximized"],
            viewport=None,
            accept_downloads=True,
        )
        page = context.pages[0] if context.pages else context.new_page()

        do_login(page, args.user, args.password)

        page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
        if not wait_products(page, timeout_sec=180):
            raise RuntimeError("Switchboard apps did not load.")

        apps = page.evaluate(
            """() => {
                const out = {modules: [], apps: []};
                if (typeof switchboardConfiguration !=="undefined" && switchboardConfiguration.Products) {
                    for (const mod of switchboardConfiguration.Products) {
                        const m = {label: mod.Label, groups: []};
                        for (const g of (mod.Groups || [])) {
                            const grp = {label: g.Label, actions: []};
                            for (const a of (g.Actions || [])) {
                                grp.actions.push({
                                    label: a.Label, view: a.View, applicationId: a.ApplicationID,
                                    icon: a.Icon, page: a.ApplicationPage || null,
                                });
                            }
                            m.groups.push(grp);
                        }
                        out.modules.push(m);
                        for (const g of (mod.Groups || [])) for (const a of (g.Actions || [])) out.apps.push(a.View);
                    }
                }
                out.user = (typeof dalc!=="undefined" && dalc.user) ? {
                    ID: dalc.user.ID, Resource: dalc.user.Resource,
                    CompanyID: dalc.user.CompanyID, loggedIn: dalc.user.loggedIn,
                } : null;
                return out;
            }"""
        )

        dalc_keys = safe_js(page, "Object.keys(dalc)")
        data_keys = safe_js(page, "Object.keys(SmartCompletions.data||{})")
        utilities_keys = safe_js(page, "Object.keys(SmartCompletions.utilities||{})")

        view_meta = {}
        view_probe = {}
        unique_apps = list(dict.fromkeys(apps.get("apps", [])))
        for view in unique_apps:
            meta = safe_js(
                page,
                """((v) => {
                    try {
                        const ctl = SmartCompletions.data.getViewController(v);
                        const out = {hasController: !!ctl};
                        if (ctl) {
                            out.controllerKeys = Object.keys(ctl);
                            try {
                                const jsonable = {};
                                for (const k of Object.keys(ctl)) {
                                    const val = ctl[k];
                                    if (val===null) jsonable[k]=null;
                                    else if (typeof val!=="object") jsonable[k]=val;
                                    else if (Array.isArray(val)) jsonable[k]=val.length;
                                    else jsonable[k]=Object.keys(val);
                                }
                                out.controllerShape = jsonable;
                            } catch(e) { out.shapeError = String(e); }
                        }
                        return out;
                    } catch(e) { return {error: String(e)}; }
                })""" % json.dumps(view),
            )
            view_meta[view] = view_meta.get(view) or meta

            probe = safe_js(
                page,
                """((v) => {
                    try {
                        return new Promise((res) => {
                            dalc.Get(v, [], [], {Distinct:false, FirstRecordOrdinal:0, NumberOfRecords:1}, "Flat",
                                (ok) => {
                                    let cols = [];
                                    let rows = 0;
                                    let total = null;
                                    if (ok && ok.Values && ok.Values.length) {
                                        cols = Object.keys(ok.Values[0]).filter(k=>!k.startsWith('__'));
                                        rows = ok.Values.length;
                                    }
                                    if (ok && ok.Request && ok.Request.TotalRecords) total = ok.Request.TotalRecords;
                                    res({ok:true, cols, rows, total, hasReq: !!ok});
                                },
                                (err) => res({ok:false, err: String((err && err.message) || err)}), false);
                        });
                    } catch(e) { return {ok:false, err:String(e)}; }
                })""" % json.dumps(view),
            )
            view_probe[view] = probe
            cols = probe.get("cols") if isinstance(probe, dict) else None
            print(f"   probe [{view}] ok={probe.get('ok')} rows={probe.get('rows')} total={probe.get('total')} cols={len(cols) if cols else 0}")

        js_src = {}
        for name, rel in JS_FILES.items():
            src = safe_js(
                page,
                "(async () => { try { const r = await fetch(%s, {headers: {'X-Auth-Token': localStorage.getItem('token') || '', 'Authorization': localStorage.getItem('token') || ''}}); if (!r.ok) return {status: r.status}; return await r.text(); } catch(e) { return {error: String(e)}; } })()"
                % json.dumps(rel),
            )
            if isinstance(src, str):
                with open(os.path.join(JS_DIR, name + ".js"), "w", encoding="utf-8") as f:
                    f.write(src)
                js_src[name] = {"saved": True, "size": len(src)}
            else:
                js_src[name] = src

        discovery = {
            "url": page.url,
            "user": apps.get("user"),
            "modules": apps.get("modules", []),
            "appCount": len(apps.get("apps", [])),
            "dalcKeys": dalc_keys,
            "smartCompletionsDataKeys": data_keys,
            "smartCompletionsUtilitiesKeys": utilities_keys,
            "viewMeta": view_meta,
            "jsSources": js_src,
        }
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(discovery, f, ensure_ascii=False, indent=2)

        print(f"\n=== DISCOVERY DONE ===\nUser: {discovery['user']}")
        print(f"Apps count: {discovery['appCount']}")
        for m in discovery["modules"]:
            print(f"\n[{m['label']}]")
            for g in m["groups"]:
                for a in g["actions"]:
                    print(f"   {a['label']}: view={a['view']}")
        print(f"\nSaved to {OUT_JSON}")

        page.wait_for_timeout(3000)


if __name__ == "__main__":
    sys.exit(main())