import json
import os
import sys

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JS_DIR = os.path.join(BASE_DIR, "js")
OUT_JSON = os.path.join(BASE_DIR, "discovery.json")

CDP_URL = "http://127.0.0.1:9222"
SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)

JS_FILES = {
    "data": "/ISC/Scripts/SmartCompletions.Data3.js",
    "view_rendering": "/ISC/Scripts/SmartCompletions.viewPageRendering2.js",
    "utilities": "/ISC/Scripts/SmartCompletions.Utilities3.js",
    "reporting": "/ISC/Scripts/SmartCompletions.Reporting3.js",
}


def safe_js(page, expr, default=None):
    try:
        return page.evaluate(expr)
    except Exception as e:
        return {"__error__": str(e)} if default is None else default


def wait_products(page, timeout_sec=180):
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
        browser = pw.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]

        page = None
        for p in ctx.pages:
            if "Switchboard" in p.url or "intergraphsmartcloud" in p.url:
                page = p
                break
        if page is None:
            for p in ctx.pages:
                if p.url in ("", "about:blank", "chrome://newtab/"):
                    page = p
                    break
        if page is None:
            page = ctx.new_page()

        try:
            page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
        except Exception as e:
            print("nav-warning:", str(e)[:200])

        if not wait_products(page, timeout_sec=120):
            state = safe_js(
                page,
                """() => ({url: location.href, title: document.title,
                    loggedIn: (typeof dalc!=="undefined" && dalc.user && dalc.user.loggedIn===true) || false,
                    loginField: document.getElementById('txtUserName') !== null})""",
            )
            print("Switchboard apps did not load; page state:", json.dumps(state, ensure_ascii=False))
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

view_prob = {}
        unique_apps = list(dict.fromkeys(apps.get("apps", [])))
        for view in unique_apps:
            probe_js = """((TOPVIEW) => {
                try {
                    return new Promise((res) => {
                        dalc.Get(TOPVIEW, [], [], {Distinct:false, FirstRecordOrdinal:0, NumberOfRecords:1}, "Flat",
                            (ok) => {
                                let cols = []; let rows = 0; let total = null;
                                if (ok && ok.Values && ok.Values.length) {
                                    cols = Object.keys(ok.Values[0]).filter(k=>!k.startsWith('__'));
                                    rows = ok.Values.length;
                                }
                                if (ok && ok.Request) {
                                    const tr = ok.Request.TotalRecords;
                                    if (typeof tr !== "undefined" && tr !== null) total = tr;
                                }
                                res({ok:true, cols, rows, total});
                            },
                            (err) => res({ok:false, err: String((err && err.message) || err)}), false);
                    });
                } catch(e) { return Promise.resolve({ok:false, err:String(e)}); }
            })(<<VIEW>>)""".replace("<<VIEW>>", json.dumps(view))
            probe = safe_js(page, probe_js)
            view_prob[view] = probe
            print(
                "   probe [%s] ok=%s rows=%s total=%s cols=%d"
                % (
                    view,
                    probe.get("ok"),
                    probe.get("rows"),
                    probe.get("total"),
                    len(probe.get("cols") or []),
                )
            )

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
            "viewProbe": view_prob,
            "jsSources": js_src,
        }
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(discovery, f, ensure_ascii=False, indent=2)

        print("\n=== DISCOVERY DONE ===")
        print("User:", discovery["user"])
        print("Apps count:", discovery["appCount"])
        for m in discovery["modules"]:
            print("\n[%s]" % m["label"])
            for g in m["groups"]:
                for a in g["actions"]:
                    print("   %s: view=%s" % (a["label"], a["view"]))
        print("\nSaved to", OUT_JSON)

        browser.close()


if __name__ == "__main__":
    sys.exit(main())