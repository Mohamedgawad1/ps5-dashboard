import json
import os
import sys

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "profile")
OUT_JSON = os.path.join(BASE_DIR, "discovery.json")

LOGIN_URL = "https://wly04-sc.intergraphsmartcloud.com/Login.aspx"
SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)


def count_products(page):
    try:
        return page.evaluate(
            """() => (typeof switchboardConfiguration!=="undefined" && switchboardConfiguration.Products
            && switchboardConfiguration.Products.length) ? switchboardConfiguration.Products.length : -1"""
        )
    except Exception:
        return -1


def wait_products(page, timeout_sec=900):
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
            PROFILE_DIR,
            channel="chrome",
            headless=False,
            args=["--start-maximized"],
            viewport=None,
            accept_downloads=True,
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        try:
            page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
        except Exception as e:
            print("nav-warning:", str(e)[:150])

        if not wait_products(page, timeout_sec=60):
            print(">>> الجلسة مش موجودة؛ هنفتح صفحة الدخول...")
            try:
                page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
            except Exception as e:
                print("nav-warning:", str(e)[:150])
            page.bring_to_front()
            print("=" * 60)
            print(">>> النافذة مفتوحة على صفحة الدخول. برجاء الآن:")
            print("    1) اضغط زر: Login with Single Sign-On (SSO)")
            print("    2) سجل الدخول بحساب Google (smartplantcloud) مثل ما بتعمل عادة")
            print("    3) لما ترجع للمنصة والأزرار تظهر، خلاص (السكريبت يكمّل لوحده)")
            print(">>> مفيش أي إعادة محاولة تلقائية.")
            print("=" * 60)
            if not wait_products(page, timeout_sec=900):
                print(">>> لم يتم العثور على أزرار المنصة خلال المهلة.")
                ctx.close()
                return 1

        apps = page.evaluate(
            """() => {
                const out = {modules: [], apps: []};
                if (typeof switchboardConfiguration !=="undefined" && switchboardConfiguration.Products) {
                    for (const mod of switchboardConfiguration.Products) {
                        const m = {label: mod.Label, groups: []};
                        for (const g of (mod.Groups || [])) {
                            const grp = {label: g.Label, actions: []};
                            for (const a of (g.Actions || [])) {
                                grp.actions.push({label: a.Label, view: a.View, applicationId: a.ApplicationID,
                                                  icon: a.Icon, page: a.ApplicationPage || null});
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

        view_probe = {}
        for view in list(dict.fromkeys(apps.get("apps", []))):
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
            probe = page.evaluate(probe_js)
            view_probe[view] = probe
            print(
                "   probe [%s] ok=%s rows=%s total=%s cols=%d"
                % (view, probe.get("ok"), probe.get("rows"), probe.get("total"), len(probe.get("cols") or []))
            )

        discovery = {
            "url": page.url,
            "user": apps.get("user"),
            "modules": apps.get("modules", []),
            "appCount": len(apps.get("apps", [])),
            "viewProbe": view_probe,
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
        print("\n>>> الجلسة محفوظة. التشغيلات الجاية هتشتغل تلقائيًا.")

        page.wait_for_timeout(4000)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())