import json
import os
import sys

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "profile")
CREDS_FILE = os.path.join(BASE_DIR, "creds.env")
OUT_JSON = os.path.join(BASE_DIR, "discovery.json")

LOGIN_URL = "https://wly04-sc.intergraphsmartcloud.com/Login.aspx"
SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)


def load_creds():
    creds = {}
    if os.path.exists(CREDS_FILE):
        with open(CREDS_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                creds[k.strip()] = v.strip().strip('"').strip("'")
    return creds, CREDS_FILE


def wait_products(page, timeout_sec=240):
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


def login_once(page, user, password):
    try:
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
    except Exception as e:
        print("nav-warning:", str(e)[:150])
    try:
        page.fill("#txtUserName", user)
        page.fill("#txtUserPass", password)
        page.click("#btnLogin", timeout=15000)
        print(">>> تم إرسال بيانات الدخول (محاولة واحدة فقط)...")
    except Exception as e:
        print(">>> تعذر إدخال البيانات:", str(e)[:200])
        return False

    import time
    for _ in range(20):
        still_login = page.evaluate(
            """() => document.getElementById('txtUserName') !== null"""
        )
        if not still_login:
            break
        page.wait_for_timeout(1000)
    page.wait_for_timeout(8000)
    still_login = page.evaluate(
        """() => document.getElementById('txtUserName') !== null"""
    )
    if still_login:
        print(">>> يبدو أن البيانات لم تُقبل (هنشوف أي رسالة):")
        print(page.evaluate("() => document.body.innerText.slice(0, 400)"))
        try:
            page.screenshot(path=os.path.join(BASE_DIR, "login_fail.png"))
        except Exception:
            pass
        return False
    return wait_products(page, timeout_sec=240)


def main():
    creds, creds_path = load_creds()
    do_login = bool(creds.get("SC_USER") and creds.get("SC_PASS"))

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

        if do_login:
            ok = login_once(page, creds["SC_USER"], creds["SC_PASS"])
            if not ok:
                print(">>> فشل تسجيل الدخول في المحاولة الأولى؛ بايفطر. هقفل من غير ما أعيد محاولة.")
                ctx.close()
                return 1
        else:
            page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
            if not wait_products(page, timeout_sec=240):
                print(">>> مش مسجل؛ هفتح صفحة الدخول، سجل يدويًا لأول مرة وبس (وقفلتنا أمان).")
                page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
                print(">>> سجل دخولك في النافذة المفتوحة...")
                if not wait_products(page, timeout_sec=900):
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

        print("\n>>> الجلسة محفوظة في البروفايل. بعد كده كل تشغيلة هتشتغل أوتوماتيك من غير دخول.")
        page.wait_for_timeout(4000)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())