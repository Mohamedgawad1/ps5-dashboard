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

APP_URLS = [
    "https://wly04-sc.intergraphsmartcloud.com/ISC/tools/vTasks_TestsPlanned/index.htm",
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/Tasks/vTasks_TestsPlanned/index.htm",
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/tasks/vTasks_TestsPlanned/index.htm",
    "https://wly04-sc.intergraphsmartcloud.com/ISC/tools/tasks/vTasks_TestsPlanned/index.htm",
]

REPORT_KEY = "Responsible Company"


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


def try_json(page):
    return page.evaluate(
        """() => {
            const out = {url: location.href, title: document.title};
            const txt = document.body.innerText;
            out.hasResponsible = txt.indexOf("Responsible") >= 0;
            out.hasDalc = typeof dalc !== "undefined";
            out.grids = Array.from(document.querySelectorAll(".k-grid")).length;
            out.bodyLen = txt.length;
            return out;
        }"""
    )


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
        except Exception as e:
            print("nav:", str(e)[:120], flush=True)
        if not wait_products(page, timeout_sec=30):
            print("NO SESSION - need login in opened window", flush=True)
            page.bring_to_front()
            if not wait_products(page, timeout_sec=900):
                print("login timeout", flush=True)
                ctx.close()
                return 1

        app_page = ctx.new_page()
        loaded = False
        for url in APP_URLS:
            print("trying:", url, flush=True)
            try:
                app_page.goto(url, wait_until="commit", timeout=45000)
            except Exception as e:
                print("  goto note:", str(e)[:100], flush=True)
            ok = False
            for i in range(20):
                app_page.wait_for_timeout(5000)
                try:
                    st = try_json(app_page)
                except Exception as e:
                    print("  eval err:", str(e)[:80], flush=True)
                    st = None
                if st:
                    print("  state%d:" % i, json.dumps(st, ensure_ascii=False), flush=True)
                    good_title = any(k in (st.get("title") or "").lower() for k in ("task", "plan", "inspection"))
                    nav_ok = "index.htm" in st["url"] and st["url"].startswith("https")
                    if nav_ok and st["hasDalc"] and not st["url"].endswith("Login.htm"):
                        ok = True
                        break
            if ok:
                loaded = True
                break

        if not loaded:
            print("app page did not load well; continuing with dump anyway", flush=True)

        app_page.bring_to_front()
        app_page.wait_for_timeout(6000)

        info = app_page.evaluate(
            """() => {
                const out = {};
                const txt = document.body.innerText;
                out.url = location.href;
                const i = txt.indexOf("Responsible");
                if (i >= 0) out.responsibleNear = txt.slice(Math.max(0,i-200), i+300).replace(/\\s+/g,' ');
                out.reportItems = [];
                const els = Array.from(document.querySelectorAll("span, div, li, a, td"));
                for (const el of els) {
                    const t = (el.innerText || "").trim();
                    if (t && t.length < 120 && /Responsible|Summary/i.test(t)) {
                        out.reportItems.push({tag: el.tagName, id: el.id, cls: (el.className||"").toString().slice(0,60), text: t.slice(0,120)});
                    }
                }
                out.reportItems = out.reportItems.slice(0, 30);
                out.grids = Array.from(document.querySelectorAll(".k-grid")).map(g => ({
                    id: g.id, cls: (g.className||"").toString().slice(0,60),
                    cols: Array.from(g.querySelectorAll("th")).map(x => (x.innerText||"").trim()).filter(Boolean).slice(0,60),
                    rows: g.querySelectorAll("tbody tr").length,
                    pager: (() => { const p = g.querySelector("[class*=pager-info], [class*=pager]"); return p ? (p.innerText||"").trim() : null; })(),
                }));
                return out;
            }"""
        )
        print(json.dumps(info, ensure_ascii=False, indent=2), flush=True)

        with open(os.path.join(BASE_DIR, "app_dump2.txt"), "w", encoding="utf-8") as f:
            f.write(app_page.evaluate("() => document.body.innerText"))
        print("[saved app_dump2.txt]", flush=True)

        click_target = None
        for item in info.get("reportItems", []):
            if "Responsible" in item["text"] or "Summary" in item["text"]:
                click_target = item
                break
        if click_target:
            print("clicking:", json.dumps(click_target, ensure_ascii=False), flush=True)
            try:
                app_page.evaluate(
                    """(args) => {
                        const els = Array.from(document.querySelectorAll("span, div, li, a, td"));
                        for (const el of els) {
                            const t = (el.innerText || "").trim();
                            if (t === args.text && (el.id === args.id || !args.id)) { el.click(); return {clicked: true, tag: el.tagName, id: el.id}; }
                        }
                        return {clicked: false};
                    }""",
                    click_target,
                )
            except Exception as e:
                print("click err:", str(e)[:120], flush=True)

        app_page.wait_for_timeout(15000)
        after = app_page.evaluate(
            """() => {
                const out = {};
                const txt = document.body.innerText;
                const i = txt.indexOf("Responsible");
                if (i >= 0) out.responsibleNear = txt.slice(Math.max(0,i-100), i+500).replace(/\\s+/g,' ');
                out.grids = Array.from(document.querySelectorAll(".k-grid")).map(g => ({
                    id: g.id, cols: Array.from(g.querySelectorAll("th")).map(x => (x.innerText||"").trim()).filter(Boolean).slice(0,80),
                    rows: g.querySelectorAll("tbody tr").length,
                    pager: (() => { const p = g.querySelector("[class*=pager]"); return p ? (p.innerText||"").trim() : null; })(),
                }));
                return out;
            }"""
        )
        print(json.dumps({"after": after}, ensure_ascii=False, indent=2), flush=True)
        with open(os.path.join(BASE_DIR, "app_dump2_after.txt"), "w", encoding="utf-8") as f:
            f.write(app_page.evaluate("() => document.body.innerText"))
        print("[saved app_dump2_after.txt]", flush=True)

        ctx.close()


if __name__ == "__main__":
    sys.exit(main())