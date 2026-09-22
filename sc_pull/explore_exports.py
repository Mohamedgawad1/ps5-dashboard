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

EXPORTS_URL = "https://wly04-sc.intergraphsmartcloud.com/ISC/tools/vExports/index.htm"
EDIT_URL = "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/EditForm.aspx?ID=1369&v=vExports"


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


def safe_eval(page, js):
    try:
        return page.evaluate(js)
    except Exception as e:
        return {"evalErr": str(e)[:150]}


def open_and_wait(page, url, ready_js, timeout_sec=150):
    try:
        page.goto(url, wait_until="commit", timeout=45000)
    except Exception as e:
        print("  goto note:", str(e)[:100], flush=True)
    start = time.time()
    while time.time() - start < timeout_sec:
        page.wait_for_timeout(5000)
        try:
            st = safe_eval(page, ready_js)
            if st and (st.get("bodyLen") or 0) > 100:
                return st
        except Exception:
            pass
    return None


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

        # 1) Exports app index
        tab = ctx.new_page()
        print("opening EXPORTS app...", flush=True)
        st = open_and_wait(
            tab, EXPORTS_URL,
            """() => ({bodyLen: document.body.innerText.length, t: document.title,
                grids: document.querySelectorAll('.k-grid').length})""",
            timeout_sec=150,
        )
        print("exports state:", json.dumps(st, ensure_ascii=False), flush=True)
        tab.bring_to_front()
        tab.wait_for_timeout(8000)
        ex = safe_eval(tab, """() => {
            const txt = document.body.innerText;
            const out = {url: location.href, title: document.title};
            out.snippets = [];
            for (const key of ['Responsible', '1369', 'Export', 'Summary', 'Refresh', 'Row', 'Status']) {
                let i = txt.indexOf(key);
                let c = 0;
                while (i >= 0 && c < 3) {
                    out.snippets.push({key, near: txt.slice(Math.max(0,i-120), i+200).replace(/\\s+/g,' ')});
                    i = txt.indexOf(key, i + 200);
                    c++;
                }
            }
            out.grids = Array.from(document.querySelectorAll('.k-grid')).map(g => ({
                id: g.id,
                cols: Array.from(g.querySelectorAll('th')).map(x => (x.innerText||'').trim()).filter(Boolean).slice(0,40),
                rows: g.querySelectorAll('tbody tr').length,
            }));
            return out;
        }""")
        print("EXPORTS:", json.dumps(ex, ensure_ascii=False, indent=2), flush=True)
        with open(os.path.join(BASE_DIR, "exports_dump.txt"), "w", encoding="utf-8") as f:
            f.write(tab.evaluate("() => document.body.innerText"))

        # 2) EditForm.aspx?id=1369
        print("opening EDIT FORM 1369...", flush=True)
        tab2 = ctx.new_page()
        st2 = open_and_wait(
            tab2, EDIT_URL,
            """() => ({bodyLen: document.body.innerText.length, t: document.title})""",
            timeout_sec=180,
        )
        print("edit state:", json.dumps(st2, ensure_ascii=False), flush=True)
        tab2.bring_to_front()
        tab2.wait_for_timeout(10000)
        ef = safe_eval(tab2, """() => {
            const txt = document.body.innerText;
            return {
                url: location.href,
                title: document.title,
                len: txt.length,
                text: txt.slice(0, 6000),
            };
        }""")
        print("EDIT1369 head:", json.dumps({k: (v[:4000] if k == 'text' else v) for k, v in ef.items()}, ensure_ascii=False), flush=True)
        with open(os.path.join(BASE_DIR, "edit1369_dump.txt"), "w", encoding="utf-8") as f:
            f.write(tab2.evaluate("() => document.body.innerText"))
        print("[saved exports_dump.txt & edit1369_dump.txt]", flush=True)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())