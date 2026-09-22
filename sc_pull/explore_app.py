import json
import os
import sys
import time

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
APP_URL = "https://wly04-sc.intergraphsmartcloud.com/ISC/tools/vTasks_TestsPlanned/index.htm"


def safe_eval(page, js, tries=4):
    last = None
    for i in range(tries):
        try:
            return page.evaluate(js)
        except Exception as e:
            last = e
            time.sleep(3)
    raise last


def main():
    b = None
    for attempt in range(5):
        try:
            b = sync_playwright().start().chromium.connect_over_cdp("http://127.0.0.1:9222")
            break
        except Exception as e:
            print("conn-retry:", str(e)[:100], flush=True)
            time.sleep(4)
    if not b:
        sys.exit("cannot connect")
    ctx = b.contexts[0]

    page = None
    for p in ctx.pages:
        try:
            t = safe_eval(p, "() => ({u: location.href, t: document.title})")
            print("tab:", json.dumps(t, ensure_ascii=False), flush=True)
            if "vTasks_TestsPlanned" in t["u"]:
                page = p
        except Exception:
            pass

    if page is None:
        page = ctx.new_page()
        try:
            page.goto(APP_URL, wait_until="commit", timeout=30000)
        except Exception as e:
            print("goto-note:", str(e)[:120], flush=True)
        page.wait_for_timeout(15000)

    for i in range(18):
        try:
            st = safe_eval(page, """() => ({u: location.href, t: document.title,
                dalc: (typeof dalc!=='undefined'), grids: document.querySelectorAll('.k-grid').length})""")
            print("state%d:" % i, json.dumps(st, ensure_ascii=False), flush=True)
            if st["u"].startswith("https://") and "index.htm" in st["u"]:
                break
        except Exception as e:
            print("err-check:", str(e)[:100], flush=True)
        page.wait_for_timeout(5000)

    page.wait_for_timeout(8000)
    info = safe_eval(page, """() => {
        const out = {};
        const txt = document.body.innerText;
        out.snippets = [];
        for (const key of ['Responsible', 'Summary', '1369', 'Planned Tasks', 'Export']) {
            const i = txt.indexOf(key);
            if (i >= 0) out.snippets.push({key: key, near: txt.slice(Math.max(0,i-160), i+220).replace(/\\s+/g,' ')});
        }
        out.grids = Array.from(document.querySelectorAll('.k-grid')).map(g => ({
            id: g.id, hdr: Array.from(g.querySelectorAll('th')).slice(0,50).map(x => (x.innerText||'').trim()).filter(Boolean),
            rows: g.querySelectorAll('tbody tr').length,
        }));
        out.reportLikes = Array.from(document.querySelectorAll('[class*="report"], [class*="Report"], [class*="list"], [class*="List"], [class*="tree"], [class*="Tree"]')).slice(0,40).map(x => ({tag:x.tagName, id:x.id, cls:(x.className||'').toString().slice(0,70), txt:(x.innerText||'').slice(0,220)}));
        return out;
    }""")
    print(json.dumps(info, ensure_ascii=False, indent=2), flush=True)
    with open(os.path.join(BASE_DIR, "app_dump.txt"), "w", encoding="utf-8") as f:
        f.write(safe_eval(page, "() => document.body.innerText"))
    print("[saved app_dump.txt]", flush=True)
    b.close()


if __name__ == "__main__":
    sys.exit(main())