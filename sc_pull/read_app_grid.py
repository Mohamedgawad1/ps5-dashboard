import json
import os
import sys
import time

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def main():
    b = None
    try:
        b = sync_playwright().start().chromium.connect_over_cdp("http://127.0.0.1:9222")
    except Exception as e:
        sys.exit("conn failed: %s" % str(e)[:120])
    ctx = b.contexts[0]
    page = None
    for p in ctx.pages:
        try:
            if "vAssets_TestForms" in p.url:
                page = p
                break
        except Exception:
            continue
    if not page:
        print("no itr tab open")
        b.close()
        return

    try:
        page.bring_to_front()
    except Exception:
        pass
    page.wait_for_timeout(4000)

    info = page.evaluate(
        """() => {
            const out = {};
            out.url = location.href;
            const grids = Array.from(document.querySelectorAll('.k-grid'));
            out.gridCount = grids.length;
            out.grids = grids.map(g => ({
                id: g.id, cls: (g.className||'').toString().slice(0,70),
                cols: Array.from(g.querySelectorAll('th')).map(x => (x.innerText||'').trim()),
                rows: g.querySelectorAll('tbody tr').length,
            }));
            out.pagers = Array.from(document.querySelectorAll('.k-pager-info, .k-pager-sizes, .k-pager-refresh, .k-pager-numeric, [class*=pager]')).map(x => (x.innerText||'').trim()).filter(t=>t).slice(0,20);
            const txt = document.body.innerText;
            out.hasStatus = txt.indexOf('Status') >= 0;
            out.hasClosed = txt.indexOf('Closed') >= 0;
            out.hasPunchlistCol = txt.indexOf('Punch') >= 0;
            const li = txt.indexOf('Status');
            out.statusNear = li>=0 ? txt.slice(Math.max(0,li-80), li+160).replace(/\\s+/g,' ') : null;
            return out;
        }"""
    )
    print(json.dumps(info, ensure_ascii=False, indent=2), flush=True)
    with open(os.path.join(BASE_DIR, "itr_grid_dump.txt"), "w", encoding="utf-8") as f:
        f.write(page.evaluate("() => document.body.innerText"))
    print("[saved itr_grid_dump.txt]", flush=True)
    b.close()


if __name__ == "__main__":
    sys.exit(main())