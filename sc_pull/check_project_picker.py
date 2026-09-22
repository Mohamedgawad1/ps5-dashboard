import json
import sys

from playwright.sync_api import sync_playwright


def main():
    with sync_playwright() as pw:
        b = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
        ctx = b.contexts[0]
        page = None
        for p in ctx.pages:
            try:
                n = p.evaluate(
                    """() => (typeof switchboardConfiguration!=='undefined'&&switchboardConfiguration.Products)?switchboardConfiguration.Products.length:-1"""
                )
                if n >= 0:
                    page = p
                    break
            except Exception:
                continue
        if not page:
            print("no switchboard page")
            b.close()
            return

        info = page.evaluate(
            """() => {
                const out = {};
                out.selects = Array.from(document.querySelectorAll('select')).map(s => ({id: s.id, name: s.name, options: Array.from(s.options).slice(0,8).map(o => o.text)}));
                out.combos = Array.from(document.querySelectorAll('.k-combobox, .k-dropdown, [role=combobox], .k-dropdowntree')).slice(0,10).map(x => ({cls: x.className.toString().slice(0,60), html: (x.innerText||'').slice(0,80)}));
                out.eacopInBody = (document.body.innerText.indexOf('EACOP') >= 0);
                const txt = document.body.innerText;
                const idx = txt.indexOf('EACOP');
                out.around = idx >= 0 ? txt.slice(Math.max(0, idx-120), idx+120) : null;
                out.toolbars = Array.from(document.querySelectorAll('[id*="hdrtbl"], [id*="header"], [id*="search"], [class*="dropdown"]')).slice(0,15).map(x => ({id: x.id, cls: (x.className||'').toString().slice(0,60), txt: (x.innerText||'').slice(0,100)}));
                return out;
            }"""
        )
        print(json.dumps(info, ensure_ascii=False, indent=2))
        b.close()


if __name__ == "__main__":
    sys.exit(main())