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

        opened = page.evaluate(
            """() => {
                const dd = document.querySelector('.k-dropdown') || document.querySelector('.k-combobox');
                if (!dd) return {err: 'no dropdown'};
                dd.click();
                return {clicked: true};
            }"""
        )
        page.wait_for_timeout(1500)
        items = page.evaluate(
            """() => {
                const pops = Array.from(document.querySelectorAll('.k-popup, .k-animation-container')).filter(p => p.offsetParent !== null || p.style.display !== 'none');
                const texts = pops.map(p => (p.innerText || '').slice(0, 400));
                const lis = Array.from(document.querySelectorAll('.k-list .k-item, .k-listbox .k-item, .k-group .k-item, .k-item')).map(x => x.innerText).filter(t => t && t.trim());
                return {popups: texts, items: lis.slice(0, 50)};
            }"""
        )
        print(json.dumps(opened, ensure_ascii=False, indent=2))
        print(json.dumps(items, ensure_ascii=False, indent=2))
        # close popup by pressing escape to avoid UI side effects
        page.keyboard.press("Escape")
        b.close()


if __name__ == "__main__":
    sys.exit(main())