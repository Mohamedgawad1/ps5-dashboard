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
        except Exception:
            pass
        if not wait_products(page, timeout_sec=40):
            print("NO SESSION", flush=True)
            page.bring_to_front()
            if not wait_products(page, timeout_sec=900):
                ctx.close()
                return 1

        info = page.evaluate(
            """() => {
                const out = {};
                const hdr = document.querySelector('#swithboard-header-right-panel');
                if (hdr) {
                    out.headerHtml = (hdr.innerText || '').slice(0, 200);
                    const dd = hdr.querySelector('.k-dropdown, .k-combobox, [role=combobox], select, .k-dropdowntree');
                    if (dd) {
                        out.dropdownTag = dd.tagName;
                        out.dropdownCls = (dd.className || '').toString().slice(0, 80);
                        out.dropdownText = (dd.innerText || dd.value || '').slice(0, 120);
                        out.dropdownId = dd.id;
                        if (dd.tagName === 'SELECT') {
                            out.selectOptions = Array.from(dd.options).map(o => ({text: o.text, value: o.value}));
                        }
                    }
                }
                out.localStorage = {pj: localStorage.getItem('ibspj'), name: localStorage.getItem('ibspjname'), ci: localStorage.getItem('ibsCIid')};
                return out;
            }"""
        )
        print(json.dumps(info, ensure_ascii=False, indent=2), flush=True)

        # click the header dropdown precisely
        click_info = page.evaluate(
            """() => {
                const hdr = document.querySelector('#swithboard-header-right-panel');
                if (!hdr) return {err: 'no header'};
                const dd = hdr.querySelector('.k-dropdown, .k-combobox, [role=combobox], select, .k-dropdowntree');
                if (!dd) return {err: 'no dropdown in header'};
                dd.click();
                return {clicked: true, tag: dd.tagName, id: dd.id};
            }"""
        )
        print("clicked:", json.dumps(click_info, ensure_ascii=False), flush=True)
        page.wait_for_timeout(2000)
        popup = page.evaluate(
            """() => {
                const out = {anns: [], lists: []};
                const anns = Array.from(document.querySelectorAll('.k-animation-container, .k-popup'));
                out.anns = anns.map(a => ({vis: a.style.display, off: a.offsetParent !== null, txt: (a.innerText || '').slice(0, 400)}));
                out.lists = Array.from(document.querySelectorAll('.k-list li, .k-listview-item, .k-item, [role=option]')).map(x => ({txt: (x.innerText||'').trim().slice(0,120), cls: (x.className||'').toString().slice(0,50)})).slice(0, 60);
                return out;
            }"""
        )
        print(json.dumps(popup, ensure_ascii=False, indent=2), flush=True)
        page.keyboard.press("Escape")
        with open(os.path.join(DATA_DIR, "projects_popup.json"), "w", encoding="utf-8") as f:
            json.dump({"switchboard": info, "clicked": click_info, "popup": popup}, f, ensure_ascii=False, indent=2)
        print("saved -> projects_popup.json", flush=True)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())