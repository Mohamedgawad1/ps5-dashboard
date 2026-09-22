import json
import os

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "profile")
SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        PROFILE_DIR,
        channel="chrome",
        headless=True,
        args=["--start-maximized"],
        viewport=None,
        accept_downloads=True,
    )
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    try:
        page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=45000)
    except Exception as e:
        print("nav-warning:", str(e)[:200])
        try:
            page.wait_for_timeout(8000)
        except Exception:
            pass
    page.wait_for_timeout(20000)
    state = page.evaluate(
        """() => {
            let out = {url: location.href, title: document.title};
            out.loginPage = document.getElementById('txtUserName') !== null;
            try {
                out.loggedIn = (typeof dalc!=='undefined' && dalc.user && dalc.user.loggedIn===true) || false;
                out.products = (typeof switchboardConfiguration!=='undefined' && switchboardConfiguration.Products)
                    ? switchboardConfiguration.Products.length : -1;
                out.bodySnippet = document.body.innerText.slice(0, 300);
            } catch(e) { out.error = String(e); }
            return out;
        }"""
    )
    print(json.dumps(state, ensure_ascii=False, indent=2))
    ctx.close()