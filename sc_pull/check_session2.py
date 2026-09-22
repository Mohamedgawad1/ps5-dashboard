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
        PROFILE_DIR, channel="chrome", headless=True, viewport={"width": 1400, "height": 900}
    )
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=45000)
    page.wait_for_timeout(12000)
    state = page.evaluate(
        """() => {
            let out = {url: location.href, title: document.title};
            out.loginField = document.getElementById('txtUserName') !== null;
            out.token = (typeof localStorage!=="undefined") ? (localStorage.getItem('token') || null) : null;
            out.localStorageKeys = (typeof localStorage!=="undefined") ? Object.keys(localStorage).slice(0,40) : [];
            try {
                out.loggedIn = (typeof dalc!=="undefined" && dalc.user && dalc.user.loggedIn===true) || false;
                out.products = (typeof switchboardConfiguration!=='undefined' && switchboardConfiguration.Products)
                    ? switchboardConfiguration.Products.length : -1;
            } catch(e) { out.error = String(e); }
            return out;
        }"""
    )
    cookies = ctx.cookies("https://wly04-sc.intergraphsmartcloud.com")
    print(json.dumps({"state": state, "cookies": [{"n": c["name"], "v": (c["value"] or "")[:30], "e": c.get("expires")} for c in cookies]}, ensure_ascii=False, indent=2))
    ctx.close()