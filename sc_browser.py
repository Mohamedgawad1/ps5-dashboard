"""Attach to the already-open SmartCloud Chrome, so scripts never end the session.

The browser is started ONCE by Start_SmartCloud.bat (with --remote-debugging-port).
Every tool here connects to that window with CDP and disconnects when it is done.
Nothing in this file ever closes Chrome, so the SSO session survives between runs.
"""

import json
import os
import time
import urllib.error
import urllib.request

from playwright.sync_api import sync_playwright

PROFILE_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.getcwd()), "sc_itr_profile")
CDP_PORT = 9222
CDP_URL = "http://127.0.0.1:%d" % CDP_PORT
SWITCHBOARD = "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
BAT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Start_SmartCloud.bat")


def log(msg):
    print(msg, flush=True)


def cdp_alive(timeout=2.0):
    try:
        with urllib.request.urlopen(CDP_URL + "/json/version", timeout=timeout) as r:
            return json.loads(r.read().decode()).get("Browser", "")
    except Exception:
        return ""


def wait_for_cdp(seconds=90):
    end = time.time() + seconds
    while time.time() < end:
        v = cdp_alive()
        if v:
            return v
        time.sleep(2)
    return ""


def products_ready(page, seconds=45):
    end = time.time() + seconds
    while time.time() < end:
        try:
            n = page.evaluate(
                """() => (typeof switchboardConfiguration!=="undefined" && switchboardConfiguration.Products
                && switchboardConfiguration.Products.length) ? switchboardConfiguration.Products.length : -1"""
            )
            if n != -1:
                return True
        except Exception:
            pass
        time.sleep(2)
    return False


def session_page(ctx):
    """Return the tab that shows the switchboard, opening one if needed."""
    home = None
    for p in ctx.pages:
        try:
            if "vDashboardsUsers/Switchboard.htm" in (p.url or ""):
                home = p
                break
        except Exception:
            continue
    if home is None:
        home = ctx.new_page()
    try:
        home.bring_to_front()
    except Exception:
        pass
    if products_ready(home, 20):
        return home
    try:
        home.goto(SWITCHBOARD, wait_until="domcontentloaded", timeout=60000)
    except Exception as exc:
        log("goto note: %s" % str(exc)[:110])
    return home


def open_browser_window():
    """Start the standalone Chrome (only if it is not running yet)."""
    if cdp_alive():
        log("المتصفح شغال بالفعل")
        return True
    log("بنحل المتصفح من: %s" % BAT)
    os.startfile(BAT)
    v = wait_for_cdp(120)
    if v:
        log("المتصفح جاهز: %s" % v)
        return True
    log("!! المتصفح مفتحش - افتح الاختصار من الـ Desktop وجرب تاني")
    return False


def connect(pw, attempts=4):
    """Playwright object + browser context attached over CDP (Chrome keeps running)."""
    v = cdp_alive()
    if not v:
        v = wait_for_cdp(10)
    if not v:
        raise SystemExit("المتصفح مش شغال. دوس اختصار Start SmartCloud في الـ Desktop الأول.")
    last = None
    for i in range(1, attempts + 1):
        # a busy browser (many SmartCloud tabs) needs a long time to attach
        timeout = 45000 if i == 1 else 120000
        try:
            browser = pw.chromium.connect_over_cdp(CDP_URL, timeout=timeout)
            ctx = browser.contexts[0] if browser.contexts else browser.new_context()
            log("اتصلت بالمتصفح: %s" % v)
            return browser, ctx
        except Exception as exc:
            last = exc
            log("  connect attempt %d/%d failed (%ds): %s"
                % (i, attempts, timeout // 1000, str(exc).splitlines()[0][:100]))
            time.sleep(6)
    raise SystemExit(
        "مش قادر أوصل للمتصفح (%s). اقفل الـ tabs القديمة اللي فتحتها و حاول تاني."
        % str(last).splitlines()[0][:120]
    )


def keep_open(ctx, minutes=25):
    """Wait until you close the tabs, refreshing the session every few minutes.

    The SmartCloud session expires quickly when it sits idle, so while you work through
    the task pages this quietly reloads the switchboard tab. The browser is never closed.
    """
    log("\nخلاص. المتصفح لسه مفتوح.")
    log("أنا هعيد تحميل الصفحة كل %d دقيقة عشان الجلسة ما تقفلش." % minutes)
    log("لما تخلص، اقفل الـ tabs أو المتصفح وأنا كده.")
    end = time.time() + minutes * 60
    while True:
        time.sleep(20)
        if time.time() > end:
            log("انتهى وقت الـ keep-alive")
            return
        try:
            if not [p for p in ctx.pages if not p.is_closed()]:
                log("قفلت كل الـ tabs")
                return
        except Exception:
            return
        # refresh the switchboard tab roughly every 5 minutes
        try:
            n = sum(1 for p in ctx.pages if "vDashboardsUsers/Switchboard.htm" in (p.url or ""))
        except Exception:
            return
        now = time.time()
        if getattr(keep_open, "_last", 0) == 0:
            keep_open._last = now
        if now - keep_open._last > 300 and n:
            keep_open._last = now
            for p in ctx.pages:
                try:
                    if "vDashboardsUsers/Switchboard.htm" in (p.url or ""):
                        p.reload(wait_until="domcontentloaded", timeout=60000)
                        log("  ♻️ جدّدت الجلسة")
                        break
                except Exception:
                    pass
