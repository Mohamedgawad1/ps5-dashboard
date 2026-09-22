import sys
import os

from playwright.sync_api import sync_playwright

path_arg = sys.argv[1] if len(sys.argv) > 1 else os.path.abspath("index.html")
path = os.path.abspath(path_arg)
url = "file:///" + path.replace("\\", "/")

with sync_playwright() as pw:
    browser = pw.chromium.launch(channel="chrome", headless=True)
    page = browser.new_page(viewport={"width": 1600, "height": 1000})
    errors = []
    stacks = []

    def on_err(msg):
        errors.append(msg.text)
        try:
            loc = msg.location
            stacks.append("loc:%s:%s" % (loc.get("lineNumber"), loc.get("columnNumber")))
        except Exception:
            pass

    page.on("console", lambda msg: on_err(msg) if msg.type in ("error",) else None)

    def on_pageerr(e):
        stacks.append("PAGEERR: " + (str(getattr(e, "stack", "") or "")[-400:]))
        errors.append("PAGEERROR: " + str(e))

    page.on("pageerror", on_pageerr)
    sess = None
    excpos = []

    def on_exc(p):
        try:
            d = p.get("exceptionDetails") or {}
            excpos.append({
                "url": d.get("url"), "line": d.get("lineNumber"),
                "col": d.get("columnNumber"), "text": d.get("text"),
            })
        except Exception:
            pass
    try:
        sess = page.context.new_cdp_session(page)
        sess.on("Runtime.exceptionThrown", on_exc)
        sess.send("Runtime.enable")
    except Exception as e:
        print("cdp err:", e)
    page.goto(url, wait_until="load", timeout=120000)
    page.wait_for_timeout(15000)
    info = page.evaluate(
        """() => {
            const c = document.getElementById('chartCombinedDaily');
            const out = {exists: !!c};
            if (!c) return out;
            const r = c.getBoundingClientRect();
            out.width = r.width; out.height = r.height;
            const ctx = c.getContext('2d');
            out.pixels = ctx.getImageData(0,0,c.width,c.height).data.length;
            let colored = 0;
            const d = ctx.getImageData(0,0,c.width,c.height).data;
            for (let i=0;i<d.length;i+=4){ if(d[i]+ d[i+1]+d[i+2] < 750) colored++; }
            out.colored = colored;
            try { out.charts = Object.keys(Chart.instances||{}).length; } catch(e){ out.charts = -1; }
            try {
              const insts = Object.keys(Chart.instances||{}).map(k=>Chart.instances[k]);
              out.winITR = window.ITR ? {daily:(window.ITR.daily||[]).length, subj:(window.ITR.daily_submitted||[]).length} : null;
            } catch(e){ out.winITR = 'err:'+e; }
            try {
              const c = window.Chart ? Object.keys(window.Chart.instances||{}).map(k=>window.Chart.instances[k]).find(x=>x.canvas && x.canvas.id==='chartCombinedDaily') : null;
              out.combined = c ? {
                datasets: (c.data.datasets||[]).map(ds=>({label:ds.label, hasDL:!!ds.datalabels, total:(ds.data||[]).reduce((a,b)=>a+(b||0),0), head:(ds.data||[]).slice(0,6)}))
              } : 'no-inst';
            } catch(e){ out.combined = 'err:'+e; }
            return out;
        }"""
    )
    print("INFO:", info)
    print("CONSOLE errors:", errors[:8])
    print("STACKS:", stacks[:8])
    print("CDP EXC:", excpos)
    snap = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chart_check.png")
    try:
        el = page.query_selector("#chartCombinedDaily")
        if el:
            el.screenshot(path=os.path.join(os.path.dirname(os.path.abspath(__file__)), "chart_el.png"))
    except Exception as e:
        print("screenshot err:", e)
    page.screenshot(path=snap, full_page=False)
    print("screenshot saved")
    browser.close()