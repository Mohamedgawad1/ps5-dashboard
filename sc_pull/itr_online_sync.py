import json
import os
import shutil
import subprocess
import sys
import time
from collections import Counter
import datetime

from playwright.sync_api import sync_playwright

import traceback

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WS = os.path.dirname(BASE_DIR)
DATA_DIR = os.path.join(BASE_DIR, "data")
# Sync uses its own dedicated profile. (Adopting the user's manual sc_itr_profile
# was tried but it fights the user's personal window + is not reliably logged in.)
PROFILE_DIR = os.path.join(BASE_DIR, "profile")
# Single deploy repo = the workspace (the one behind update_dashboard + github.io/ps5-dashboard)
CLONE = WS
# Live ITR page repo (the PUBLIC platform site). The sync pushes itr_live files here too
# so the live page always shows fresh data.
PAGE_REPO = "PS5-COMPLETION-PLATFORM"
PAGE_CLONE = os.path.join(os.environ.get("TEMP", BASE_DIR), "ps5_pages_deploy")
TOK_FILE = os.path.join(WS, "github_token.txt")
LOG = os.path.join(BASE_DIR, "itr_online_log.txt")

MILESTONE_MAP = {}
_MS_LOADED = False


def load_milestone_map():
    global MILESTONE_MAP, _MS_LOADED
    if _MS_LOADED:
        return
    p = os.path.join(DATA_DIR, "milestone_map.json")
    try:
        with open(p, "r", encoding="utf-8") as f:
            MILESTONE_MAP = {str(k).strip(): str(v).strip() for k, v in json.load(f).items()}
        pr("milestone_map: %d subsystems" % len(MILESTONE_MAP))
    except Exception as e:
        pr("milestone_map load err:", str(e)[:120])
    _MS_LOADED = True


def milestone_of(process_breakdown):
    if not process_breakdown:
        return None
    s = str(process_breakdown).strip()
    raw = MILESTONE_MAP.get(s)
    if raw:
        return raw
    for k, v in MILESTONE_MAP.items():
        if s.startswith(k) or k in s:
            return v
    return None

SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)
# URL marker for the sync's own dedicated pull tab inside the user's browser.
# Reusing our tab keeps pull-login sticky without touching any work tab.
MARK_URL = SWITCHBOARD_URL + "#itr-sync"
CPP_CID = 1021
PAGE = 4000
# Keep the browser session ALIVE after each update (reuse via debug port next run).
# Set ITR_KEEP_OPEN=0 to restore old close-after-pull behavior.
# 9222 is the user's SmartCloud browser (the one they sign in, close ITRs, and work in).
# The sync REUSES that one browser in its own dedicated tab - it never launches a second
# window, never closes the browser, and never touches the user's other tabs.
DEBUG_PORT = int(os.environ.get("ITR_DEBUG_PORT", "9222"))
KEEP_OPEN = os.environ.get("ITR_KEEP_OPEN", "1") != "0"
# Loop interval between pulls. Default 5 minutes; override with ITR_LOOP_SECONDS.
LOOP_SECONDS = int(os.environ.get("ITR_LOOP_SECONDS", "300"))

FIELDS = [
    "ID", "TaskName", "AssetTag", "AssetDescription", "Description", "TaskCategorySummary",
    "TaskDiscipline", "TaskDisciplineSummary", "Discipline", "TaskType", "ResponsibleCompany",
    "ResponsibleCompanyID", "ProcessBreakdown", "ProcessBreakdownIdentifier",
    "ProcessBreakdownUp3Identifier", "PhysicalLocationUp1Summary", "ApprovedDate", "ApprovedBy",
    "TaskState", "AssetType", "LoopName", "TaskModelName", "TaskPrioritySummary",
    "DocumentDescription", "TaskPriority", "TaskCategory",
]

GET_JS = """((P) => {
    return new Promise((res) => {
        let done = false;
        const t = setTimeout(() => { if (!done) { done = true; res({timeout: true}); } }, P.to);
        dalc.Get(P.v, P.f, P.r || [], {Distinct:false, FirstRecordOrdinal: P.off, NumberOfRecords:P.n}, "Flat",
            (ok) => {
                if (done) return; done = true; clearTimeout(t);
                let cols = [], rows = [], total = null;
                const vals = (ok && ok.Values) || [];
                if (vals.length) cols = Object.keys(vals[0]).filter(k=>!k.startsWith('__'));
                rows = JSON.parse(JSON.stringify(vals || []));
                if (ok && ok.Request) {
                    const tr = ok.Request.TotalRecords;
                    if (typeof tr !== 'undefined' && tr !== null) total = tr;
                }
                res({cols: cols, rows: rows, total: total, n: rows.length});
            },
            (err) => { if (!done) { done = true; clearTimeout(t); res({err: String((err && err.message) || err)}); } }, false);
    });
})"""


def pr(*a):
    s = " ".join(str(x) for x in a)
    try:
        print(s, flush=True)
    except Exception:
        pass
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(time.strftime("[%Y-%m-%d %H:%M:%S] ") + s + "\n")
    except Exception:
        pass


def call(page, view, fields, restrictions=None, n=PAGE, off=0, to=120000):
    return page.evaluate("(%s)(%s)" % (GET_JS, json.dumps({
        "v": view, "f": fields, "r": restrictions or [], "n": n, "off": off, "to": to,
    })))


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


def js_fill(page, sel, val):
    try:
        return page.evaluate("""([sel, val]) => {
            const el = document.querySelector(sel);
            if (!el) return 'no-el';
            const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
            const setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
            setter.call(el, val);
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
            return 'ok';
        }""", [sel, val])
    except Exception as e:
        return 'err:' + str(e)[:50]


def auto_login(page):
    creds = {}
    try:
        with open(os.path.join(BASE_DIR, "creds.env")) as f:
            for line in f:
                line = line.strip()
                if "=" in line and not line.startswith("#"):
                    k, v = line.split("=", 1)
                    creds[k.strip()] = v.strip()
    except Exception as e:
        pr("creds err", str(e)[:120])
    root_login = "https://wly04-sc.intergraphsmartcloud.com//Login.aspx?referrer=/ISC/Tools/vDashboardsUsers/Switchboard.htm"
    try:
        page.goto(root_login, wait_until="domcontentloaded", timeout=45000)
    except Exception:
        pass
    page.wait_for_timeout(2500)
    js_fill(page, "#txtUserName", creds.get("SC_USER", ""))
    js_fill(page, "#txtUserPass", creds.get("SC_PASS", ""))
    page.wait_for_timeout(400)
    clicked = page.evaluate("""() => {
        const el = document.getElementById('btnSSOLogin');
        if (el) { el.click(); return 'ok'; }
        return 'no-btn';
    }""")
    pr("ssologin click:", clicked)
    page.wait_for_timeout(10000)
    if "auth.intergraphsmartcloud.com" in page.url:
        pr("sso handoff -> verify password page")
        page.wait_for_timeout(2000)
        js_fill(page, "input[type=password]", creds.get("SC_PASS", ""))
        page.wait_for_timeout(400)
        sub = page.evaluate("""() => {
            const cand = document.querySelector('input[type=submit], button[type=submit], button[data-action-button-primary]');
            if (cand) { cand.click(); return 'submitted'; }
            return 'no-submit';
        }""")
        pr("verify submit:", sub)
        page.wait_for_timeout(12000)
    cnt = 0
    while cnt < 6:
        try:
            if count_products(page) >= 0:
                return True
        except Exception:
            pass
        page.wait_for_timeout(5000)
        cnt += 1
    return False


def dc(v):
    if v is None:
        return None
    s = str(v).strip()
    if s in ("E", "I", "T", "M", "S", "P"):
        return s
    for p in ("E -", "I -", "T -", "M -", "S -", "P -"):
        if s.startswith(p):
            return p[0]
    return s[:2].upper()


DISC_LABELS = {
    "E": "Electrical (E)",
    "I": "Instrumentation (I)",
    "T": "Telecom (T)",
    "M": "Mechanical (M)",
    "S": "S (S)",
    "P": "Process (P)",
}


def date_key(v):
    s = str(v or "").split(" ")[0]
    if "T" in s:
        s = s.split("T")[0]
    if len(s) == 10 and s[4] == "-":
        return s
    parts = s.split("/")
    if len(parts) == 3:
        try:
            m, d, y = int(parts[0]), int(parts[1]), int(parts[2])
            if y < 100:
                y += 2000
            return "%04d-%02d-%02d" % (y, m, d)
        except Exception:
            return ""
    return ""


def _listener_pid(port):
    try:
        out = subprocess.run(["netstat", "-ano", "-p", "TCP"], capture_output=True, text=True, timeout=60).stdout
        for line in out.splitlines():
            parts = line.split()
            if len(parts) >= 5 and parts[0] == "TCP" and parts[1].endswith(":%d" % port) and "LISTENING" in line:
                return int(parts[-1])
    except Exception as e:
        pr("netstat err:", str(e)[:80])
    return None


def _pid_cmdline(pid):
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "(Get-CimInstance Win32_Process -Filter 'ProcessId=%d').CommandLine" % pid],
            capture_output=True, text=True, timeout=30)
        return (r.stdout or "").strip()
    except Exception:
        return ""


def _our_port():
    """Debug port of a LISTENING chrome that uses OUR profile dir (the kept-open session)."""
    for port in range(DEBUG_PORT, DEBUG_PORT + 10):
        pid = _listener_pid(port)
        if not pid:
            continue
        if PROFILE_DIR.replace("/", "\\").lower() in _pid_cmdline(pid).lower():
            return port
    return None


def _free_port():
    for port in range(DEBUG_PORT, DEBUG_PORT + 10):
        if not _listener_pid(port):
            return port
    return DEBUG_PORT


def _kill_profile_chrome():
    """Never kill the browser window.

    The user works in this same Chrome (that is where they sign in and open the
    task pages). Killing it closes their tabs and loses the session, so the sync
    only ever backs off and retries later.
    """
    pr("keeping the browser open - we never kill it (user works in it)")
    return False


def open_session(pw):
    """Attach to the one always-open browser the user works in (port 9222) and read
    from a dedicated tab, so there is never a second browser window."""
    if _listener_pid(DEBUG_PORT):
        port = DEBUG_PORT
    else:
        port = _our_port()
    if port:
        try:
            conn = pw.chromium.connect_over_cdp("http://127.0.0.1:%d" % port, timeout=15000)
            # reuse OUR dedicated pull tab if it exists, else open a BRAND-NEW tab.
            # never reuse the user's work tabs (EditForm / Switchboard / anything they opened).
            ours = None
            for c in (conn.contexts or []):
                for p in c.pages:
                    try:
                        if p.url.startswith(MARK_URL):
                            ours = p
                            break
                    except Exception:
                        pass
                if ours:
                    break
            if ours is None:
                home = None
                try:
                    home = conn.contexts[0] if conn.contexts else conn.new_context()
                except Exception:
                    pass
                if home is None or (hasattr(home, "new_page") is False):
                    home = conn.contexts[0] if conn.contexts else None
                if home is not None:
                    try:
                        ours = home.new_page()
                        try:
                            ours.goto(MARK_URL, wait_until="domcontentloaded", timeout=60000)
                        except Exception:
                            pass
                    except Exception as e:
                        pr("new tab err: %s -> using first page file-less" % str(e)[:60])
                        ours = None
            if ours is None:
                raise RuntimeError("could not create a pull tab")
            pr("reused the user's open browser (port %d) in a fresh tab" % port)
            return None, ours
        except Exception as e:
            pr("reuse failed: %s -> leaving the browser alone, will retry later" % str(e)[:70])
            time.sleep(2)
            raise
    # no debuggable browser yet - launch ONE kept-open browser and reuse it forever
    port = _free_port()
    last = None
    for attempt in range(4):
        try:
            ctx = pw.chromium.launch_persistent_context(
                PROFILE_DIR, channel="chrome", headless=False,
                args=["--remote-debugging-port=%d" % port],
                viewport=None, timeout=60000,
            )
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            pr("launched fresh kept-open browser (port %d) profile=%s" % (port, PROFILE_DIR))
            return ctx, page
        except Exception as e:
            last = e
            pr("launch attempt %d failed: %s" % (attempt + 1, str(e)[:100]))
            time.sleep(3)
    raise last


def pull_live(page):
    try:
        page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
    except Exception as e:
        pr("goto err:", str(e)[:100])
    if not wait_products(page, timeout_sec=45):
        pr("session expired -> auto-login")
        login_ok = False
        for attempt in range(3):
            auto_login(page)
            try:
                page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=90000)
            except Exception as e:
                pr("post-login goto err:", str(e)[:100])
            if wait_products(page, timeout_sec=60):
                login_ok = True
                break
            pr("auto-login attempt %d still not logged in" % (attempt + 1))
        if not login_ok:
            for i in range(120):
                if count_products(page) >= 0:
                    break
                page.wait_for_timeout(5000)
    if count_products(page) < 0:
        pr("WAITING FOR MANUAL LOGIN...")
        if not wait_products(page, timeout_sec=120):
            return None
    pr("LOGGED IN")
    restrictions = [
        {"FieldName": "ResponsibleCompanyID", "Term": "=", "ValueCollection": [CPP_CID]},
        {"FieldName": "PhysicalLocationUp1Summary", "Term": "like", "ValueCollection": ["PS5 -"]},
        {"FieldName": "TaskCategorySummary", "Term": "like", "ValueCollection": ["PCOM - Precommissioning"]},
    ]
    all_rows = []
    off = 0
    while True:
        r = call(page, "vTasks_TestsPlanned", FIELDS, restrictions, off=off)
        if r.get("timeout"):
            pr("timeout off=%d" % off)
            break
        if r.get("err"):
            pr("err off=%d: %s" % (off, r["err"][:120]))
            break
        n = r.get("n", 0)
        all_rows.extend(r.get("rows", []))
        pr("pulled off=%d n=%d accum=%d" % (off, n, len(all_rows)))
        if n < PAGE:
            break
        off += PAGE
    return all_rows


def build_state(rows):
    eit = [r for r in rows if dc(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline")) in ("E", "I", "T")]
    by = Counter(); closed = Counter(); opened = Counter(); today = 0
    today_key = time.strftime("%Y-%m-%d")
    daily_ts = {}
    ms_today = Counter()
    for r in eit:
        d = dc(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline"))
        by[d] += 1
        if str(r.get("TaskState")) == "Closed":
            closed[d] += 1
            k = date_key(r.get("ApprovedDate"))
            if k == today_key:
                today += 1
                ms = milestone_of(r.get("ProcessBreakdown"))
                if ms:
                    ms_today[(ms, d)] += 1
            if k:
                daily_ts.setdefault(k, Counter())[d] += 1
        else:
            opened[d] += 1
    daily = []
    for i in range(29, -1, -1):
        t = time.localtime(time.time() - i * 86400)
        label = time.strftime("%Y-%m-%d", t)
        c = daily_ts.get(label, {})
        daily.append({"label": label, "E": c.get("E", 0), "I": c.get("I", 0), "T": c.get("T", 0)})
        daily[-1]["Total"] = daily[-1]["E"] + daily[-1]["I"] + daily[-1]["T"]
    recent = recent_closed_rows(rows)
    closed_this_week, closed_this_month = 0, 0
    now_dt = datetime.datetime.now()
    wk = datetime.date.today() - datetime.timedelta(days=datetime.date.today().weekday())
    wk_start = "%04d-%02d-%02d" % (wk.year, wk.month, wk.day)
    for r in eit:
        if str(r.get("TaskState")) != "Closed":
            continue
        k = date_key(r.get("ApprovedDate"))
        if k and k >= wk_start:
            closed_this_week += 1
        if k and k[:7] == now_dt.strftime("%Y-%m"):
            closed_this_month += 1
    today_milestone = []
    for (ms, d), c in sorted(ms_today.items()):
        label = str(ms)
        for prefix in ("PS5 - ", "PS5-"):
            if label.startswith(prefix):
                label = label[len(prefix):]
                break
        today_milestone.append({
            "milestone": label,
            "disc": d,
            "discipline": DISC_LABELS.get(d, d),
            "count": c,
        })
    MS_GROUPS = {"Closed": "Closed", "Submitted": "Submitted", "To be completed": "To be completed"}
    MS_ORDER = ["Closed", "Submitted", "To be completed", "Other"]
    M = {}
    for r in rows:
        ms = milestone_of(r.get("ProcessBreakdown"))
        if not ms:
            continue
        sub = M.setdefault(ms, Counter())
        st = str(r.get("TaskState") or "").strip()
        sub["total"] += 1
        if st == "Closed":
            sub["closed"] += 1
        sub[MS_GROUPS.get(st, "Other")] += 1
    milestone_summary = []
    for ms in sorted(M):
        sub = M[ms]
        label = str(ms)
        for prefix in ("PS5 - ", "PS5-"):
            if label.startswith(prefix):
                label = label[len(prefix):]
                break
        tot = sub["total"]
        cls = sub["closed"]
        milestone_summary.append({
            "label": label,
            "raw": ms,
            "total": tot,
            "closed": cls,
            "remaining": tot - cls,
            "pct": round(cls * 100.0 / tot, 1) if tot else 0,
            "status": {k: sub.get(k, 0) for k in MS_ORDER},
        })
    return {
        "source": "Smart Completions (live authenticated pull)",
        "scope": "PS5 / CPP AGI / PCOM (ITR tests) / E&I&T / vTasks_TestsPlanned",
        "updated": time.strftime("%Y-%m-%d %H:%M:%S"),
        "total_tasks_ps5": len(rows),
        "eit_total": len(eit),
        "eit_by_discipline": dict(by),
        "closed": sum(closed.values()),
        "open": sum(opened.values()),
        "today_closed": today,
        "closed_by_discipline": dict(closed),
        "open_by_discipline": dict(opened),
        "daily": daily,
        "closed_this_week": closed_this_week,
        "closed_this_month": closed_this_month,
        "today_milestone": today_milestone,
        "today_milestone_total": sum(ms_today.values()),
        "milestone_summary": milestone_summary,
        "recent_closed": recent,
    }


def parse_dt(v):
    s = str(v or "").strip()
    if not s:
        return ""
    t = s.replace("T", " ")
    if "." in t:
        t = t.split(".")[0]
    for fmt in ("%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %H:%M:%S",
                "%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            import datetime
            return datetime.datetime.strptime(t.strip(), fmt).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            continue
    return ""


def recent_closed_rows(rows, n=120):
    out = []
    eit = [r for r in rows if dc(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline")) in ("E", "I", "T")]
    for r in eit:
        if str(r.get("TaskState")) != "Closed":
            continue
        raw = str(r.get("ApprovedDate") or "")
        out.append({
            "tag": r.get("AssetTag") or "",
            "task": r.get("TaskName") or r.get("ID") or "",
            "system": r.get("ProcessBreakdown") or "",
            "loop": r.get("LoopName") or "",
            "disc": r.get("TaskDisciplineSummary") or r.get("TaskDiscipline") or "",
            "cat": r.get("TaskCategorySummary") or "",
            "approved": raw[:19],
            "_dt": parse_dt(raw),
        })
    out.sort(key=lambda x: x["_dt"], reverse=True)
    for x in out:
        x.pop("_dt", None)
    return out[:n]


def write_local(state):
    with open(os.path.join(WS, "itr_live_state.json"), "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    with open(os.path.join(DATA_DIR, "itr_kpi.json"), "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    src = os.path.join(WS, "dashboard_data.json")
    if os.path.exists(src):
        try:
            with open(src, "r", encoding="utf-8") as f:
                dd = json.load(f)
            dd.setdefault("itr_summary", {})
            dd["itr_summary"]["live"] = {
                "closed": state["closed"],
                "open": state["open"],
                "total": state["eit_total"],
                "by_discipline": state["closed_by_discipline"],
                "updated": state["updated"],
            }
            dd["date"] = time.strftime("%d %b %Y").upper()
            with open(src, "w", encoding="utf-8") as f:
                json.dump(dd, f, ensure_ascii=False, indent=1)
            pr("updated dashboard_data.json itr_summary.live")
        except Exception as e:
            pr("dashboard_data.json update err:", str(e)[:120])


def ensure_web_files():
    local = {
        "live_itr.js": LIVE_ITER_JS,
        "itr_live.html": ITR_LIVE_HTML,
    }
    for name, content in local.items():
        with open(os.path.join(WS, name), "w", encoding="utf-8") as f:
            f.write(content)


def inject_into_index(clone):
    if not os.path.isdir(clone):
        return False
    p = os.path.join(clone, "index.html")
    if not os.path.exists(p):
        return False
    s = open(p, "r", encoding="utf-8", errors="ignore").read()
    marker = '<script src="live_itr.js"'
    if marker in s:
        return True
    tag = '<script src="live_itr.js" charset="utf-8"></script>'
    idx = s.rfind("</body>")
    if idx != -1:
        s = s[:idx] + tag + "\n" + s[idx:]
    else:
        s = s + "\n" + tag
    open(p, "w", encoding="utf-8").write(s)
    return True


def git_push(clone, state):
    if not os.path.isdir(clone):
        return None
    git = ["git", "-C", clone]
    targeted = ["itr_live_state.json", "live_itr.js", "itr_live.html", "index.html"]
    a0 = git + ["add", "--"] + targeted
    r = subprocess.run(a0, capture_output=True, text=True, timeout=600)
    pr(" >> " + " ".join(a0) + " rc=%d" % r.returncode)
    a1 = git + ["commit", "-m", "ITR live update %s (closed=%d)" % (
        time.strftime("%Y-%m-%d %H:%M"), state["closed"])]
    r = subprocess.run(a1, capture_output=True, text=True, timeout=600)
    pr(" >> " + " ".join(a1) + " rc=%d" % r.returncode, (r.stdout or "").strip()[-120:])
    for attempt in range(4):
        r = subprocess.run(git + ["push", "origin", "main"], capture_output=True, text=True, timeout=1200)
        if r.returncode == 0:
            pr(" >> git push OK")
            return 0
        err = (r.stderr or "").strip()
        if "fetch first" in err or "rejected" in err:
            pr(" >> pull --rebase (attempt %d)" % (attempt + 1))
            rr = subprocess.run(git + ["pull", "--rebase", "origin", "main"],
                                capture_output=True, text=True, timeout=1200)
            rrerr = (rr.stderr or "").strip()
            if rr.returncode != 0:
                pr(" !! rebase failed:", " | ".join(rrerr.splitlines()[-5:])[:400])
                return rr.returncode
            continue
        pr(" !! push failed:", err.splitlines()[-3:])
        return r.returncode
    return 1


def _token():
    try:
        with open(TOK_FILE, "r", encoding="utf-8") as f:
            t = f.read().strip()
        return t or None
    except Exception:
        return None


def ensure_pages_clone():
    if not os.path.isdir(PAGE_CLONE):
        r = subprocess.run(
            ["git", "clone", "--depth", "1",
             "https://github.com/Mohamedgawad1/%s.git" % PAGE_REPO, PAGE_CLONE],
            capture_output=True, text=True, timeout=300)
        if r.returncode != 0:
            pr(" !! pages clone failed:", (r.stderr or "")[-200:])
            return False
    git = ["git", "-C", PAGE_CLONE]
    subprocess.run(git + ["remote", "set-url", "origin",
                          "https://github.com/Mohamedgawad1/%s.git" % PAGE_REPO],
                   capture_output=True, text=True, timeout=120)
    subprocess.run(git + ["fetch", "--depth", "1", "origin", "main"],
                   capture_output=True, text=True, timeout=300)
    subprocess.run(git + ["reset", "--hard", "origin/main"],
                   capture_output=True, text=True, timeout=120)
    return True


def pages_push(state):
    """Sync the ITR LIVE files to the PUBLIC platform repo so the live page stays fresh."""
    tok = _token()
    if not tok:
        pr(" !! no token found — skipping pages push")
        return None
    if not ensure_pages_clone():
        return None
    git = ["git", "-C", PAGE_CLONE]
    for fn in ("itr_live_state.json", "live_itr.js", "itr_live.html"):
        src = os.path.join(WS, fn)
        if os.path.exists(src):
            try:
                shutil.copy2(src, os.path.join(PAGE_CLONE, fn))
            except Exception as e:
                pr(" !! copy %s err: %s" % (fn, str(e)[:80]))
    subprocess.run(git + ["add", "--"] + ["itr_live_state.json", "live_itr.js", "itr_live.html"],
                   capture_output=True, text=True, timeout=300)
    r = subprocess.run(
        git + ["commit", "-m", "ITR live update %s (closed=%d)" % (
            time.strftime("%Y-%m-%d %H:%M"), state["closed"])],
        capture_output=True, text=True, timeout=300)
    pr(" >> pages commit rc=%d %s" % (r.returncode, (r.stdout or "").strip()[-80:]))
    for attempt in range(4):
        push = git + ["push",
                      "https://x-access-ps5:%s@github.com/Mohamedgawad1/%s.git" % (tok, PAGE_REPO),
                      "main"]
        r = subprocess.run(push, capture_output=True, text=True, timeout=600)
        if r.returncode == 0:
            pr(" >> pages push OK (closed=%d)" % state["closed"])
            return 0
        err = (r.stderr or "").strip()
        if "fetch first" in err or "rejected" in err:
            pr(" >> pages pull --rebase (attempt %d)" % (attempt + 1))
            rr = subprocess.run(git + ["pull", "--rebase", "origin", "main"],
                                capture_output=True, text=True, timeout=600)
            if rr.returncode != 0:
                pr(" !! pages rebase failed:", (rr.stderr or "")[-300:])
                return rr.returncode
            continue
        pr(" !! pages push failed:", err.splitlines()[-3:])
        return r.returncode
    return 1


def run_once():
    pr("=" * 60)
    pr("ITR ONLINE SYNC START")
    try:
        return _run_once()
    except Exception:
        pr("!! CRASH:\n" + traceback.format_exc()[-4000:])
        return 1


def _run_once():
    pr("=" * 60)
    pr("ITR ONLINE SYNC START")
    with sync_playwright() as pw:
        ctx, page = open_session(pw)
        if page is None:
            pr("SKIPPED this cycle - browser is busy with the user (nothing was closed)")
            return 0
        rows = None
        for attempt in range(2):
            try:
                rows = pull_live(page)
                break
            except Exception as e:
                pr("pull attempt %d failed: %s — reopening session" % (attempt + 1, str(e)[:120]))
                try:
                    ctx, page = open_session(pw)
                except Exception as e2:
                    pr("reopen failed too: %s" % str(e2)[:120])
                    break
        if not rows:
            pr("!! no rows — abort")
            if ctx is not None and KEEP_OPEN:
                try:
                    ctx.close()
                except Exception:
                    pass
            return 1
        load_milestone_map()
        state = build_state(rows)
        pr("RESULT:", json.dumps(state, ensure_ascii=False))
        write_local(state)
        ensure_web_files()
        injected = inject_into_index(CLONE)
        pr("index.html injection ok:", injected)
        with open(os.path.join(CLONE, "itr_live_state.json") if os.path.isdir(CLONE) else os.devnull,
                  "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
            pr("wrote itr_live_state.json -> ps5-dashboard")
        pr("pushing online...")
        rc = git_push(CLONE, state)
        pr("PUSH rc=%s" % (rc if rc is not None else "no clone"))
        pr("pushing to live platform repo...")
        prc = pages_push(state)
        pr("PAGES PUSH rc=%s" % (prc if prc is not None else "skipped"))
        pr("DONE")
        if not KEEP_OPEN:
            if ctx is not None:
                try:
                    ctx.close()
                except Exception:
                    pass
            return 0
        # keep the session alive: switch to Switchboard, wait until user closes it
        try:
            page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
            page.bring_to_front()
        except Exception:
            pass
        pr("Session KEPT OPEN (Switchboard). Close the browser window to finish.")
        while not page.is_closed():
            time.sleep(5)
        pr("Browser window closed — done")
    return 0
    return 0


LIVE_ITER_JS = """(() => {
  const LS = 'live_itr_state';
  function fmt(n){ return (n||0).toLocaleString('en-US'); }
  function build(update){
    const e = document.querySelector('#itr-live-badge');
    if(!e) return;
    const cb = update.closed, eb = (update.closed_by_discipline||{}).E||0,
          ib = (update.closed_by_discipline||{}).I||0, tb = (update.closed_by_discipline||{}).T||0;
    e.innerHTML =
      '<div class="itr-badge-title">ITR Closed &middot; Live</div>' +
      '<div class="itr-badge-big">' + fmt(cb) + '</div>' +
      '<div class="itr-badge-sub">E ' + fmt(eb) + ' &middot; I ' + fmt(ib) + ' &middot; T ' + fmt(tb) +
      ' &middot; Total E&I&T ' + fmt(update.eit_total) + '</div>' +
      '<div class="itr-badge-time">Updated ' + (update.updated||'') + '</div>';
  }
  function syncCards(update){
    const set = (id, v) => { const el = document.getElementById(id); if(el) el.textContent = v; };
    const tot = update.closed||0, all = update.eit_total||0, today = update.today_closed||0;
    set('todayClosed', fmt(today));
    set('totalClosed', fmt(tot));
    set('totalPct', all ? (tot/all*100).toFixed(2) : '0.00');
    document.querySelectorAll('.kpi').forEach(c => {
      const lbl = c.querySelector('.lbl'); if(!lbl) return;
      const val = c.querySelector('.val'); if(!val) return;
      const t = lbl.textContent||'';
      if(t.indexOf('CPP AGI EIT') !== -1 && t.indexOf('Total ITRs Closed') !== -1){
        val.textContent = fmt(tot) + ' / ' + fmt(all);
      } else if(t.trim() === 'Today Closed'){
        val.textContent = fmt(today);
      } else if(t.trim() === 'Closed This Week'){
        val.textContent = fmt(update.closed_this_week);
      } else if(t.trim() === 'Closed This Month'){
        val.textContent = fmt(update.closed_this_month);
      }
    });
  }
  function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  function discChip(d){
    const c = (d||'').toLowerCase();
    const cls = c.indexOf('e')===0 ? '#065f46' : c.indexOf('i')===0 ? '#78350f' : c.indexOf('t')===0 ? '#831843' : '#1e293b';
    const col = c.indexOf('e')===0 ? '#a7f3d0' : c.indexOf('i')===0 ? '#fde68a' : c.indexOf('t')===0 ? '#fbcfe8' : '#cbd5e1';
    return '<span style="display:inline-block;padding:1px 8px;border-radius:10px;font-size:10px;background:'+cls+';color:'+col+'">'+esc(d)+'</span>';
  }
  function syncRecent(update){
    const list = Array.isArray(update.recent_closed) ? update.recent_closed : [];
    let sec = document.getElementById('live-recent-sec');
    if(!sec){
      sec = document.createElement('section');
      sec.id = 'live-recent-sec';
      sec.className = 'section active';
      sec.style.border = '1px solid var(--border)';
      sec.style.background = 'var(--panel2)';
      sec.style.borderRadius = '8px';
      sec.style.padding = '14px';
      sec.style.marginBottom = '16px';
      sec.innerHTML =
        '<div class="section-title" style="margin:0 0 10px">🕒 Latest ITR Closures &mdash; Live (Subsystem / Asset / Task)</div>' +
        '<div class="wrap" style="overflow-x:auto">' +
        '<table id="live-recent-tbl" style="width:100%;border-collapse:collapse;font-size:12px;white-space:nowrap">' +
        '<thead><tr style="background:var(--panel);text-transform:uppercase;font-size:10px;letter-spacing:.5px;color:var(--teal)">' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Task ID</th>' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Asset Tag</th>' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Subsystem</th>' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Loop</th>' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Disc</th>' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Category</th>' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Approved</th>' +
        '</tr></thead><tbody></tbody></table></div>' +
        '<div style="font-size:11px;color:var(--muted);margin-top:8px">Live from <b>itr_live_state.json</b> &mdash; updates automatically every 2 min. Time shown: <b id="live-recent-time">&mdash;</b></div>';
      const kpi = document.getElementById('sec-kpi');
      if(kpi && kpi.parentElement) kpi.parentElement.insertBefore(sec, kpi.nextSibling);
      else document.body.appendChild(sec);
    }
    const tb = sec.querySelector('#live-recent-tbl tbody');
    const tm = sec.querySelector('#live-recent-time');
    if(tm) tm.textContent = update.updated || '—';
    if(tb){
      tb.innerHTML = list.map(r =>
        '<tr>' +
        '<td style="padding:6px 9px;border:1px solid var(--border)">'+esc(r.task)+'</td>' +
        '<td style="padding:6px 9px;border:1px solid var(--border);color:var(--teal);font-weight:600">'+esc(r.tag)+'</td>' +
        '<td style="padding:6px 9px;border:1px solid var(--border)">'+esc(r.system)+'</td>' +
        '<td style="padding:6px 9px;border:1px solid var(--border)">'+esc(r.loop)+'</td>' +
        '<td style="padding:6px 9px;border:1px solid var(--border)">'+discChip(r.disc)+'</td>' +
        '<td style="padding:6px 9px;border:1px solid var(--border)">'+esc(r.cat)+'</td>' +
        '<td style="padding:6px 9px;border:1px solid var(--border);color:var(--muted)">'+esc(r.approved)+'</td>' +
        '</tr>'
      ).join('') || '<tr><td colspan="7" style="padding:10px;text-align:center;color:var(--muted)">No closures yet &mdash; waiting for live data&hellip;</td></tr>';
    }
  }
  function syncChart(update){
    const daily = Array.isArray(update.daily) ? update.daily : null;
    if(!daily || !daily.length) return;
    const labels = daily.map(d=>d.label);
    const tot = daily.map(d=>d.Total||0);
    const upd = (id, apply) => {
      const cv = document.getElementById(id);
      if(!cv) return;
      let ch = null;
      try{ ch = window.Chart && Chart.getChart ? Chart.getChart(cv) : null; }catch(e){}
      if(!ch || !ch.data || !ch.data.datasets) return;
      ch.data.labels = labels;
      apply(ch.data.datasets);
      ch.update();
    };
    upd('chartCombinedDaily', ds => { if(ds[0]) ds[0].data = tot; });
    upd('chartDaily', ds => {
      ['E','I','T'].forEach((k,i)=>{ if(ds[i]) ds[i].data = daily.map(d=>d[k]||0); });
    });
  }
  function syncMilestone(update){
    const t = document.getElementById('todayMsTotal');
    if(t) t.textContent = fmt(update.today_milestone_total||0);
    const body = document.getElementById('todayMsBody');
    if(!body) return;
    const tm = Array.isArray(update.today_milestone) ? update.today_milestone : [];
    if(!tm.length){
      if(!body.dataset._live) body.innerHTML = '<tr><td colspan="5" style="padding:10px;text-align:center;color:var(--muted)">No closures today yet &mdash; waiting for live data&hellip;</td></tr>';
      return;
    }
    body.dataset._live = '1';
    const msMap = {};
    tm.forEach(r => {
      if(!msMap[r.milestone]) msMap[r.milestone] = {E:0, I:0, T:0};
      const d = String(r.disc||'');
      if(msMap[r.milestone][d] !== undefined) msMap[r.milestone][d] = r.count;
    });
    body.innerHTML = Object.keys(msMap).sort().map(ms => {
      const d = msMap[ms];
      const rowTotal = d.E + d.I + d.T;
      return '<tr>' +
        '<td style="text-align:left;border:1px solid var(--border);font-weight:bold;color:var(--red);background:var(--card3);">'+esc(ms)+'</td>' +
        '<td style="text-align:center;border:1px solid var(--border);background:#FCE4D6;color:var(--red);">'+fmt(d.E)+'</td>' +
        '<td style="text-align:center;border:1px solid var(--border);background:#FFF2CC;color:var(--red);">'+fmt(d.I)+'</td>' +
        '<td style="text-align:center;border:1px solid var(--border);background:#D9D9D9;color:var(--red);">'+fmt(d.T)+'</td>' +
        '<td style="text-align:center;border:1px solid var(--border);font-weight:bold;color:var(--red);background:var(--card3);">'+fmt(rowTotal)+'</td>' +
        '</tr>';
    }).join('');
  }
  const MS_ORDER = ['Closed','Submitted','To be completed','Other'];
  const MS_COLORS = {'Closed':'#1a8a4a','Submitted':'#2563eb','To be completed':'#c8940a','Other':'#9a8d7c'};
  function syncMilestoneSummary(update){
    const list = Array.isArray(update.milestone_summary) ? update.milestone_summary : [];
    const cards = document.getElementById('msCards');
    if(cards){
      cards.innerHTML = list.map(s =>
        '<div class="kpi">' +
          '<div class="val">' + fmt(s.closed) + ' / ' + fmt(s.total) + '</div>' +
          '<div class="lbl">🎯 ' + esc(s.label) + ' &mdash; ' + s.pct + '%</div>' +
          '<div class="progress-bar"><div class="progress-fill" style="width:' + s.pct + '%"></div></div>' +
        '</div>'
      ).join('') ||
        '<div class="kpi"><div class="val">0 / 0</div><div class="lbl">🎯 &mdash;</div><div class="progress-bar"><div class="progress-fill" style="width:0%"></div></div></div>';
    }
    const cv = document.getElementById('chartMilestone');
    if(!cv || !list.length) return;
    let ch = null;
    try{ ch = window.Chart && Chart.getChart ? Chart.getChart(cv) : null; }catch(e){}
    if(!ch || !ch.data || !ch.data.datasets) return;
    ch.data.labels = list.map(s=>s.label);
    MS_ORDER.forEach((st,i)=>{
      let ds = ch.data.datasets[i];
      if(!ds || ds.label !== st) ds = ch.data.datasets.find(d=>d.label===st);
      if(ds) ds.data = list.map(s=> (s.status&&s.status[st]) || 0);
    });
    ch.update();
  }
  function fetchItr(){
    fetch('itr_live_state.json?v=' + Math.floor(Date.now()/120000), {cache:'no-store'})
      .then(r => r.json())
      .then(u => { try{ localStorage.setItem(LS, JSON.stringify(u)); }catch(e){} build(u); syncCards(u); syncRecent(u); syncChart(u); syncMilestone(u); syncMilestoneSummary(u); })
      .catch(() => { try{ const o=localStorage.getItem(LS); if(o){ const u=JSON.parse(o); build(u); syncCards(u); syncRecent(u); syncChart(u); syncMilestone(u); syncMilestoneSummary(u); } }catch(e){} });
  }
  if(!document.querySelector('#itr-live-css')) {
    const st = document.createElement('style'); st.id='itr-live-css';
    st.textContent = '#itr-live-badge{position:fixed;right:16px;bottom:16px;z-index:99999;background:#0b2f56;color:#fff;border:1px solid #38bdf8;border-radius:12px;padding:10px 14px;font-family:Segoe UI,Arial,sans-serif;box-shadow:0 6px 18px rgba(0,0,0,.35);min-width:210px}'
      + '.itr-badge-title{font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#7dd3fc}'
      + '.itr-badge-big{font-size:30px;font-weight:800;line-height:1.1}'
      + '.itr-badge-sub{font-size:12px;color:#cbd5e1;margin-top:2px}'
      + '.itr-badge-time{font-size:10px;color:#94a3b8;margin-top:4px}';
    document.head.appendChild(st);
  }
  if(!document.querySelector('#itr-live-badge')){
    const d = document.createElement('div'); d.id='itr-live-badge';
    d.innerHTML = '<div class="itr-badge-title">ITR Closed &middot; Live</div><div class="itr-badge-big">&hellip;</div>';
    document.body.appendChild(d);
  }
  fetchItr();
  setInterval(fetchItr, 120000);
})();"""

ITR_LIVE_HTML = """<!DOCTYPE html>
<html lang="ar">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ITR Live — PS5 CPP AGI</title>
<style>
  body{font-family:'Segoe UI',Arial,sans-serif;background:#0f172a;color:#e2e8f0;margin:0;padding:20px}
  h1{font-size:20px;color:#38bdf8;margin:0 0 4px}
  .sub{font-size:12px;color:#94a3b8;margin-bottom:18px}
  .cards{display:flex;flex-wrap:wrap;gap:14px}
  .card{background:#0b2f56;border:1px solid #164e63;border-radius:12px;padding:16px 20px;min-width:150px;flex:1}
  .card .lbl{font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#7dd3fc}
  .card .num{font-size:34px;font-weight:800;margin-top:4px}
  .card.e .num{color:#34d399}.card.i .num{color:#fbbf24}.card.t .num{color:#f472b6}
  .card.total{background:#155e75;border-color:#38bdf8}
  .meta{margin-top:22px;font-size:12px;color:#94a3b8;line-height:1.7}
  .bar{height:10px;background:#1e293b;border-radius:6px;overflow:hidden;margin-top:8px}
  .bar span{display:block;height:100%}
  .err{color:#f87171}
  table{width:100%;border-collapse:collapse;margin-top:22px;font-size:12px}
  th,td{border:1px solid #1e293b;padding:6px 9px;text-align:left;white-space:nowrap}
  th{background:#0b2f56;color:#7dd3fc;text-transform:uppercase;font-size:10px;letter-spacing:.5px}
  tr:nth-child(even){background:#111c33}
  .wrap{overflow-x:auto}
  .chip{display:inline-block;padding:1px 8px;border-radius:10px;font-size:10px}
  .chip.e{background:#065f46;color:#a7f3d0}.chip.i{background:#78350f;color:#fde68a}.chip.t{background:#831843;color:#fbcfe8}
</style>
</head>
<body>
<h1>كم ITR اتقفل — CPP AGI (E&I&amp;T) — PS5</h1>
<div class="sub">تحديث مباشر من منصة Smart Completions (automatic zoom)</div>
<div id="content" class="cards">
  <div class="card total"><div class="lbl">عدد المقفول الآن</div><div class="num" id="closed">&hellip;</div></div>
  <div class="card e"><div class="lbl">Electrical</div><div class="num" id="e">&hellip;</div></div>
  <div class="card i"><div class="lbl">Instrumentation</div><div class="num" id="i">&hellip;</div></div>
  <div class="card t"><div class="lbl">Telecom</div><div class="num" id="t">&hellip;</div></div>
</div>
<div class="cards" style="margin-top:14px">
  <div class="card"><div class="lbl">شغّال (Open)</div><div class="num" id="open" style="color:#fca5a5">&hellip;</div></div>
  <div class="card"><div class="lbl">إجمالي E&amp;I&amp;T</div><div class="num" id="total">&hellip;</div></div>
</div>
<div class="meta" id="meta">جارٍ التحديث…</div>
<h2 style="font-size:15px;color:#7dd3fc;margin-top:28px">آخر ITRs متقفلة — بالتفاصيل (باشر)</h2>
<div class="wrap"><table id="tbl">
<thead><tr><th>Task ID</th><th>Asset Tag</th><th>النظام (System)</th><th>Loop</th><th>Disc</th><th>التشغيل</th><th>تاريخ القفل</th></tr></thead>
<tbody id="rows"></tbody>
</table></div>
<script>
  function fmt(n){ return (n||0).toLocaleString('en-US'); }
  function render(u){
    document.getElementById('closed').textContent = fmt(u.closed);
    document.getElementById('open').textContent = fmt(u.open);
    document.getElementById('total').textContent = fmt(u.eit_total);
    document.getElementById('e').textContent = fmt((u.closed_by_discipline||{}).E||0);
    document.getElementById('i').textContent = fmt((u.closed_by_discipline||{}).I||0);
    document.getElementById('t').textContent = fmt((u.closed_by_discipline||{}).T||0);
    document.getElementById('meta').innerHTML =
      'آخر تحديث من المنصة: <b>' + (u.updated||'') + '</b><br>النطاق: ' + (u.scope||'') +
      '<br>يتحدث تلقائياً كل دقيقة. <span class="err" id="st"></span>';
    document.title = 'ITR Closed: ' + fmt(u.closed) + ' | PS5 CPP AGI';
    const tb = document.getElementById('rows');
    if(tb && Array.isArray(u.recent_closed)){
      tb.innerHTML = u.recent_closed.map(r=>{
        const esc = s => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
        return '<tr><td>'+esc(r.task)+'</td><td>'+esc(r.tag)+'</td><td>'+esc(r.system)+'</td>'+
               '<td>'+esc(r.loop)+'</td><td><span class="chip '+(r.disc||'').toLowerCase()+'">'+esc(r.disc)+'</span></td>'+
               '<td>'+esc(r.cat)+'</td><td>'+esc(r.approved)+'</td></tr>';
      }).join('');
    }
  }
  function go(){
    fetch('itr_live_state.json?v=' + Math.floor(Date.now()/60000), {cache:'no-store'})
      .then(r => r.json()).then(render)
      .catch(e => { document.getElementById('meta').innerHTML = '<span class="err">تعذر التحديث: ' + e + '</span>'; });
  }
  go(); setInterval(go, 60000);
</script>
</body>
</html>
"""

if __name__ == "__main__":
    if not KEEP_OPEN:
        pass
    if "--loop" in sys.argv:
        # keep ONE live session open across all pulls (never close it)
        with sync_playwright() as pw:
            ctx, page = open_session(pw)
            pr("LOOP READY (interval %ds, keep_open=%s)" % (LOOP_SECONDS, KEEP_OPEN))
            while True:
                try:
                    if page.is_closed():
                        pr("browser was closed — reopening session")
                        ctx, page = open_session(pw)
                    rows = pull_live(page)
                    if rows:
                        load_milestone_map()
                        state = build_state(rows)
                        write_local(state)
                        ensure_web_files()
                        inject_into_index(CLONE)
                        with open(os.path.join(CLONE, "itr_live_state.json") if os.path.isdir(CLONE) else os.devnull,
                                  "w", encoding="utf-8") as f:
                            json.dump(state, f, ensure_ascii=False, indent=2)
                        git_push(CLONE, state)
                        pages_push(state)
                    try:
                        page.goto(MARK_URL, wait_until="domcontentloaded", timeout=60000)
                    except Exception:
                        pass
                    if page.is_closed():
                        pr("browser closed — exiting loop")
                        break
                except Exception as e:
                    pr("LOOP ERR:", str(e)[:200])
                for _ in range(LOOP_SECONDS // 5):
                    if page.is_closed():
                        break
                    time.sleep(5)
    else:
        sys.exit(run_once())