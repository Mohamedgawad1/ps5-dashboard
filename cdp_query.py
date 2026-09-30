"""Query the tasks grid through a raw CDP websocket - no Playwright attach.

The long-running itr_online_sync.py holds a CDP connection, so Playwright's
connect_over_cdp never finishes. A plain DevTools websocket coexists with it.
"""

import json
import time
import urllib.parse
import urllib.request

import websocket

CDP = "http://127.0.0.1:9222"
TASKS_APP = "https://wly04-sc.intergraphsmartcloud.com/ISC/tools/vTasks_TestsPlanned/index.htm"
EDIT_URL = ("https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/EditForm.aspx"
            "?ID=%s&v=vTasks_TestsCompletion")
TAB = "cdp_tasks"


def log(msg):
    print(msg, flush=True)


def _json(path, method="GET", timeout=20):
    url = "%s%s" % (CDP, path)
    req = urllib.request.Request(url, method=method)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8", "ignore"))


def new_tab(url, timeout=25):
    return _json("/json/new?" + urllib.parse.quote(url, safe=""), method="PUT", timeout=timeout)


def tabs():
    return [t for t in _json("/json/list") if t.get("type") == "page"]


def close_tab(tab_id):
    try:
        _json("/json/close/" + tab_id, timeout=8)
    except Exception:
        pass


class Tab(object):
    """Minimal CDP client on one tab."""

    def __init__(self, ws_url, timeout=45):
        self.ws = websocket.create_connection(ws_url, timeout=timeout, suppress_origin=True)
        self.n = 0

    def send(self, method, params=None, timeout=45):
        self.n += 1
        mid = self.n
        self.ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        end = time.time() + timeout
        while time.time() < end:
            self.ws.settimeout(max(1, end - time.time()))
            try:
                msg = json.loads(self.ws.recv())
            except Exception as exc:
                return {"error": str(exc)[:120]}
            if msg.get("id") == mid:
                return msg.get("result") or {"error": str(msg.get("error"))[:200]}
        return {"error": "timeout"}

    def eval(self, expression, timeout=45, awaitp=False):
        """expression must be a real JS expression - wrap arrow functions yourself."""
        res = self.send("Runtime.evaluate", {
            "expression": expression.strip(),
            "returnByValue": True,
            "awaitPromise": awaitp,
            "userGesture": True,
        }, timeout=timeout)
        if "error" in res:
            return None
        r = res.get("result") or {}
        if r.get("subtype") == "error":
            return None
        return r.get("value")

    def close(self):
        try:
            self.ws.close()
        except Exception:
            pass


APP_READY = "!!document.getElementById('textbox_TaskName_search_vTasks_TestsPlanned')"
APP_HREF = "location.href"

def open_app_tab(reuse=True, seconds=120):
    """One tasks-app tab driven over a raw websocket."""
    tab = None
    if reuse:
        for t in tabs():
            if "vTasks_TestsPlanned" in (t.get("url") or "") and t.get("webSocketDebuggerUrl"):
                tab = Tab(t["webSocketDebuggerUrl"])
                if tab.eval(APP_HREF):
                    if tab.eval(APP_READY):
                        log("  (re-using the open tasks app tab)")
                        return tab
                tab.close()
                tab = None
    info = new_tab(TASKS_APP)
    tab = Tab(info["webSocketDebuggerUrl"])
    end = time.time() + seconds
    while time.time() < end:
        time.sleep(2)
        try:
            if tab.eval(APP_READY):
                log("  tasks app ready (raw cdp)")
                return tab
        except Exception:
            break
    tab.close()
    raise RuntimeError("tasks app did not load")


SEARCH = r"""(async (tid) => {
  const want = String(tid).trim().toUpperCase();
  const f = document.getElementById('textbox_TaskName_search_vTasks_TestsPlanned');
  if (!f) return {error: 'no search field'};
  f.scrollIntoView({block: 'center'}); f.focus(); f.value = tid;
  f.dispatchEvent(new Event('input', {bubbles: true}));
  f.dispatchEvent(new Event('change', {bubbles: true}));
  f.blur();
  const b = document.getElementById('btn_vTasks_TestsPlanned_Search_Search');
  if (b) b.click();

  const g = document.getElementById('vTasks_TestsPlanned_PrimaryList_grid');
  const read = () => {
    const kg = $(g).data('kendoGrid');
    if (!kg) return [];
    return kg.dataSource.view().map(r => {
      const tr = kg.tbody.find('tr[data-uid="' + r.uid + '"]');
      const a = tr && tr.length ? tr[0].querySelector('a[onclick*="vTasks_TestsCompletion"]') : null;
      const m = a ? (a.getAttribute('onclick') || '').match(/vTasks_TestsCompletion",\s*(\d+)/) : null;
      return {id: String(r.ID || ''), link: m ? m[1] : '',
              task: (r.TaskName || '').trim(), asset: r.AssetTag, type: r.TaskType};
    });
  };
  const end = Date.now() + 35000;
  let seen = [];
  while (Date.now() < end) {
    await new Promise(r => setTimeout(r, 500));
    seen = read();
    const hit = seen.filter(r => r.task.toUpperCase() === want && (r.id || r.link));
    if (hit.length) return {row: hit[0], total: $(g).data('kendoGrid').dataSource.total()};
  }
  return {row: null, total: seen.length, seen: seen.slice(0, 3)};
})"""


GET_JS = r"""((P) => {
    return new Promise((res) => {
        let done = false;
        const t = setTimeout(() => { if (!done) { done = true; res({timeout: true}); } }, P.to);
        const fin = (v) => { if (!done) { done = true; clearTimeout(t); res(v); } };
        try {
            dalc.Get(P.v, P.f, P.r || [], {Distinct:false, FirstRecordOrdinal: P.off, NumberOfRecords:P.n}, "Flat",
                (ok) => {
                    const vals = (ok && ok.Values) || [];
                    let total = null;
                    if (ok && ok.Request && typeof ok.Request.TotalRecords !== 'undefined') total = ok.Request.TotalRecords;
                    fin({rows: JSON.parse(JSON.stringify(vals || [])), total: total, n: vals.length});
                },
                (err) => fin({err: String((err && err.message) || err)}), false);
        } catch (e) { fin({err: String(e)}); }
    });
})"""


def find_task_dalc(app, task_id, timeout=45):
    """Ask the platform's own data layer for one TaskName - same call the live
    dashboard sync uses, so it answers in a fraction of a second instead of waiting
    for the grid to redraw. Returns {record, task, asset, type} or None."""
    payload = {
        "v": "vTasks_TestsPlanned",
        "f": ["ID", "TaskName", "AssetTag", "TaskType", "TaskState"],
        "r": [{"FieldName": "TaskName", "Term": "like", "ValueCollection": [task_id]}],
        "n": 20, "off": 0, "to": (timeout - 5) * 1000,
    }
    res = app.eval(GET_JS + "(" + json.dumps(payload) + ")", timeout=timeout, awaitp=True)
    if not isinstance(res, dict) or res.get("timeout") or res.get("err") or not res.get("rows"):
        return None
    want = str(task_id).strip().upper()
    for r in res["rows"]:
        name = str(r.get("TaskName") or "").strip()
        if name.upper() != want:
            continue
        rec = r.get("ID")
        if rec in (None, ""):
            continue
        return {"record": str(rec), "task": name, "asset": r.get("AssetTag"),
                "type": r.get("TaskType"), "state": r.get("TaskState"), "via": "dalc"}
    return None


def find_task_grid(app, task_id, timeout=60):
    """Grid search - the slower path, kept as the fallback."""
    call = "%s(%s)" % (SEARCH, json.dumps(task_id))
    res = app.eval(call, timeout=timeout, awaitp=True)
    if not isinstance(res, dict):
        return None
    row = res.get("row")
    if not row:
        app.last_seen = res
        return None
    rec = row.get("link") or row.get("id")
    return {"record": str(rec), "task": row.get("task"), "asset": row.get("asset"), "type": row.get("type")}


def find_task(app, task_id, timeout=60):
    """Prefer the fast data-layer call, fall back to the grid search."""
    hit = find_task_dalc(app, task_id, timeout=min(timeout, 45))
    if hit:
        return hit
    return find_task_grid(app, task_id, timeout=timeout)
