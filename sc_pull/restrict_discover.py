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

GET_JS = """((P) => {
    return new Promise((res) => {
        dalc.Get(P.v, P.f, P.r || [], {Distinct:false, FirstRecordOrdinal:0, NumberOfRecords:P.n}, "Flat",
            (ok) => {
                let cols = [], rows = [], total = null;
                const vals = (ok && ok.Values) || [];
                if (vals.length) cols = Object.keys(vals[0]).filter(k=>!k.startsWith('__'));
                rows = JSON.parse(JSON.stringify(vals || []));
                if (ok && ok.Request) {
                    const tr = ok.Request.TotalRecords;
                    if (typeof tr !== "undefined" && tr !== null) total = tr;
                }
                res({cols: cols, rows: rows, total: total, n: rows.length});
            },
            (err) => res({err: String((err && err.message) || err)}), false);
    });
})"""


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


def get(page, view, fields, restrictions=None, n=5):
    return page.evaluate("(%s)(%s)" % (
        GET_JS,
        json.dumps({"v": view, "f": fields, "r": restrictions or [], "n": n}),
    ))


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
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
        except Exception as e:
            print("nav:", str(e)[:100], flush=True)
        if not wait_products(page, timeout_sec=40):
            print("NO SESSION", flush=True)
            page.bring_to_front()
            if not wait_products(page, timeout_sec=900):
                ctx.close()
                return 1

        out = {}

        dis = get(page, "vTaskDisciplines", ["ID", "Summary", "Code", "Name", "Description"], n=60)
        out["vTaskDisciplines"] = dis
        print("== vTaskDisciplines ==", flush=True)
        for r in (dis.get("rows") or [])[:40]:
            print("  ", json.dumps({k: str(v)[:40] for k, v in r.items()}, ensure_ascii=False), flush=True)

        pbs = get(page, "vProcessBreakdownsProjects", ["ID", "Summary", "ProcessBreakdownUp1Summary", "ProcessBreakdown", "ProcessBreakdownUp1ID"], n=200)
        out["vProcessBreakdownsProjects"] = pbs
        print("== vProcessBreakdownsProjects ==", flush=True)
        for r in (pbs.get("rows") or []):
            s = json.dumps({k: str(v)[:60] for k, v in r.items()}, ensure_ascii=False)
            if "CPP" in s.upper() or "AGI" in s.upper():
                print("  (match) ", s, flush=True)
        print("  first3:", [ {k: str(v)[:40] for k,v in r.items()} for r in (pbs.get("rows") or [])[:3] ], flush=True)

        dis_rows = dis.get("rows") or []
        pbs_rows = pbs.get("rows") or []

        target_dis = {}
        for r in dis_rows:
            summ = str(r.get("Summary") or "").upper()
            if summ in ("E", "I", "T"):
                target_dis[summ] = r.get("ID")

        agi_rows = [r for r in pbs_rows if ("CPP" in json.dumps(r, ensure_ascii=False).upper()) or ("AGI" in json.dumps(r, ensure_ascii=False).upper())]
        agi_ids = list({r.get("ProcessBreakdownUp1ID") for r in agi_rows if r.get("ProcessBreakdownUp1ID") is not None})

        print("target disciplines:", target_dis, flush=True)
        print("agi system ids:", agi_ids, flush=True)

        fields_tasks = ["TaskDisciplineID", "TaskDiscipline", "ApprovedDate", "ApprovedBy", "AssetTag", "ProcessBreakdownUp1ID", "ProcessBreakdownUp1Summary", "ProcessBreakdown", "WeekOverdue", "TaskType", "TaskName", "ID"]

        tests = []
        for code, d_id in target_dis.items():
            r = get(page, "vTasks_TestsPlanned", fields_tasks, [{"FieldName": "TaskDisciplineID", "Term": "=", "ValueCollection": [d_id]}], n=3)
            tests.append({"what": "discipline " + code, "res": {k: (v if k != "rows" else None) for k, v in r.items()}, "sample": (r.get("rows") or [])[:2]})
        for sys_id in agi_ids[:3]:
            r = get(page, "vTasks_TestsPlanned", fields_tasks, [{"FieldName": "TaskDisciplineID", "Term": "=", "ValueCollection": [target_dis.get("E") or target_dis.get("I") or 0]}], n=1)
            r2 = get(page, "vTasks_TestsPlanned", fields_tasks, [{"FieldName": "ProcessBreakdownUp1ID", "Term": "=", "ValueCollection": [sys_id]}], n=3)
            tests.append({"what": "system " + str(sys_id), "resTotal": r2.get("total"), "err": r2.get("err"), "sample": (r2.get("rows") or [])[:2]})

        # also try 'like' term on display fields
        rl = get(page, "vTasks_TestsPlanned", fields_tasks, [{"FieldName": "TaskDiscipline", "Term": "like", "ValueCollection": ["E"]}], n=3)
        tests.append({"what": "TaskDiscipline like E", "res": {k: (v if k != "rows" else None) for k, v in rl.items()}})

        out["restrictionTests"] = tests
        for t in tests:
            print("TEST:", t, flush=True)

        with open(os.path.join(DATA_DIR, "restrict_test.json"), "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        print("saved ->", os.path.join(DATA_DIR, "restrict_test.json"), flush=True)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())