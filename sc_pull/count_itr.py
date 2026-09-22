import json
import os
import sys
import time
from collections import Counter

import openpyxl
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
PROFILE_DIR = os.path.join(BASE_DIR, "profile")

SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)

PAGE = 2000

TASK_FIELDS = [
    "ID", "TaskName", "AssetTag", "AssetDescription", "Description", "TaskCategorySummary",
    "TaskDiscipline", "TaskDisciplineSummary", "Discipline", "TaskType", "ResponsibleCompany",
    "ResponsibleCompanyID", "ProcessBreakdown", "ProcessBreakdownIdentifier",
    "ProcessBreakdownUp3Identifier", "PhysicalLocationUp1Summary", "ApprovedDate", "ApprovedBy",
    "TaskState", "AssetType", "LoopName", "TaskModelName", "TaskPrioritySummary", "DocumentDescription",
]

GET_JS = """((P) => {
    return new Promise((res) => {
        let done = false;
        const t = setTimeout(() => { if (!done) { done = true; res({timeout: true}); } }, P.to);
        dalc.Get(P.v, P.f, P.r || [], {Distinct:false, FirstRecordOrdinal:0, NumberOfRecords:P.n}, "Flat",
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


def call(page, view, fields, restrictions=None, n=PAGE, to=150000):
    return page.evaluate("(%s)(%s)" % (GET_JS, json.dumps({
        "v": view, "f": fields, "r": restrictions or [], "n": n, "to": to,
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


def excel_reference():
    p = os.path.join(BASE_DIR, "ovTasks_TestsPlanned_1369.xlsx")
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    ws = wb["Exported from SC"]
    rows = ws.iter_rows(values_only=True)
    header = list(next(rows))
    idx = {h: i for i, h in enumerate(header)}
    cpp = eit = closed = 0
    by_disc = Counter()
    for r in rows:
        comp = str(r[idx["Responsible Company (Summary)"]] or "")
        if comp != "CPP AGI":
            continue
        disc = str(r[idx["Discipline (Summary)"]] or "").strip()
        disc = disc.split(" - ")[0] if " - " in disc else disc
        if disc not in ("E", "I", "T"):
            continue
        eit += 1
        by_disc[disc] += 1
        state = str(r[idx["Task State"]] or "")
        if state == "Closed":
            closed += 1
    return {"rows": eit, "closed": closed, "by_disc": dict(by_disc)}


def main():
    os.makedirs(DATA_DIR, exist_ok=True)

    ref = None
    try:
        ref = excel_reference()
        print("EXCEL REFERENCE:", json.dumps(ref, ensure_ascii=False), flush=True)
    except Exception as e:
        print("excel ref err:", str(e)[:200], flush=True)

    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            PROFILE_DIR, channel="chrome", headless=False,
            args=["--start-maximized"], viewport=None,
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
        if not wait_products(page, timeout_sec=120):
            print("NO SESSION (long poll)", flush=True)
            for i in range(60):
                if count_products(page) >= 0:
                    break
                page.wait_for_timeout(5000)
        if count_products(page) < 0:
            page.bring_to_front()
            if not wait_products(page, timeout_sec=900):
                ctx.close()
                return 1
        print("LOGGED IN", flush=True)

        out = {"reference_excel": ref}

        # find CPP AGI company id
        companies = call(page, "vCompanies", ["ID", "Name", "ShortName", "Summary", "FullName"], n=1000)
        out["companies"] = {k: (len(v) if k in ("rows",) else v) for k, v in companies.items() if k != "rows"}
        cpp_cid = None
        for r in (companies.get("rows") or []):
            vals = " | ".join(str(v) for v in r.values())
            if "CPP AGI" in vals.upper():
                print("COMPANY MATCH:", json.dumps(r, ensure_ascii=False), flush=True)
                cpp_cid = r.get("ID")
                break
        print("CPP AGI company ID:", cpp_cid, flush=True)

        restrictions = []
        if cpp_cid:
            restrictions = [{"FieldName": "ResponsibleCompanyID", "Term": "=", "ValueCollection": [cpp_cid]}]
        else:
            restrictions = [{"FieldName": "ResponsibleCompany", "Term": "=", "ValueCollection": ["CPP AGI"]}]

        print("pulling CPP AGI tasks...", flush=True)
        data = call(page, "vTasks_TestsPlanned", TASK_FIELDS, restrictions)
        print("pull:", json.dumps({k: (v if k != "rows" else None) for k, v in data.items()}, ensure_ascii=False), flush=True)
        rows = data.get("rows") or []
        print("rows got:", len(rows), flush=True)
        for r in rows[:5]:
            print(" row:", json.dumps({k: (str(v)[:60] if v is not None else None) for k, v in r.items()}, ensure_ascii=False), flush=True)

        out["pulled"] = len(rows)
        out["totalReported"] = data.get("total")
        out["cols"] = data.get("cols")

        if rows:
            ke = Counter()
            closed = Counter()
            closed_by_disc = 0
            closed_by_state_date = 0
            for r in rows:
                disc = str(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline") or "").strip()
                disc = disc.split(" - ")[0] if " - " in disc else disc
                ke[disc] += 1
                state = str(r.get("TaskState") or "")
                if state == "Closed":
                    closed[disc] += 1
                    if r.get("ApprovedDate") not in (None, ""):
                        closed_by_state_date += 1
            out["keyCounts"] = {
                "total": len(rows),
                "E_I_T_total": sum(ke.get(d, 0) for d in ("E", "I", "T")),
                "by_discipline": dict(ke),
                "closed_E_I_T": sum(closed.get(d, 0) for d in ("E", "I", "T")),
                "closed_by_discipline": dict(closed),
                "closed_with_ApprovedDate": closed_by_state_date,
            }
            print("COUNTS:", json.dumps(out["keyCounts"], ensure_ascii=False), flush=True)

            with open(os.path.join(DATA_DIR, "cpp_agi_tasks.json"), "w", encoding="utf-8") as f:
                json.dump({"rows": rows, "meta": out["keyCounts"]}, f, ensure_ascii=False)

        with open(os.path.join(DATA_DIR, "counts_report.json"), "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, default=str, indent=2)
        print("saved -> data/counts_report.json", flush=True)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())