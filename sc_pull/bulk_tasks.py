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

GRID_FIELDS = [
    "WeekOverdue", "ID", "TaskName", "TaskID", "AssetTag", "TaskDescription",
    "IsException", "ExecutionType", "TaskDiscipline", "TaskType", "ResponsibleCompany",
    "PDFFileID", "HasPackage", "HasPLs", "JobName", "AssetPackName", "LoopName",
    "WorkBreakdownUp2Name", "WorkBreakdownUp1Name", "Discipline", "AssetType",
    "CertificateCategorySummary", "ActualEndDate", "CompletedBy", "ApprovedDate",
    "ApprovedBy", "ProcessBreakdownUp1Summary", "ProcessBreakdown", "ProcessBreakdownCustodyID",
]

PAGE = 2000
MAX_ROWS = 250000
CALL_TIMEOUT_MS = 120000


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


PULL_JS = """((P) => {
    return new Promise((res) => {
        let settled = false;
        const timer = setTimeout(() => { if (!settled) { settled = true; res({timeout: true}); } }, P.timeoutMs);
        dalc.Get(P.v, P.f, [], {Distinct:false, FirstRecordOrdinal: P.off, NumberOfRecords: P.size}, "Flat",
            (ok) => {
                if (settled) return;
                settled = true; clearTimeout(timer);
                let cols = [], rows = [], total = null;
                const vals = (ok && ok.Values) || [];
                if (vals.length) cols = Object.keys(vals[0]).filter(k=>!k.startsWith('__'));
                rows = vals && vals.length ? JSON.parse(JSON.stringify(vals)) : [];
                if (ok && ok.Request) {
                    const tr = ok.Request.TotalRecords;
                    if (typeof tr !== "undefined" && tr !== null) total = tr;
                }
                res({cols: cols, rows: rows, total: total, n: rows.length});
            },
            (err) => { if (!settled) { settled = true; clearTimeout(timer); res({err: String((err && err.message) || err)}); } }, false);
    });
})"""


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
        except Exception:
            pass
        if not wait_products(page, timeout_sec=40):
            print("NO SESSION", flush=True)
            page.bring_to_front()
            if not wait_products(page, timeout_sec=900):
                ctx.close()
                return 1

        # sample first
        s = page.evaluate("(%s)(%s)" % (
            PULL_JS,
            json.dumps({"v": "vTasks_TestsPlanned", "f": GRID_FIELDS, "off": 0, "size": 3, "timeoutMs": 120000}),
        ))
        print("SAMPLE:", json.dumps({k: (v if k != "rows" else None) for k, v in s.items()}, ensure_ascii=False), flush=True)
        if s.get("n", 0) == 0:
            print("empty cols:", s.get("cols"), "err:", s.get("err"), flush=True)
            ctx.close()
            return 1
        for row in s.get("rows", [])[:3]:
            print("row:", json.dumps({k: (str(v)[:120] if v is not None else None) for k, v in row.items()}, ensure_ascii=False), flush=True)

        all_rows = []
        total = s.get("total")
        off = 0
        while True:
            r = page.evaluate("(%s)(%s)" % (
                PULL_JS,
                json.dumps({"v": "vTasks_TestsPlanned", "f": GRID_FIELDS, "off": off, "size": PAGE, "timeoutMs": CALL_TIMEOUT_MS}),
            ))
            n = r.get("n", 0)
            if r.get("timeout"):
                print("timeout at off=%d; aborting" % off, flush=True)
                break
            if r.get("err"):
                print("err at off=%d: %s" % (off, r["err"][:150]), flush=True)
                break
            if total is None and r.get("total") is not None:
                total = r["total"]
            all_rows.extend(r.get("rows", []))
            print("pulled off=%d n=%d total=%s accumulated=%d" % (off, n, total, len(all_rows)), flush=True)
            if n < PAGE:
                break
            off += PAGE
            if len(all_rows) >= MAX_ROWS:
                print("MAX_ROWS reached", flush=True)
                break

        out = {"total": total, "pulled": len(all_rows), "fields": GRID_FIELDS, "rows": all_rows}
        with open(os.path.join(DATA_DIR, "tasks_bulk.json"), "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False)
        print("saved data/tasks_bulk.json pulled=%d total=%s" % (len(all_rows), total), flush=True)
        if all_rows:
            only = {k: all_rows[0][k] for k in all_rows[0]}
            print("first row keys:", list(all_rows[0].keys()), flush=True)

        ctx.close()


if __name__ == "__main__":
    sys.exit(main())