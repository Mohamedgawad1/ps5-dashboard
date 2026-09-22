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
LOG = os.path.join(BASE_DIR, "kpi_out.log")


class Tee:
    def __init__(self, *streams):
        self.streams = streams

    def write(self, s):
        for st in self.streams:
            try:
                st.write(s)
                st.flush()
            except Exception:
                pass

    def flush(self):
        for st in self.streams:
            try:
                st.flush()
            except Exception:
                pass


def pr(*a):
    s = " ".join(str(x) for x in a)
    try:
        print(s, flush=True)
    except Exception:
        pass
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(time.strftime("[%H:%M:%S] ") + s + "\n")
    except Exception:
        pass

SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)

CPP_CID = 1021
PAGE = 4000

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
    urls = [
        "https://wly04-sc.intergraphsmartcloud.com/ISC/Home/Login.aspx",
        "https://wly04-sc.intergraphsmartcloud.com/ISC/SSO/Login.aspx",
        "https://wly04-sc.intergraphsmartcloud.com/ISC/Login.aspx",
    ]
    for u in urls:
        try:
            page.goto(u, wait_until="domcontentloaded", timeout=45000)
        except Exception as e:
            pr("goto", u, "err", str(e)[:80])
            continue
        page.wait_for_timeout(1500)
        try:
            pw = page.query_selector("input[type=password]")
        except Exception:
            pw = None
        if pw:
            pr("LOGIN FORM at", u)
            try:
                user = page.query_selector("input[type=text], input[name*=User], input[id*=User]")
                if user:
                    user.fill(creds.get("SC_USER", ""))
                pw.fill(creds.get("SC_PASS", ""))
                btn = page.query_selector("input[type=submit], button[type=submit]")
                if btn:
                    btn.click()
                else:
                    pw.press("Enter")
                page.wait_for_timeout(8000)
                return True
            except Exception as e:
                pr("autologin err", str(e)[:150])
    return False


EXPORT_COLS = [
    "Task ID", "Asset - Tag", "Asset Pack ID", "Cable Pack ID", "Pipe Pack ID", "Loop Name",
    "Description", "Category (Summary)", "Discipline", "Task Type (Name)",
    "Discipline (Summary)", "Asset - Description", "Systemization - Process Plant ID",
    "Responsible Company (Summary)", "Actual Duration", "Scheduled Duration", "Task Duration",
    "Title - Actual MH/Person", "Title - Planned MH/Person", "Field005 (Custom)",
    "Physical Location - Area (Summary)", "Systemization - Subsystem (Summary)",
    "Closing Date", "Subsystem ID", "Task Priority", "Priority", "Test Form - Description",
    "Asset Discipline", "Task State", "Subsystem Priority", "Task Model", "Closed By",
]


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
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
        if not wait_products(page, timeout_sec=45):
            pr("NO SESSION -> try auto-login")
            auto_login(page)
            page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
            if not wait_products(page, timeout_sec=45):
                pr("NO SESSION after auto-login (long poll)")
                for i in range(60):
                    if count_products(page) >= 0:
                        break
                    page.wait_for_timeout(5000)
        if count_products(page) < 0:
            page.bring_to_front()
            pr("waiting for manual login...")
            if not wait_products(page, timeout_sec=900):
                ctx.close()
                return 1
        pr("LOGGED IN")

        restrictions = [
            {"FieldName": "ResponsibleCompanyID", "Term": "=", "ValueCollection": [CPP_CID]},
            {"FieldName": "PhysicalLocationUp1Summary", "Term": "like", "ValueCollection": ["PS5 -"]},
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
        pr("PS5 CPP AGI rows:", len(all_rows))

        eit = [r for r in all_rows if dc(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline")) in ("E", "I", "T")]
        by = Counter(); closed = Counter(); opened = Counter()
        for r in eit:
            d = dc(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline"))
            by[d] += 1
            st = str(r.get("TaskState"))
            if st == "Closed":
                closed[d] += 1
            else:
                opened[d] += 1
        res = {
            "scope": "PS5 / CPP AGI / vTasks_TestsPlanned",
            "as_of": time.strftime("%Y-%m-%d %H:%M:%S"),
            "ps5_total_rows": len(all_rows),
            "eit_total": len(eit),
            "eit_closed": sum(closed.values()),
            "eit_open": sum(opened.values()),
            "eit_by_discipline": dict(by),
            "eit_closed_by_discipline": dict(closed),
            "eit_open_by_discipline": dict(opened),
        }
        pr("RESULT:", json.dumps(res, ensure_ascii=False))

        with open(os.path.join(DATA_DIR, "itr_kpi.json"), "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=2)
        with open(os.path.join(DATA_DIR, "itr_ps5_rows.json"), "w", encoding="utf-8") as f:
            json.dump(eit, f, ensure_ascii=False, indent=2)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Exported from SC"
        ws.append(EXPORT_COLS)
        for r in sorted(eit, key=lambda x: str(x.get("TaskName") or "")):
            ws.append([
                r.get("TaskName"), r.get("AssetTag"), None, None, None, r.get("LoopName"),
                r.get("Description"), r.get("TaskCategorySummary"), r.get("TaskDiscipline"),
                r.get("TaskType"), r.get("TaskDisciplineSummary"), r.get("AssetDescription"), None,
                r.get("ResponsibleCompany"), None, None, None, None, None, None,
                r.get("PhysicalLocationUp1Summary"), r.get("ProcessBreakdown"),
                r.get("ApprovedDate"), r.get("ProcessBreakdownIdentifier"),
                r.get("TaskPrioritySummary"), None, r.get("DocumentDescription"),
                None, r.get("TaskState"), None, r.get("TaskModelName"), r.get("ApprovedBy"),
            ])
        out = os.path.join(BASE_DIR, "Live_PS5_CPP_AGI_EIT_Tasks.xlsx")
        wb.save(out)
        pr("saved:", out)
        ctx.close()

    pr("DONE")
    return 0


if __name__ == "__main__":
    sys.exit(main())