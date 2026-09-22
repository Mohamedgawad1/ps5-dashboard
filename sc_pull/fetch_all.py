import json
import os
import re
import sys

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "profile")
OUT_DIR = os.path.join(BASE_DIR, "data")

SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)
LOGIN_URL = "https://wly04-sc.intergraphsmartcloud.com/Login.aspx"

PAGE_SIZE = 2000
MAX_ROWS = 400000

FOCUS_VIEWS = [
    "vAssets_TestForms",
    "vTasks_TestsPlanned",
    "vPunchlist",
    "vInspectionNonCompliances",
]

COMMON_FIELDS = [
    "ID", "AssetTag", "Description", "AssetSystemID", "AssetSubsystemID", "TagNumber",
    "TaskCategoryID", "TaskCategory", "TaskSubCategoryID", "TaskSubCategory",
    "TaskStatus", "TaskStatusID", "Status", "StatusID", "Closed", "ClosedBy",
    "DateClosed", "ClosedDate", "IsClosed", "ClosedID", "CompletionStatus",
    "SignoffStatus", "SignOffStatus", "SignedOff", "Inspector", "Discipline",
    "DisciplineID", "ProjectID", "WorkGroupID", "CompanyID", "CompanyInstanceID",
    "RecordCreated", "DateModified", "ModifiedBy", "Number", "Title", "Reference",
    "Description1", "Description2", "TCategory", "TSubCategory", "InstArea",
    "DocTypeName", "FileName", "LoopName", "SystemID", "SubsystemID", "Category",
    "CategoryID", "PunchItemNumber", "AssignedTo", "DueDate", "Priority", "Severity",
    "PriorityID", "SeverityID", "TaskType", "TaskTypeID", "IsITR", "ITRNumber",
    "TestForm", "TestDescription", "InspectionStatus", "Inspection", "InspectionID",
]

EXTRA_FIELDS = {
    "vAssets_TestForms": [
        "TaskTAT", "TaskAT", "PreviewFileID", "Path", "NextTestDate", "FreqValue", "Frequency",
        "Units", "Method", "RoutineStatus", "Result1", "Result2", "Result3", "Result4",
    ],
    "vTasks_TestsPlanned": [
        "TaskTAT", "TaskAT", "NextTestDate", "FreqValue", "Frequency", "Units", "Method",
    ],
    "vPunchlist": [
        "PunchItemNumber", "PunchItem", "ResponsibleCompanyID", "ResponsibleCompany",
        "Discipline", "SupervisorID", "WorkOrder", "SectionID", "PunchStatus", "OpenDate",
    ],
    "vInspectionNonCompliances": [
        "NCNumber", "NonConformance", "CorrectiveAction", "ResponsibleCompanyID", "TypeID",
    ],
}


def count_products(page):
    try:
        return page.evaluate(
            """() => (typeof switchboardConfiguration!=="undefined" && switchboardConfiguration.Products
            && switchboardConfiguration.Products.length) ? switchboardConfiguration.Products.length : -1"""
        )
    except Exception:
        return -1


def wait_products(page, timeout_sec):
    import time
    start = time.time()
    while time.time() - start < timeout_sec:
        if count_products(page) >= 0:
            return True
        try:
            page.wait_for_timeout(2000)
        except Exception:
            pass
    return False


def ensure_logged_in(page):
    page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
    if wait_products(page, timeout_sec=20):
        return True
    page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
    page.bring_to_front()

    creds = {}
    creds_path = os.path.join(BASE_DIR, "creds.env")
    if os.path.exists(creds_path):
        with open(creds_path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if "=" in line and not line.startswith("#"):
                    k, v = line.split("=", 1)
                    creds[k.strip()] = v.strip().strip('"').strip("'")

    if creds.get("SC_USER") and creds.get("SC_PASS"):
        try:
            page.fill("#txtUserName", creds["SC_USER"], timeout=5000)
            page.fill("#txtUserPass", creds["SC_PASS"], timeout=5000)
            page.click("#btnLogin", timeout=5000)
        except Exception:
            pass

    print("=" * 60)
    print(">>> سجل الدخول بنفسك في النافذة المفتوحة (يوزر/باسورد أو التحقق اللي بتعمله)")
    print(">>> وارجع للمنصة. محاولة واحدة فقط — مفيش إعادة محاولة.")
    print("=" * 60)
    return wait_products(page, timeout_sec=900)


def js_get_all(view, fields):
    return """(async () => {
        const out = {rows: [], cols: [], total: null, error: null};
        const PAGE = PAGE_SIZE;
        let off = 0;
        try {
            for (let i = 0; i < 400; i++) {
                const ok = await new Promise((res, rej) => {
                    dalc.Get(VIEW, FIELDS, [], {Distinct:false, FirstRecordOrdinal: off, NumberOfRecords: PAGE}, "Flat",
                        (r) => res(r), (err) => rej(err), false);
                });
                const vals = (ok && ok.Values) || [];
                if (!out.cols.length && vals.length) {
                    out.cols = Object.keys(vals[0]).filter(k => !k.startsWith('__'));
                }
                if (ok && ok.Request) {
                    const tr = ok.Request.TotalRecords;
                    if (typeof tr !== 'undefined' && tr !== null) out.total = tr;
                }
                out.rows.push.apply(out.rows, vals);
                if (vals.length < PAGE) break;
                off += PAGE;
                if (out.rows.length >= MAXROWS) break;
            }
        } catch (e) { out.error = String(e && e.message || e); }
        return out;
    })()""".replace("VIEW", json.dumps(view)).replace("FIELDS", json.dumps(fields)).replace("PAGE_SIZE", str(PAGE_SIZE)).replace("MAXROWS", str(MAX_ROWS))


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            PROFILE_DIR, channel="chrome", headless=False, args=["--start-maximized"], viewport=None
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        if not ensure_logged_in(page):
            print(">>> لم يكتمل الدخول. هقفل.")
            ctx.close()
            return 1

        context_info = page.evaluate(
            """() => {
                const info = {url: location.href};
                if (typeof localStorage !== 'undefined') {
                    info.ibspj = localStorage.getItem('ibspj');
                    info.ibspjname = localStorage.getItem('ibspjname');
                    info.ibsCIid = localStorage.getItem('ibsCIid');
                    info.ibsCIname = localStorage.getItem('ibsCIname');
                }
                if (typeof dalc !== 'undefined' && dalc.user) {
                    info.userID = dalc.user.ID;
                    info.resource = dalc.user.Resource;
                    info.companyID = dalc.user.CompanyID;
                    info.pj = JSON.parse(JSON.stringify(dalc.user.pj || null));
                    info.loc = JSON.parse(JSON.stringify(dalc.user.loc || null));
                }
                return info;
            }"""
        )
        print("context:", json.dumps(context_info, ensure_ascii=False)[:2000])

        results = {}
        for view in FOCUS_VIEWS:
            fields = list(dict.fromkeys(COMMON_FIELDS + EXTRA_FIELDS.get(view, [])))
            print("pulling:", view, "(%d candidate fields)" % len(fields))
            res = page.evaluate(js_get_all(view, fields))
            results[view] = {
                "cols": res.get("cols", []),
                "rowCount": len(res.get("rows", [])),
                "totalReported": res.get("total"),
                "error": res.get("error"),
            }
            with open(os.path.join(OUT_DIR, view + ".json"), "w", encoding="utf-8") as f:
                json.dump({"cols": res.get("cols", []), "rows": res.get("rows", [])}, f, ensure_ascii=False)
            print("   -> rows:", results[view]["rowCount"], "total:", res.get("total"), "err:", res.get("error"))

        with open(os.path.join(OUT_DIR, "context.json"), "w", encoding="utf-8") as f:
            json.dump({"context": context_info, "results": results}, f, ensure_ascii=False, indent=2)

        print("\n=== PULL DONE ===")
        for view, r in results.items():
            print("  %s: %d rows | cols: %s" % (view, r["rowCount"], ", ".join(r["cols"] or [])))
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())