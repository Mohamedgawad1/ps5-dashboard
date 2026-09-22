import json
import os
import sys
import time

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "profile")
OUT_DIR = os.path.join(BASE_DIR, "data")

CDP_URL = "http://127.0.0.1:9222"
SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)

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
    "LoopName", "SystemID", "SubsystemID", "Category", "CategoryID",
    "PunchItemNumber", "AssignedTo", "DueDate", "Priority", "Severity",
    "TaskType", "TaskTypeID", "IsITR", "ITRNumber", "TestDescription",
    "InspectionStatus", "InspectionID", "PunchStatus", "OpenDate",
]

EXTRA_FIELDS = {
    "vAssets_TestForms": [
        "TaskTAT", "TaskAT", "NextTestDate", "FreqValue", "Frequency", "Units", "Method",
        "RoutineStatus", "Result1", "Result2", "Result3", "Result4",
    ],
    "vTasks_TestsPlanned": [
        "TaskTAT", "TaskAT", "NextTestDate", "FreqValue", "Frequency", "Units", "Method",
    ],
    "vPunchlist": [
        "PunchItem", "ResponsibleCompanyID", "ResponsibleCompany", "SupervisorID",
        "WorkOrder", "SectionID",
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
        browser = pw.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]

        import time as _t
        page = None
        for p in ctx.pages:
            try:
                if count_products(p) >= 0:
                    page = p
                    break
            except Exception:
                continue
        if page is None:
            for p in ctx.pages:
                try:
                    has_dalc = p.evaluate(
                        """() => (typeof dalc !== 'undefined' && dalc.user && typeof dalc.Get === 'function')"""
                    )
                    if has_dalc and "intergraphsmartcloud" in p.url:
                        page = p
                        break
                except Exception:
                    continue
        if page is None:
            for p in ctx.pages:
                if p.url in ("", "about:blank", "chrome://newtab/"):
                    page = p
                    break
        if page is None:
            page = ctx.new_page()

        if count_products(page) < 0:
            try:
                page.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=90000)
            except Exception as e:
                print("nav-warning:", str(e)[:120], flush=True)

        logged_in = False
        deadline = time.time() + 900
        while time.time() < deadline:
            if count_products(page) >= 0:
                logged_in = True
                break
            print("... بانتظار تسجيل الدخول في النافذة المفتوحة (كون صح انك شايف صفحة الأزرار)...", flush=True)
            time.sleep(5)

        if not logged_in:
            print(">>> لم يكتمل الدخول خلال المهلة.", flush=True)
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
                }
                return info;
            }"""
        )
        print("context:", json.dumps(context_info, ensure_ascii=False)[:3000], flush=True)

        results = {}
        for view in FOCUS_VIEWS:
            fields = list(dict.fromkeys(COMMON_FIELDS + EXTRA_FIELDS.get(view, [])))
            print("pulling:", view, flush=True)
            res = page.evaluate(js_get_all(view, fields))
            rows = res.get("rows", [])
            results[view] = {
                "cols": res.get("cols", []),
                "rowCount": len(rows),
                "totalReported": res.get("total"),
                "error": res.get("error"),
            }
            with open(os.path.join(OUT_DIR, view + ".json"), "w", encoding="utf-8") as f:
                json.dump({"cols": res.get("cols", []), "rows": rows}, f, ensure_ascii=False)
            print("   -> rows:", len(rows), "cols:", results[view]["cols"], "err:", res.get("error"), flush=True)

        with open(os.path.join(OUT_DIR, "context.json"), "w", encoding="utf-8") as f:
            json.dump({"context": context_info, "results": results}, f, ensure_ascii=False, indent=2)

        print("\n=== PULL DONE ===", flush=True)
        for view, r in results.items():
            print("  %s: %d rows | %s" % (view, r["rowCount"], ", ".join(r["cols"] or [])), flush=True)
        browser.close()


if __name__ == "__main__":
    sys.exit(main())