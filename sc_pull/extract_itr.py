import json
import os
import sys
import time
from datetime import datetime as _dt

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "profile")
DATA_DIR = os.path.join(BASE_DIR, "data")

SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)
LOGIN_URL = "https://wly04-sc.intergraphsmartcloud.com/Login.aspx"

FOCUS_VIEWS = [
    "vAssets_TestForms",
    "vTasks_TestsPlanned",
    "vPunchlist",
    "vInspectionNonCompliances",
]

BIG_FIELD_CANDIDATES = [
    "ID", "AssetTag", "AssetId", "Description", "AssetSystemID", "AssetSubsystemID",
    "TagNumber", "TaskCategoryID", "TaskCategory", "TaskSubCategoryID", "TaskSubCategory",
    "TaskStatus", "TaskStatusID", "Status", "StatusID", "Closed", "ClosedBy",
    "DateClosed", "ClosedDate", "IsClosed", "ClosedID", "CompletionStatus",
    "SignoffStatus", "SignOffStatus", "SignedOff", "Inspector", "Discipline",
    "DisciplineID", "ProjectID", "WorkGroupID", "CompanyID", "CompanyInstanceID",
    "RecordCreated", "DateModified", "ModifiedBy", "Number", "Title", "Reference",
    "LoopName", "SystemID", "SubsystemID", "SystemName", "Area", "AreaID",
    "PunchItemNumber", "AssignedTo", "DueDate", "Priority", "Severity",
    "TaskType", "TaskTypeID", "IsITR", "ITRNumber", "TestDescription",
    "InspectionStatus", "InspectionID", "PunchStatus", "OpenDate", "TaskName",
    "TaskNumber", "TestFormsCount", "TestFormID", "TestFormNumber", "TestStatus",
    "CertificateStatus", "Approved", "ApprovedBy", "ApprovedDate", "SignedBy", "SignDate",
    "TestDate", "PlannedDate", "ExecutionDate", "Executed", "TaskResponsibleCompanyID",
    "TaskResponsibleCompany", "ResponsibleCompanyID", "ResponsibleCompany",
    "TaskDueDateTimeStamp", "MSF", "PageNumber", "Tabs", "TaskModelID",
    "StatusColor", "StatusRank", "PhaseNumber", "Phase", "Milestone", "PriorityName",
    "PredecessorTypeName", "CostRollingStatus", "IsROT", "ROTDate", "DSSColor",
    "NoVerificationYet", "Revision", "Version", "Mode", "Flags", "SortIndex", "RequestStatus",
]

DEFAULT_FILTER_FIELDS = [
    "ProjectID", "CompanyInstanceID", "InstallationID",
]

BROAD_CANDIDATES = list(dict.fromkeys(["RowNumber", "AssetTag", "Description"] + BIG_FIELD_CANDIDATES))


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


def get_controller_deep(page, view):
    js = """((VIEW) => {
        const out = {view: VIEW};
        try {
            const ctl = SmartCompletions.data.getViewController(VIEW);
            out.hasController = !!ctl;
            if (ctl) {
                out.controllerKeys = Object.keys(ctl);
                const bh = ctl.Behaviors;
                if (bh) {
                    let names = null;
                    const n = bh.Names || bh;
                    if (Array.isArray(n)) names = n;
                    else if (n && typeof n === 'object') { names = []; for (const k in n) names.push(k); }
                    out.behaviorNames = names;
                }
                const gb = ctl.getBehavior;
                if (typeof gb === 'function') {
                    out.behaviors = {};
                    for (const name of (out.behaviorNames || []).slice(0, 20)) {
                        try {
                            const b = gb.call ? gb.call(ctl, name) : gb(name);
                            out.behaviors[name] = sanitize(b);
                        } catch(e) { out.behaviors[name] = {__err: String(e).slice(0,200)}; }
                    }
                }
            }
        } catch(e) { out.controllerError = String(e).slice(0,500); }
        return out;

        function sanitize(x, depth) {
            depth = depth || 0;
            if (depth > 6) return {__depth: depth};
            if (x === null || x === undefined) return x;
            const t = typeof x;
            if (t === 'function' || t === 'symbol') return {__fn: true};
            if (t !== 'object') return x;
            if (Array.isArray(x)) {
                if (x.length > 40) return {__arrLen: x.length, head: x.slice(0, 3).map(v => sanitize(v, depth+1))};
                if (x.length === 0) return [];
                return x.slice(0, 40).map(v => sanitize(v, depth+1));
            }
            const out = {};
            let keys = Object.keys(x);
            if (keys.length > 60) return {__objKeys: keys.slice(0, 60)};
            for (const k of keys) {
                try { out[k] = sanitize(x[k], depth+1); } catch(e) { out[k] = {__err: String(e).slice(0,100)}; }
            }
            return out;
        }
    })(%s)""" % json.dumps(view)
    return page.evaluate(js)


def sample_pull(page, view, fields, n=5, restrictions=None):
    js = """((P) => {
        return new Promise((res) => {
            const restr = P.r || [];
            dalc.Get(P.v, P.f, restr, {Distinct:false, FirstRecordOrdinal:0, NumberOfRecords:P.n}, "Flat",
                (ok) => {
                    let cols = [], rows = 0, total = null, sample = null;
                    if (ok && ok.Values && ok.Values.length) {
                        cols = Object.keys(ok.Values[0]).filter(k=>!k.startsWith('__'));
                        rows = ok.Values.length;
                        sample = JSON.parse(JSON.stringify(ok.Values.slice(0, P.n)));
                    }
                    if (ok && ok.Request) {
                        const tr = ok.Request.TotalRecords;
                        if (typeof tr !== "undefined" && tr !== null) total = tr;
                    }
                    res({ok:true, cols, rows, total, sample});
                },
                (err) => res({ok:false, err: String((err && err.message) || err)}), false);
        });
    })(%s)""" % json.dumps({"v": view, "f": fields, "n": n, "r": restrictions or []})
    return page.evaluate(js)


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
            print("nav:", str(e)[:120], flush=True)

        if not wait_products(page, timeout_sec=25):
            print("=== الجلسة مش حية هنا؛ لازم تسجيل دخول في النافذة اللي فتحت ===", flush=True)
            try:
                page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
            except Exception:
                pass
            page.bring_to_front()
            print(">>> سجل دخولك وانزل للمنصة (أنا مستني لحد ما تشوف الأزرار)", flush=True)
            if not wait_products(page, timeout_sec=900):
                print(">>> مهلة الدخول خلصت بدون دخول.", flush=True)
                ctx.close()
                return 1

        print("LOGGED IN", flush=True)
        ctx_info = page.evaluate(
            """() => {
            const d = typeof dalc!=='undefined' && dalc.user ? dalc.user : null;
            return {
                ID: d && d.ID, Resource: d && d.Resource, loggedIn: d && d.loggedIn,
                pj: (typeof localStorage!=='undefined') ? {pj: localStorage.getItem('ibspj'), name: localStorage.getItem('ibspjname'), ci: localStorage.getItem('ibsCIid')} : null,
            };
        }"""
        )
        print("context:", json.dumps(ctx_info, ensure_ascii=False), flush=True)

        schema = {}
        for view in FOCUS_VIEWS:
            print("controller deep:", view, flush=True)
            try:
                schema[view] = get_controller_deep(page, view)
            except Exception as e:
                schema[view] = {"error": str(e)[:200]}
                print("  controller ERR:", str(e)[:150], flush=True)

            print("sample pull:", view, flush=True)
            try:
                res = sample_pull(page, view, BROAD_CANDIDATES, n=3)
                schema[view + "_sample"] = res
                print("  cols:", res.get("cols"), "total:", res.get("total"), flush=True)
                for i, row in enumerate(res.get("sample") or []):
                    keep = {k: str(v)[:100] for k, v in (row or {}).items()}
                    print("   row%d: %s" % (i, json.dumps(keep, ensure_ascii=False)), flush=True)
            except Exception as e:
                print("  sample ERR:", str(e)[:150], flush=True)

        out_path = os.path.join(DATA_DIR, "schema_behaviors.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(schema, f, ensure_ascii=False, indent=2)
        print("saved ->", out_path, flush=True)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())