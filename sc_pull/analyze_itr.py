import json
import sys

from playwright.sync_api import sync_playwright

COMMON_FIELDS = [
    "ID", "AssetTag", "Description", "AssetSystemID", "AssetSubsystemID", "TagNumber",
    "TaskCategoryID", "TaskCategory", "TaskSubCategoryID", "TaskSubCategory",
    "TaskStatus", "TaskStatusID", "Status", "StatusID", "Closed", "ClosedBy",
    "DateClosed", "ClosedDate", "IsClosed", "ClosedID", "CompletionStatus",
    "SignoffStatus", "SignOffStatus", "SignedOff", "Inspector", "Discipline",
    "DisciplineID", "ProjectID", "WorkGroupID", "CompanyID", "CompanyInstanceID",
    "RecordCreated", "DateModified", "ModifiedBy", "Number", "Title", "Reference",
    "LoopName", "SystemID", "SubsystemID", "SystemName", "Area", "AreaID",
    "PunchItemNumber", "AssignedTo", "DueDate", "Priority", "Severity",
    "TaskType", "TaskTypeID", "IsITR", "ITRNumber", "TestDescription",
    "InspectionStatus", "InspectionID", "PunchStatus", "OpenDate", "AssetTagFrom",
    "AssetTagTo", "Category", "CategoryID", "Subsystem", "System", "TaskName",
]


def main():
    with sync_playwright() as pw:
        b = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
        ctx = b.contexts[0]
        page = None
        for p in ctx.pages:
            try:
                n = p.evaluate(
                    """() => (typeof switchboardConfiguration!=='undefined'&&switchboardConfiguration.Products)?switchboardConfiguration.Products.length:-1"""
                )
                if n >= 0:
                    page = p
                    break
            except Exception:
                continue
        if not page:
            print("no switchboard page")
            b.close()
            return

        views = sys.argv[1:] or ["vTasks_TestsPlanned", "vAssets_TestForms"]
        for view in views:
            print("=" * 40, view, "=" * 40, flush=True)
            js = """((P) => {
                return new Promise((res) => {
                    dalc.Get(P.v, P.f, [], {Distinct:false, FirstRecordOrdinal:0, NumberOfRecords:20}, "Flat",
                        (ok) => {
                            const vals = (ok && ok.Values) || [];
                            let cols = [];
                            let total = null;
                            if (vals.length) cols = Object.keys(vals[0]).filter(k=>!k.startsWith('__'));
                            if (ok && ok.Request) {
                                const tr = ok.Request.TotalRecords;
                                if (typeof tr !== 'undefined' && tr !== null) total = tr;
                            }
                            res({cols, total, rows: vals.length, sample: vals.slice(0, 5)});
                        },
                        (err) => res({ok:false, err: String((err && err.message) || err)}), false);
                });
            })(%s)""" % json.dumps({"v": view, "f": COMMON_FIELDS})

            try:
                r = page.evaluate(js)
            except Exception as e:
                print("EVAL ERR:", str(e)[:200], flush=True)
                continue
            if r.get("ok") is False:
                print("DALC ERR:", r.get("err"), flush=True)
                continue
            cols = r.get("cols", [])
            print("COLS (%d):" % len(cols))
            for c in cols:
                print("   ", c, flush=True)
            print("TotalRecords:", r.get("total"), flush=True)
            for i, row in enumerate(r.get("sample", [])):
                small = {}
                for k in list(row.keys())[:40]:
                    v = row[k]
                    if v is None:
                        continue
                    if isinstance(v, (dict, list)):
                        v = json.dumps(v, ensure_ascii=False)[:80]
                    small[k] = str(v)[:60]
                print("ROW%d: %s" % (i, json.dumps(small, ensure_ascii=False)), flush=True)
        b.close()


if __name__ == "__main__":
    sys.exit(main())