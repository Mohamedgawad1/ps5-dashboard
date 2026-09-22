import json
import os
from collections import Counter

import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WS_DIR = os.path.dirname(BASE_DIR)
DATA_DIR = os.path.join(BASE_DIR, "data")

with open(os.path.join(DATA_DIR, "cpp_agi_tasks_full.json"), "r", encoding="utf-8") as f:
    live = json.load(f)
rows = live["rows"]
live_by_id = {str(r.get("TaskName")): r for r in rows}
print("live rows:", len(rows), "live by name:", len(live_by_id))

wb = openpyxl.load_workbook(os.path.join(WS_DIR, "ovTasks_TestsPlanned_1369.xlsx"), read_only=True, data_only=True)
ws = wb["Exported from SC"]
it = ws.iter_rows(values_only=True)
header = list(next(it))
idx = {h: i for i, h in enumerate(header)}
xl = []
for r in it:
    if str(r[idx["Responsible Company (Summary)"]] or "") == "CPP AGI":
        xl.append(r)
wb.close()
print("excel CPP AGI rows:", len(xl))

xl_ids = set()
xl_ids_not_in_live = []
for r in xl:
    tid = str(r[idx["Task ID"]] or "").strip()
    if tid:
        xl_ids.add(tid)
        if tid not in live_by_id:
            xl_ids_not_in_live.append(r)
print("excel unique Task Names:", len(xl_ids), "not found in live:", len(xl_ids_not_in_live))

# states of excel IDs as found in live
st_count = Counter()
loc = Counter()
pb_set = Counter()
for tid in xl_ids:
    r = live_by_id[tid]
    st_count[str(r.get("TaskState"))] += 1
    loc[str(r.get("PhysicalLocationUp1Summary") or "")] += 1
    pb_set[str(r.get("ProcessBreakdown") or "")] += 1
print("\nstates of exported IDs (live):", dict(st_count))
print("physical location of exported IDs top:", loc.most_common(6))
print("process breakdown top:", pb_set.most_common(4))

# excel rows PS5-ness via direct cols
xl_ps5_final = 0
for r in xl:
    s = " | ".join(str(x) for x in (r[idx["Asset - Tag"]], r[idx["Systemization - Subsystem (Summary)"]], r[idx["Physical Location - Area (Summary)"]]))
    if "PS5" in s.upper() or "P5" in s.upper():
        xl_ps5_final += 1
print("\nexcel rows with PS5/P5 in tag/subsys/area:", xl_ps5_final, "of", len(xl))

# Now: which live rows (whole company) are NOT exported -> what distinguishes?
# Compare exported IDs vs all live IDs on: station, TaskState, TaskType, Category
live_st = Counter(); live_typed = Counter(); live_station = Counter()
for r in rows:
    live_st[str(r.get("TaskState"))] += 1
    live_typed[str(r.get("TaskType"))] += 1
    loc = str(r.get("PhysicalLocationUp1Summary") or "")
    live_station[loc.split(" - ")[0] if loc else r.get("ProcessBreakdown", "")] += 1
print("\nLIVE all-company states:", dict(live_st))
print("LIVE all-company types:", live_typed.most_common(6))

# live rows on excel IDs - discipline split vs excel discipline split
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

xd = Counter(); xdc = Counter()
for r in xl:
    d = dc(r[idx["Discipline (Summary)"]])
    if d in ("E", "I", "T"):
        xd[d] += 1
        if str(r[idx["Task State"]] or "") == "Closed":
            xdc[d] += 1
print("\nEXCEL EIT by disc:", dict(xd), "closed:", dict(xdc))

# live rows whose ID is in excel, EIT
ld = Counter(); ldc = Counter()
for tid in xl_ids:
    r = live_by_id[tid]
    d = dc(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline"))
    if d in ("E", "I", "T"):
        ld[d] += 1
        if str(r.get("TaskState")) == "Closed":
            ldc[d] += 1
print("LIVE(on excel IDs) EIT by disc:", dict(ld), "closed:", dict(ldc))

# Which live ID rows are NOT exported? maybe restrict on the view: the export might live in a DIFFERENT view
# (vTasks_All / planned vs tests). Count overlaps fully.
print("\nrows not in excel -> sample 20 (id, tag, state, loc, type, cat):")
shown = 0
for r in rows:
    if str(r.get("ID")) in xl_ids:
        continue
    if shown >= 20:
        break
    loc = str(r.get("PhysicalLocationUp1Summary") or "")
    if loc.split(" - ")[0] != "PS5":
        continue
    print(" ", r.get("ID"), r.get("AssetTag"), r.get("TaskState"), loc[:28], r.get("TaskType"), str(r.get("TaskCategorySummary"))[:18])
    shown += 1