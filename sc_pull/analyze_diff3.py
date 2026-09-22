import json
import os
from collections import Counter

import openpyxl

WS = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

rows = json.load(open(os.path.join(DATA, "cpp_agi_tasks_full.json"), encoding="utf-8"))["rows"]
by_name = {str(r.get("TaskName")): r for r in rows}

wb = openpyxl.load_workbook(os.path.join(WS, "ovTasks_TestsPlanned_1369.xlsx"), read_only=True, data_only=True)
ws = wb["Exported from SC"]
it = ws.iter_rows(values_only=True)
hdr = list(next(it))
idx = {h: i for i, h in enumerate(hdr)}
excel_names = set()
for r in it:
    if str(r[idx["Responsible Company (Summary)"]] or "") == "CPP AGI":
        excel_names.add(str(r[idx["Task ID"]] or ""))
wb.close()
print("excel CPP AGI task names:", len(excel_names))

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

included = []
excluded = []
for r in rows:
    if str(r.get("PhysicalLocationUp1Summary") or "").split(" - ")[0] != "PS5":
        continue
    if str(r.get("TaskName")) in excel_names:
        included.append(r)
    else:
        excluded.append(r)
print("included:", len(included), "excluded:", len(excluded))

def field_counter(group, key):
    c = Counter()
    for r in group:
        v = str(r.get(key) or "")
        if key == "TagPrefix":
            v = str(r.get("AssetTag") or "")[:9]
        c[v] += 1
    return c

for key in ["TagPrefix", "ProcessBreakdown", "TaskType", "TaskCategorySummary", "TaskDisciplineSummary", "TaskModelName"]:
    inc = field_counter(included, key)
    exc = field_counter(excluded, key)
    intop = inc.most_common(8)
    extop = exc.most_common(8)
    print("\n== %s ==" % key)
    print("  INCLUDED top:", json.dumps(intop, ensure_ascii=False))
    print("  EXCLUDED top:", json.dumps(extop, ensure_ascii=False))
    onlyex = {k: v for k, v in exc.items() if k not in inc}
    if onlyex:
        print("  EXCLUDED-ONLY values:", json.dumps(list(onlyex.items())[:8], ensure_ascii=False))

print("\n--- excluded EIT closed count:", sum(1 for r in excluded if dc(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline")) in ("E","I","T") and str(r.get("TaskState"))=="Closed"))
print("--- included EIT closed count:", sum(1 for r in included if dc(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline")) in ("E","I","T") and str(r.get("TaskState"))=="Closed"))