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
print("live rows:", len(rows))


def disc_code(v):
    if v is None:
        return None
    s = str(v).strip()
    if s in ("E", "I", "T", "M", "S", "P", "B", "H"):
        return s
    for pre in ("E -", "I -", "T -", "M -", "S -", "P -", "B -", "H -"):
        if s.startswith(pre):
            return pre[0]
    return s[:2].upper() if len(s) >= 2 else s.upper()


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
print("excel cpp agi rows:", len(xl))

xl_tags = {}
for r in xl:
    tag = str(r[idx["Asset - Tag"]] or "").strip()
    if tag:
        xl_tags.setdefault(tag, []).append(r)
print("excel unique tags:", len(xl_tags))

# build live EIT rows grouped by tag
live_by_tag = {}
for r in rows:
    tag = str(r.get("AssetTag") or "").strip()
    disc = disc_code(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline"))
    if disc not in ("E", "I", "T"):
        continue
    live_by_tag.setdefault(tag, []).append(r)
print("live EIT unique tags:", len(live_by_tag))

xl_eit_tags = {t for t, rs in xl_tags.items() for r in rs if disc_code(r[idx["Discipline (Summary)"]]) in ("E", "I", "T")}
live_eit_tags = set(live_by_tag.keys())

only_live = live_eit_tags - xl_eit_tags
only_xl = xl_eit_tags - live_eit_tags
both = live_eit_tags & xl_eit_tags
print("xl EIT tags:", len(xl_eit_tags), "live EIT tags:", len(live_eit_tags))
print("only_live tags:", len(only_live), "only_xl tags:", len(only_xl), "both:", len(both))

# analyze only_live rows: what do they have that only_xl/both lack?
def p5(t):
    return t.startswith("PS5") or "PS5" in t

only_live_rows = []
for t in only_live:
    for r in live_by_tag[t]:
        only_live_rows.append(r)
print("\nonly_live rows total:", len(only_live_rows))
cc = Counter()
cc2 = Counter()
for r in only_live_rows:
    cc[str(r.get("TaskState"))] += 1
    cat = str(r.get("TaskCategorySummary"))
    cc2[str(r.get("TaskType"))] += 1
print("only_live by TaskState:", dict(cc))
print("only_live by TaskType:", cc2.most_common(8))

# in only_xl, check what excel has
print("\n--- sample only_xl rows (in excel, not live):", len(only_xl))
for t in list(only_xl)[:8]:
    for r in xl_tags[t][:1]:
        print(" tag:", t, "state:", r[idx["Task State"]], "disc:", r[idx["Discipline (Summary)"]])

# compare closed counts on the SAME tag basis
xl_disc_tag_closed = Counter()
xl_disc_tag = Counter()
for t, rs in xl_tags.items():
    for r in rs:
        d = disc_code(r[idx["Discipline (Summary)"]])
        if d not in ("E", "I", "T"):
            continue
        xl_disc_tag[d] += 1
        if str(r[idx["Task State"]] or "") == "Closed":
            xl_disc_tag_closed[d] += 1
print("\nEXCEL by discipline (all tags):", dict(xl_disc_tag), dict(xl_disc_tag_closed))

# live restricted to tags found in excel
lv = Counter(); lc = Counter()
for t in both:
    for r in live_by_tag[t]:
        d = disc_code(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline"))
        if d not in ("E", "I", "T"):
            continue
        lv[d] += 1
        if str(r.get("TaskState")) == "Closed":
            lc[d] += 1
print("LIVE (tags also in excel) by discipline:", dict(lv), dict(lc), "sum=", sum(lv.values()), "closed=", sum(lc.values()))

# live restricted to tags only-active-in-both AND PS5 by any marker
p5_live_rows = [r for r in rows if p5(str(r.get("AssetTag") or "")) and disc_code(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline")) in ("E", "I", "T")]
print("live EIT tag.startswith PS5 rows:", len(p5_live_rows), "closed:", sum(1 for r in p5_live_rows if str(r.get("TaskState")) == "Closed"))

# what states exist in only_live? maybe they are all 'Not started' type excluded by export