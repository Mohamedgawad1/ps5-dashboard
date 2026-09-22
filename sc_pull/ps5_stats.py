import json
from collections import Counter

rows = json.load(open(r"sc_pull\data\cpp_agi_tasks_full.json", encoding="utf-8"))["rows"]


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


ps5 = [r for r in rows if str(r.get("PhysicalLocationUp1Summary") or "").split(" - ")[0] == "PS5"]
print("ALL live PS5 rows:", len(ps5))
print("states:", dict(Counter(str(r.get("TaskState")) for r in ps5)))

eit = [r for r in ps5 if dc(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline")) in ("E", "I", "T")]
cd = Counter(); ct = Counter(); cop = Counter()
for r in eit:
    d = dc(r.get("TaskDisciplineSummary") or r.get("TaskDiscipline"))
    ct[d] += 1
    if str(r.get("TaskState")) == "Closed":
        cd[d] += 1
    else:
        cop[d] += 1
print("PS5 EIT:", len(eit))
print("PS5 EIT by disc:", dict(ct))
print("PS5 EIT closed by disc:", dict(cd), "sum=%d" % sum(cd.values()))
print("PS5 EIT open by disc:", dict(cop), "sum=%d" % sum(cop.values()))
print("closed with ApprovedDate:", sum(1 for r in eit if str(r.get("TaskState")) == "Closed" and r.get("ApprovedDate")))

ids = sorted(int(r["ID"]) for r in ps5 if str(r.get("ID")).isdigit())
print("ps5 id range:", ids[0], "-", ids[-1], "count", len(ids))
print("min id of extras(newer half) distribution:", ids[len(ids) // 3], ids[2 * len(ids) // 3])