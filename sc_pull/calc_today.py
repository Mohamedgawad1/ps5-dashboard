import json

rows = json.load(open("sc_pull/data/itr_ps5_rows.json", encoding="utf-8"))

def parse_mdy(v):
    s = str(v)
    datepart = s.split(" ")[0]
    parts = datepart.split("/")
    if len(parts) == 3:
        try:
            m, d, y = int(parts[0]), int(parts[1]), int(parts[2])
            if y < 100:
                y += 2000
            return "%04d-%02d-%02d" % (y, m, d)
        except Exception:
            return None
    return None

closed_today = 0
for r in rows:
    if str(r.get("TaskState")) != "Closed":
        continue
    k = parse_mdy(r.get("ApprovedDate"))
    if k == "2026-09-22":
        closed_today += 1
print("today 2026-09-22 closed (M/D/YYYY parse):", closed_today)

# sanity: show yesterday vs today counts
from collections import Counter
cnt = Counter()
for r in rows:
    if str(r.get("TaskState")) == "Closed":
        k = parse_mdy(r.get("ApprovedDate"))
        if k:
            cnt[k] += 1
for k in sorted(cnt, reverse=True)[:5]:
    print(k, cnt[k])
print("total closed in file:", sum(cnt.values()))