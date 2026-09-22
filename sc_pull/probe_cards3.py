import re, html as H
s = open("sc_pull/online_verified_index.html", encoding="utf-8", errors="ignore").read()

# Find where the big card numbers live in HTML (text nodes near 4,991 and Today/labels)
for kw in ["4,991", "4,992", "Closed ITR", "ITR Closed", "Closed", "Today", "TOTAL"]:
    idxs = [m.start() for m in re.finditer(re.escape(kw), s)]
    print("== kw", kw, "count", len(idxs))
    for i in idxs[:2]:
        seg = re.sub(r'\s+', ' ', s[max(0,i-260):i+120])
        print("   ...", seg[-360:])
        print("   ----")
    print()

# structure around daily_total (the +1 area)
m = re.search(r'.{0,10}daily_total.{0,80}', s)
print("daily_total ctx:", (m.group(0) if m else None))
