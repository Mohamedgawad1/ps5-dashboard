import re
s = open("sc_pull/online_verified_index.html", encoding="utf-8", errors="ignore").read()

# find the big-number area near 4991 / Total
for pat in ["4991</", "4991 ", ">4,991<", ">4,991 ", "closed_week", "closed_eit", "ITR Closed", "ITR Closed"]:
    print(pat, "->", s.count(pat))

m = re.search(r'.{200}4991.{200}', s)
print("CTX-4991:", (m.group(0)[:420] if m else "NONE"))
print()
# find big numeric cards: look for known labels
for lbl in ["Total", "Closed", "Open", "Progress", "Today", "TOTAL"]:
    idx = [mm.start() for mm in re.finditer('>'+lbl+'<', s)][:2]
    for i in idx[:1]:
        print("lbl", lbl, "->", repr(s[max(0,i-160):i+60])[-220:])
