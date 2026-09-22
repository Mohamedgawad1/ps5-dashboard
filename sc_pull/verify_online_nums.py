import re

s = open("sc_pull/online_verified_index.html", encoding="utf-8", errors="ignore").read()
print("refs live_itr_state:", s.count("itr_live_state.json"))
print("refs live_itr.js:", s.count("live_itr.js"))
for pat in ["4,991", "4,966", "5,514", "5,520", "closed", "Closed"]:
    print(repr(pat), "->", s.count(pat))
for m in re.finditer(r'"Total"\s*:\s*(\d+)', s):
    pass
m = re.search(r'closed:\s*(\d+)', s)
print("closed: kw ->", m.group(1) if m else None)
m = re.search(r'ITR.{0,30}closed', s[:8000], re.S)
print("head ctx:", (m.group(0)[:120] if m else None))
m = re.search(r'(\d{1,3}(?:\.\d+)?%?)', s[:6000])
print("first num:", m.group(1) if m else None)
m = re.search(r'((?:Total|total)?\s*genes|Progress[^,;]{0,20})', s, re.S)
print(m.group(0)[:80] if m else "no progress kw")
