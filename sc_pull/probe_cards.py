import re
s = open("sc_pull/online_verified_index.html", encoding="utf-8", errors="ignore").read()
for pat in ['4,991', '4,966', '5,514', '5,520', '5,522', '30,800', 'Total']:
    print(pat, '->', s.count(pat))
print('---- find KPI big-number card ----')
for m in re.finditer(r'.{80}(?:4,?991|4,?966|Total ITR|ITR Closed).{90}', s):
    print(repr(m.group(0)[:260]))
    print('~~~~')
    if m.start() > 2000000: break
