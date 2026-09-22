import re
s = open("index.html", encoding="utf-8", errors="ignore").read()

# Find KPI card markup
for m in re.finditer(r'<div class="kpi[^"]*">', s):
    i = m.start()
    block = s[i:i+320]
    print(re.sub(r'\s+',' ', block)[:320])
    print('----')
    if sum(1 for _ in re.finditer(r'<div class="kpi', s[:i])) >= 6:
        break

print('==== total id-based spans ====')
for mm in re.finditer(r'<span[^>]*id="(totalClosed|todayClosed|totalPct|todayPct|closedTotal|todayCount|progressBar|totalPctBar|todayClosedSpan)"[^>]*>[^<]*</span>', s):
    print(mm.group(1), '->', mm.group(0)[:120])

print('==== #totalPct occurrences ====')
print(s.count('id="totalPct"'), s.count('totalPct'), s.count('id="todayClosed"'), s.count('id="totalClosed"'))
