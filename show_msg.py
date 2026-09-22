import json

with open('index.html', encoding='utf-8') as f:
    html = f.read()

start = html.find('const ITR = ')
s = html.index('{', start)
d = 0; in_str = False; esc = False
for i in range(s, len(html)):
    c = html[i]
    if esc: esc = False; continue
    if c == '\\' and in_str: esc = True; continue
    if c == '"': in_str = not in_str; continue
    if in_str: continue
    if c == '{': d += 1
    elif c == '}':
        d -= 1
        if d == 0:
            itr = json.loads(html[s:i+1])
            break

tc = itr['hourly_closed_eit']
ts = itr['today_submitted_assets']
e, i, t = itr['eit_summary']

print(f'ITR Today Close {tc} Submit {ts}')
print(f'E Total {e["total"]} Closed {e["closed"]} ({e["pct"]}%) Open {e["total"]-e["closed"]}')
print(f'I Total {i["total"]} Closed {i["closed"]} ({i["pct"]}%) Open {i["total"]-i["closed"]}')
print(f'T Total {t["total"]} Closed {t["closed"]} ({t["pct"]}%) Open {t["total"]-t["closed"]}')
