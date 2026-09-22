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

print(f'hourly_closed_eit: {itr["hourly_closed_eit"]}')
print(f'hourly_total: {itr["hourly_total"]}')
print(f'hourly_submitted: {itr["hourly_submitted"]}')
print(f'hourly_submitted_eit: {itr["hourly_submitted_eit"]}')
print(f'today_submitted_assets: {itr["today_submitted_assets"]}')
print(f'today_closed_assets: {itr["today_closed_assets"]}')
print(f'daily today: {[d for d in itr["daily"] if d["label"] == itr["today_label"]]}')
