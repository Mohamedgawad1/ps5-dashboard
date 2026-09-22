import urllib.request, json

url = 'https://mohamedgawad1.github.io/ps5-dashboard/index.html'
resp = urllib.request.urlopen(url, timeout=20)
html = resp.read().decode()
start = html.find('const ITR = ')
s = html.index('{', start)
d=0;in_str=0;esc=0
for i in range(s, len(html)):
    c=html[i]
    if esc: esc=0; continue
    if c == '\\' and in_str: esc=1; continue
    if c == '"': in_str ^= 1; continue
    if in_str: continue
    if c == '{': d+=1
    elif c == '}':
        d-=1
        if d==0:
            itr = json.loads(html[s:i+1])
            break

print('=== GitHub Pages ===')
print('hourly_total:', itr.get('hourly_total'))
print('hourly_closed_eit:', itr.get('hourly_closed_eit'))
print('today_submitted_assets:', itr.get('today_submitted_assets'))
e, i, t = itr.get('eit_summary', [{}, {}, {}])
print(f'E: {e.get("closed")}/{e.get("total")} ({e.get("pct")}%)')
print(f'I: {i.get("closed")}/{i.get("total")} ({i.get("pct")}%)')
print(f'T: {t.get("closed")}/{t.get("total")} ({t.get("pct")}%)')
