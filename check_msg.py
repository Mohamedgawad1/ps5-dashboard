import urllib.request, json

INDEX_URL = "https://mohamedgawad1.github.io/ps5-dashboard/index.html"

def extract_itr(html):
    start = html.find('const ITR = ')
    if start < 0: return None
    start = html.index('{', start)
    depth = 0; in_str = False; esc = False
    for i in range(start, len(html)):
        ch = html[i]
        if esc: esc = False; continue
        if ch == '\\' and in_str: esc = True; continue
        if ch == '"': in_str = not in_str; continue
        if in_str: continue
        if ch == '{': depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                return json.loads(html[start:i+1])
    return None

resp = urllib.request.urlopen(INDEX_URL, timeout=20)
html = resp.read().decode()
itr = extract_itr(html)

tc = int(itr.get('hourly_closed_eit', 0))
ts = int(itr.get('today_submitted_assets', 0))
eit = itr.get('eit_summary', [])
disc = {}
for d in eit:
    code = d['label'][0]
    disc[code] = d

e, i, t = disc.get('E',{}), disc.get('I',{}), disc.get('T',{})

msg = (
    f"ITR Today Close {tc} Submit {ts}\n"
    f"E Total {e.get('total',0)} Closed {e.get('closed',0)} ({e.get('pct',0)}%) Open {e.get('total',0)-e.get('closed',0)}\n"
    f"I Total {i.get('total',0)} Closed {i.get('closed',0)} ({i.get('pct',0)}%) Open {i.get('total',0)-i.get('closed',0)}\n"
    f"T Total {t.get('total',0)} Closed {t.get('closed',0)} ({t.get('pct',0)}%) Open {t.get('total',0)-t.get('closed',0)}"
)
print(msg)
