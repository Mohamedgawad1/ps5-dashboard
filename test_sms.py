import urllib.request, re, json

URL = "https://mohamedgawad1.github.io/ps5-dashboard/index.html"

resp = urllib.request.urlopen(URL, timeout=20)
html = resp.read().decode()

start = html.find("const ITR = ")
start = html.index("{", start)
depth = 0; instr = False; esc = False
for i in range(start, len(html)):
    c = html[i]
    if esc: esc = False; continue
    if c == "\\" and instr: esc = True; continue
    if c == '"': instr = not instr; continue
    if instr: continue
    if c == "{": depth += 1
    elif c == "}":
        depth -= 1
        if depth == 0:
            data = json.loads(html[start:i+1])
            break

tc = int(data.get("hourly_closed_eit", 0))
ts = int(data.get("hourly_submitted", 0))

def disc(code):
    for d in data.get("eit_summary", []):
        if d.get("label", "")[:1] == code: return d
    return {"total": 0, "closed": 0}

e = disc("E"); i = disc("I"); t = disc("T")

def pct(c, tot):
    return round(c / tot * 100, 1) if tot else 0

msg = (
    f"ITR Today Close {tc} Submit {ts}\n"
    f"E Total {e['total']} Closed {e['closed']} ({pct(e['closed'],e['total'])}%) Open {e['total']-e['closed']}\n"
    f"I Total {i['total']} Closed {i['closed']} ({pct(i['closed'],i['total'])}%) Open {i['total']-i['closed']}\n"
    f"T Total {t['total']} Closed {t['closed']} ({pct(t['closed'],t['total'])}%) Open {t['total']-t['closed']}"
)

print("=== TEST SMS (not sent) ===")
print(msg)
print("===========================")
