import json, re, subprocess, time, datetime, sys, os, urllib.request

RECIPIENTS = [
    "+255684064292",
    "+255677114429",
    "+255674643758",
    "+255699065843",
    "+255699016146",
]

GITHUB_URL = "https://mohamedgawad1.github.io/ps5-dashboard/sms_data.json"
GITHUB_HTML = "https://mohamedgawad1.github.io/ps5-dashboard/index.html"

def send_sms(number, text):
    safe = text.replace('"', "'")
    r = subprocess.run(["termux-sms-send", "-n", number, safe],
        capture_output=True, text=True, timeout=30)
    out = (r.stdout + r.stderr).strip()
    print(f"  {number}: {out[:80] if out else 'sent'}")

def fetch_data():
    local = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sms_data.json")
    if os.path.exists(local):
        print("  Using local sms_data.json")
        with open(local, 'r') as f:
            return json.load(f)

    print("  Fetching from GitHub (sms_data.json)...")
    try:
        resp = urllib.request.urlopen(GITHUB_URL, timeout=20)
        return json.loads(resp.read().decode())
    except:
        pass

    print("  Fetching from GitHub (index.html)...")
    try:
        resp = urllib.request.urlopen(GITHUB_HTML, timeout=20)
        html = resp.read().decode()
        start = html.find('const ITR = ')
        if start < 0: return None
        start = html.index('{', start)
        depth = 0; instr = False; esc = False
        for i in range(start, len(html)):
            c = html[i]
            if esc: esc = False; continue
            if c == '\\' and instr: esc = True; continue
            if c == '"': instr = not instr; continue
            if instr: continue
            if c == '{': depth += 1
            elif c == '}':
                depth -= 1
                if depth == 0:
                    try: return json.loads(html[start:i+1])
                    except: return None
    except:
        pass
    return None

def main():
    now = datetime.datetime.now()
    today = now.strftime('%Y-%m-%d')

    data = fetch_data()
    if not data:
        print("ERROR: No data found!")
        return

    data_date = data.get('date', data.get('today_label', ''))
    print(f"  Data date: {data_date}")

    tc = int(data.get('today_closed', data.get('hourly_closed_eit', 0)))
    ts = int(data.get('today_submitted', 0))

    def get_disc(code):
        for d in data.get('eit_summary', []):
            if d.get('label', '')[:1] == code:
                return d
        return {'total': 0, 'closed': 0}

    e = get_disc('E')
    i = get_disc('I')
    t = get_disc('T')

    def pct(c, tot):
        return round(c / tot * 100, 1) if tot else 0

    msg = (
        f"ITR Today Close {tc} Submit {ts}\n"
        f"E Total {e['total']} Closed {e['closed']} ({pct(e['closed'],e['total'])}%) Open {e['total']-e['closed']}\n"
        f"I Total {i['total']} Closed {i['closed']} ({pct(i['closed'],i['total'])}%) Open {i['total']-i['closed']}\n"
        f"T Total {t['total']} Closed {t['closed']} ({pct(t['closed'],t['total'])}%) Open {t['total']-t['closed']}"
    )

    print(f"\n[{now.strftime('%H:%M')}] Sending:")
    print(msg)
    print()
    for r in RECIPIENTS:
        send_sms(r, msg)
        time.sleep(2)
    print("\nDone.")

if __name__ == '__main__':
    main()
