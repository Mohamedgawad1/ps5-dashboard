"""
Watchdog - keeps server, auto_sync, and tunnel alive.
Runs continuously. Restarts any dead process within 10 seconds.
"""
import subprocess
import time
import os
import sys
import re
import threading
from datetime import datetime

BASE = r"C:\Users\mylap\OneDrive\Desktop\dashboard"
LOG = os.path.join(BASE, "watchdog_log.txt")

PROCESSES = {
    "server": ["python", os.path.join(BASE, "mobile_app", "server.py")],
    "auto_sync": ["python", os.path.join(BASE, "auto_sync_rfi.py")],
    "cloud_sync": ["python", os.path.join(BASE, "cloud_auto_update.py")],
    "tunnel_named": [os.path.join(BASE, "cloudflared.exe"), "tunnel", "run", "cpp-eit"],
    "tunnel_backup": [os.path.join(BASE, "cloudflared.exe"), "tunnel", "--url", "http://localhost:8080"],
}

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def is_running(proc):
    return proc.poll() is None

def start_process(name, cmd):
    try:
        kwargs = {"cwd": BASE}
        if sys.platform == 'win32':
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        if name == "tunnel_backup":
            kwargs["stdout"] = subprocess.PIPE
            kwargs["stderr"] = subprocess.STDOUT
        p = subprocess.Popen(cmd, **kwargs)
        log(f"STARTED {name} (PID {p.pid})")
        if name == "tunnel_backup":
            def read_output(proc):
                try:
                    for line in iter(proc.stdout.readline, b''):
                        text = line.decode('utf-8', errors='replace').strip()
                        if 'trycloudflare.com' in text:
                            urls = re.findall(r'https://[^\s]+\.trycloudflare\.com', text)
                            if urls:
                                url = urls[0]
                                log(f"TUNNEL URL: {url}")
                                try:
                                    with open(os.path.join(BASE, "tunnel_url.txt"), "w") as f:
                                        f.write(url)
                                except Exception:
                                    pass
                                try:
                                    desktop = os.path.join(os.environ['USERPROFILE'], 'OneDrive', 'Desktop')
                                    html_path = os.path.join(desktop, 'EIT App URL.html')
                                    html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>EIT App - Current URL</title>
<meta http-equiv="refresh" content="120">
<style>
body{{font-family:Segoe UI,sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;background:linear-gradient(135deg,#0d1b4a,#1a237e);color:#fff}}
.box{{text-align:center;padding:30px 40px;background:rgba(255,255,255,.1);border-radius:16px;backdrop-filter:blur(10px)}}
h2{{margin-bottom:16px}}
a{{color:#64ffda;font-size:16px;text-decoration:none;word-break:break-all;display:block;margin:8px 0}}
a:hover{{text-decoration:underline}}
.label{{font-size:11px;opacity:.5;margin-top:4px}}
hr{{border:none;border-top:1px solid rgba(255,255,255,.15);margin:16px 0}}
</style>
</head>
<body>
<div class="box">
<h2>CPP AGI - EIT Cables</h2>
<div class="label">Remote (Quick Tunnel)</div>
<a href="{url}" target="_blank">{url}</a>
<hr>
<div class="label">Local (this PC)</div>
<a href="http://localhost:8080" target="_blank">http://localhost:8080</a>
<hr>
<div class="label">Auto-refreshes every 2 min</div>
</div>
</body>
</html>'''
                                    with open(html_path, 'w', encoding='utf-8') as f:
                                        f.write(html)
                                except Exception:
                                    pass
                except Exception:
                    pass
            threading.Thread(target=read_output, args=(p,), daemon=True).start()
        return p
    except Exception as e:
        log(f"FAILED to start {name}: {e}")
        return None

FIXED_URL = "https://app.mohamedgawwad.is-a.dev"
FIXED_URL_HTTP = "http://app.mohamedgawwad.is-a.dev"

def main():
    log("=" * 50)
    log("Watchdog started - keeping everything alive")
    log(f"Fixed URL: {FIXED_URL}")

    procs = {}
    for name, cmd in PROCESSES.items():
        procs[name] = start_process(name, cmd)
        time.sleep(1)

    check_count = 0

    while True:
        time.sleep(15)
        check_count += 1

        for name, cmd in PROCESSES.items():
            p = procs.get(name)
            if p is None or not is_running(p):
                log(f"RESTARTING {name}")
                procs[name] = start_process(name, cmd)
                time.sleep(3)

        if check_count % 240 == 0:
            log(f"Watchdog alive - {check_count} checks. URL: {FIXED_URL}")

if __name__ == "__main__":
    main()
