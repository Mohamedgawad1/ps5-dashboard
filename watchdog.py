"""
Watchdog - keeps server, auto_sync, and tunnel alive.
Runs continuously. Restarts any dead process within 10 seconds.
"""
import subprocess
import time
import os
import sys
from datetime import datetime

BASE = r"C:\Users\mylap\OneDrive\Desktop\dashboard"
LOG = os.path.join(BASE, "watchdog_log.txt")

PROCESSES = {
    "server": ["python", os.path.join(BASE, "mobile_app", "server.py")],
    "auto_sync": ["python", os.path.join(BASE, "auto_sync_rfi.py")],
    "tunnel_quick": [os.path.join(BASE, "cloudflared.exe"), "tunnel", "--url", "http://localhost:8080"],
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
        if name == "tunnel_quick":
            logfile = open(os.path.join(BASE, "tunnel_quick.log"), "w")
            kwargs["stdout"] = logfile
            kwargs["stderr"] = subprocess.STDOUT
        p = subprocess.Popen(cmd, **kwargs)
        log(f"STARTED {name} (PID {p.pid})")
        return p
    except Exception as e:
        log(f"FAILED to start {name}: {e}")
        return None

def save_tunnel_url():
    logfile = os.path.join(BASE, "tunnel_quick.log")
    url_file = os.path.join(BASE, "tunnel_url.txt")
    if not os.path.exists(logfile):
        return
    try:
        import re
        with open(logfile, "r", errors="ignore") as f:
            for line in f:
                m = re.search(r'(https://[a-z0-9-]+\.trycloudflare\.com)', line)
                if m:
                    url = m.group(1)
                    with open(url_file, "w") as out:
                        out.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Current URL: {url}\n")
                    log(f"TUNNEL URL: {url}")
                    return
    except:
        pass

def main():
    log("=" * 50)
    log("Watchdog started - keeping everything alive")

    procs = {}
    for name, cmd in PROCESSES.items():
        procs[name] = start_process(name, cmd)
        time.sleep(1)

    save_tunnel_url()
    check_count = 0

    while True:
        time.sleep(15)
        check_count += 1

        for name, cmd in PROCESSES.items():
            p = procs.get(name)
            if p is None or not is_running(p):
                log(f"RESTARTING {name}")
                procs[name] = start_process(name, cmd)
                time.sleep(2)
                if name == "tunnel_quick":
                    time.sleep(15)
                    save_tunnel_url()

        if check_count % 10 == 0:
            save_tunnel_url()

if __name__ == "__main__":
    main()
