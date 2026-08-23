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
        p = subprocess.Popen(cmd, **kwargs)
        log(f"STARTED {name} (PID {p.pid})")
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
