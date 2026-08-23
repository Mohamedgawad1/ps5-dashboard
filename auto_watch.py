"""
PS5 PLATFORM AUTO WATCHER
Watches project Excel files -> any change triggers:
  data.json + index.html rebuild + direct git push (GitHub Pages)
Runs silently at logon (pythonw). Log: auto_watch_log.txt
"""
import os, sys, glob, time, shutil, subprocess, traceback
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(BASE, 'auto_watch_log.txt')
COOLDOWN = 600
DEBOUNCE = 90

WATCH_FILES = [
    'ovTasks_TestsPlanned_1369.xlsx',
    'ovPunchlist_1399.xlsx',
    'PS-5 EIT PUNCH LIST REGISTER.xlsx',
    'PS-5 INSPECTION REGISTER.xlsx',
    'PS5 Master tracker EIT Combined.xlsx',
    'PS5 EIT CPP AGI Dashboard.xlsx',
]

def log(msg):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    try:
        with open(LOG, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception:
        pass

def snapshot():
    st = {}
    for f in WATCH_FILES:
        p = os.path.join(BASE, f)
        if os.path.exists(p):
            st[f] = os.path.getmtime(p)
    for p in glob.glob(os.path.join(BASE, 'PS-5 COMPLETIONS DPR SUMMERY*.xlsx')):
        st[os.path.basename(p)] = os.path.getmtime(p)
    return st

def file_ready(path):
    tmp = os.path.join(BASE, '_temp', '_watch_probe.tmp')
    try:
        os.makedirs(os.path.dirname(tmp), exist_ok=True)
        shutil.copy2(path, tmp)
        os.remove(tmp)
        return True
    except Exception:
        return False

def run_cycle():
    log("change detected -> running update cycle")
    for f in WATCH_FILES:
        p = os.path.join(BASE, f)
        if os.path.exists(p):
            for i in range(30):
                if file_ready(p):
                    break
                time.sleep(10)
                if i == 29:
                    log(f"  {f} still locked, aborting this cycle")
                    return False
    steps = [
        ['python', 'daily_update.py'],
        ['python', 'cpp_agi_dashboard.py'],
    ]
    for cmd in steps:
        r = subprocess.run(cmd, cwd=BASE, capture_output=True, text=True,
                           encoding='utf-8', errors='replace')
        tail = (r.stdout or '').strip().splitlines()
        tail = tail[-1] if tail else ''
        log(f"  {'$'+' '.join(cmd[1:])} rc={r.returncode} {tail}")
        if r.returncode != 0:
            err = (r.stderr or '').strip().splitlines()
            log('   ERR: ' + (err[-1] if err else 'unknown'))
            return False
    subprocess.run(['git', 'add', '-u'], cwd=BASE, capture_output=True)
    d = subprocess.run(['git', 'diff', '--cached', '--quiet'], cwd=BASE, capture_output=True)
    if d.returncode == 0:
        log("  no changes to push")
        return True
    c = subprocess.run(['git', '-c', 'user.name=ps5-bot', '-c', 'user.email=ps5@local',
                        'commit', '-m', f"platform update {datetime.now():%Y-%m-%d %H:%M}"],
                       cwd=BASE, capture_output=True, text=True)
    log("  committed")
    p = subprocess.run(['git', 'push', 'origin', 'main'], cwd=BASE,
                       capture_output=True, text=True)
    ok = p.returncode == 0
    log("  PUSHED OK" if ok else f"  PUSH FAILED {(p.stderr or '').strip()[-200:]}")
    return ok

def main():
    os.chdir(BASE)
    now_mode = '--now' in sys.argv
    log(f"watcher started pid={os.getpid()}{' (immediate)' if now_mode else ''}")
    last_run = 0
    if now_mode:
        run_cycle()
        last_run = time.time()
    prev = snapshot()
    while True:
        time.sleep(60)
        cur = snapshot()
        changed = [k for k in cur if prev.get(k) != cur[k]]
        if not changed:
            prev = cur
            continue
        if time.time() - last_run < COOLDOWN:
            prev = cur
            log(f"changed {len(changed)} file(s) during cooldown, skipped")
            continue
        time.sleep(DEBOUNCE)
        prev = snapshot()
        if all(prev.get(k) == cur.get(k) for k in changed):
            pass
        else:
            log("still changing, will re-check next minute")
            continue
        if run_cycle():
            prev = snapshot()
            last_run = time.time()

if __name__ == '__main__':
    try:
        main()
    except Exception:
        log('CRASH: ' + traceback.format_exc())
