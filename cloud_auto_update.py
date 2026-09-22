# -*- coding: utf-8 -*-
"""
PS5 - cloud_auto_update.py
Keeps the DOWNLOADED files + dashboard + deployed pages in sync with the
ONLINE platform writes (cloud platform_state.json) and the local explorer
edits (_explorer_state.json), automatically, every 5 minutes.

Chain (only when cloud/local state changed, min 5 min between builds):
  1) sync_cloud_to_excel.py        (cloud -> DPR SUMMERY, incl added rows)
  2) apply_platform_edits_to_summery.py (local explorer -> DPR SUMMERY)
  3) punch_itr_explorer.py         (rebuilds SUBSYSTEM EXPLORER.xlsx/html)
  4) dpr_dashboard.py              (rebuilds PS5 DPR DASHBOARD.xlsm)
  5) sync_all.py                   (rebuilds platform index.html +
                                     numbered Excel downloads + git push)
  6) mirror repo PAGES/EXCEL -> Downloads ..._files folder
Run continuously (registered in watchdog.py PROCESSES).
"""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request

DASH = r"C:\Users\mylap\OneDrive\Desktop\dashboard"
PLAT = r"C:\Users\mylap\OneDrive\Desktop\PS5-COMPLETION-PLATFORM"
DL_SUB = os.path.join(os.path.expanduser('~'), 'Downloads',
                      'PS5 - CPP AGI Completion Progress Dashboard_files')
PY = sys.executable
INTERVAL = 300  # seconds (5 minutes)
LOCK = os.path.join(DASH, '_cloud_sync.lock')
LOG = os.path.join(DASH, '_cloud_sync_log.txt')

STATE_URL = ('https://raw.githubusercontent.com/Mohamedgawad1/'
             'PS5-COMPLETION-PLATFORM/main/platform_state.json')
LOCAL_STATE = os.path.join(DL_SUB, '_explorer_state.json')


def log(msg):
    line = '%s %s' % (time.strftime('%Y-%m-%d %H:%M:%S'), msg)
    print(line, flush=True)
    try:
        with open(LOG, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception:
        pass


def cloud_hash():
    try:
        req = urllib.request.Request(STATE_URL + '?t=' + str(int(time.time() * 1000)),
                                     headers={'User-Agent': 'ps5-cloud-watch'})
        with urllib.request.urlopen(req, timeout=30) as r:
            return hashlib.sha256(r.read()).hexdigest()
    except Exception as e:
        log('[warn] cloud fetch failed: %s' % e)
        return None


def local_hash():
    try:
        with open(LOCAL_STATE, 'rb') as f:
            return hashlib.sha256(f.read()).hexdigest()
    except Exception:
        return None


def locked():
    if not os.path.exists(LOCK):
        return False
    try:
        age = time.time() - os.path.getmtime(LOCK)
    except Exception:
        return True
    if age > 2700:  # stale after 45 min
        try:
            os.remove(LOCK)
        except Exception:
            pass
        return False
    return True


def run_step(name, cmd, cwd):
    log('>> ' + name)
    try:
        r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True,
                           timeout=1800)
    except Exception as e:
        log('!! %s crashed: %s' % (name, e))
        return 1
    out = (r.stdout or '').strip().splitlines()
    for ln in out[-6:]:
        log('   ' + ln)
    if r.returncode != 0:
        err = (r.stderr or '').strip().splitlines()
        for ln in err[-4:]:
            log('   !! ' + ln)
        log('!! %s FAILED (rc=%s)' % (name, r.returncode))
    return r.returncode


def run_build():
    if locked():
        log('[skip] another build in progress')
        return
    open(LOCK, 'w').close()
    t0 = time.time()
    try:
        run_step('cloud->DPR',
                 [PY, os.path.join(PLAT, 'sync_cloud_to_excel.py'), '--force'],
                 PLAT)
        run_step('local->DPR',
                 [PY, os.path.join(DASH, 'apply_platform_edits_to_summery.py')],
                 DASH)
        run_step('explorer rebuild',
                 [PY, os.path.join(DASH, 'punch_itr_explorer.py')], DASH)
        run_step('dashboard xlsm',
                 [PY, os.path.join(DASH, 'dpr_dashboard.py')], DASH)
        rc = run_step('platform full sync',
                      [PY, os.path.join(PLAT, 'sync_all.py')], PLAT)
        if rc == 0:
            mirror_downloads()
    finally:
        try:
            os.remove(LOCK)
        except Exception:
            pass
    log('[done] build in %.0fs' % (time.time() - t0))


def mirror_downloads():
    """Copy the freshly built numbered downloads into the local Downloads
    folder (PAGES/*.xlsx + EXCEL/7 - COMPLETE.xlsm)."""
    try:
        pages = os.path.join(PLAT, 'PAGES')
        excel = os.path.join(PLAT, 'EXCEL')
        if os.path.isdir(pages):
            for f in os.listdir(pages):
                if f.lower().endswith('.xlsx'):
                    shutil.copy2(os.path.join(pages, f),
                                 os.path.join(DL_SUB, f))
            log('[mirror] PAGES -> _files')
        for name in ('7 - COMPLETE.xlsm',):
            p = os.path.join(excel, name)
            if os.path.isfile(p):
                shutil.copy2(p, os.path.join(DL_SUB, name))
        # fresh explorer + dashboard outputs land here already
    except Exception as e:
        log('[warn] mirror failed: %s' % e)


def main():
    if '--once' in sys.argv:
        log('cloud_auto_update --once (manual run)')
        run_build()
        return
    log('=' * 50)
    log('cloud_auto_update started (interval %ss)' % INTERVAL)
    seen = {'cloud': None, 'local': None, 'pending': 0}
    last_run = 0
    first = True
    while True:
        ch = cloud_hash()
        lh = local_hash()
        changed = (ch is not None and (seen['cloud'] is None or ch != seen['cloud'])) \
            or (lh is not None and (seen['local'] is None or lh != seen['local']))
        if seen['cloud'] is None and seen['local'] is None and ch is None:
            time.sleep(INTERVAL)
            continue
        seen['cloud'] = ch if ch is not None else seen['cloud']
        seen['local'] = lh if lh is not None else seen['local']
        if changed:
            seen['pending'] += 1
        now = time.time()
        if (first or changed) and seen['pending'] >= 1 and (now - last_run) >= INTERVAL - 60:
            log('[trigger] state changed -> rebuilding downloads+pages')
            run_build()
            last_run = time.time()
            seen['pending'] = 0
            first = False
        time.sleep(INTERVAL - 60 if first else INTERVAL)


if __name__ == '__main__':
    main()