"""Auto-sync RFI from Downloads + watch WIRING-MASTER for any new PDF.
Checks every 30 seconds, rebuilds data.json when new files found.
"""
import os
import shutil
import time
import subprocess
from pathlib import Path
from datetime import datetime

DOWNLOADS = Path.home() / "Downloads"
MASTER = Path(r"C:\Users\mylap\OneDrive\Desktop\dashboard\WIRING - MASTER")
LOG = Path(r"C:\Users\mylap\OneDrive\Desktop\dashboard\auto_sync_log.txt")
SERVER_DIR = Path(r"C:\Users\mylap\OneDrive\Desktop\dashboard")

seen_files = set()

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def init_seen():
    global seen_files
    if MASTER.exists():
        seen_files = {f.name for f in MASTER.glob("*.pdf")}
    log(f"Tracking {len(seen_files)} existing PDFs in WIRING - MASTER")

def sync_rfi_from_downloads():
    copied = 0
    if not DOWNLOADS.exists():
        return 0
    for pdf in DOWNLOADS.glob("CPP-RFI*.pdf"):
        if pdf.is_dir():
            continue
        dest = MASTER / pdf.name
        if not dest.exists():
            try:
                shutil.copy2(str(pdf), str(dest))
                log(f"COPIED from Downloads: {pdf.name}")
                seen_files.add(pdf.name)
                copied += 1
            except Exception as e:
                log(f"ERROR copying {pdf.name}: {e}")
    return copied

def check_master_for_new():
    new = 0
    if not MASTER.exists():
        return 0
    for pdf in MASTER.glob("*.pdf"):
        if pdf.name not in seen_files:
            seen_files.add(pdf.name)
            log(f"NEW in MASTER: {pdf.name}")
            new += 1
    return new

def rebuild_data():
    log("Rebuilding data.json ...")
    result = subprocess.run(
        ["python", "daily_update.py"],
        cwd=str(SERVER_DIR),
        capture_output=True, text=True, timeout=300
    )
    for line in (result.stdout + result.stderr).splitlines():
        if "DONE" in line:
            log(line.strip())
    log("Done.")

def main():
    log("=== auto_sync started ===")
    log(f"Watching Downloads: {DOWNLOADS}")
    log(f"Watching MASTER: {MASTER}")

    init_seen()
    rebuild_needed = False

    while True:
        try:
            c1 = sync_rfi_from_downloads()
            c2 = check_master_for_new()
            total = c1 + c2
            if total > 0:
                rebuild_needed = True
                log(f"{total} new file(s) detected")

            if rebuild_needed and total == 0:
                time.sleep(5)
                extra = sync_rfi_from_downloads() + check_master_for_new()
                if extra > 0:
                    continue
                rebuild_data()
                rebuild_needed = False
        except Exception as e:
            log(f"ERROR: {e}")

        time.sleep(30)

if __name__ == "__main__":
    main()
