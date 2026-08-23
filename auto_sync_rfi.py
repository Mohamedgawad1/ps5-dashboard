"""
Auto-sync: scans ALL Downloads subfolders for PDFs (especially CPP-RFI).
Checks every 30 seconds for new files.
Full rebuild every 2 hours.
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

SKIP_PATTERNS = ['mohamed', 'cv', 'attendance', 'punch', 'punchlist',
                 'oil.pdf', 'task datasheet', 'document_control',
                 '1784709', '1786128', 'resume', 'letter', 'contract',
                 'invoice', 'receipt', 'photo', 'image', 'screenshot']

REBUILD_INTERVAL = 7200  # 2 hours in seconds
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
        seen_files = {f.name for f in MASTER.rglob("*.pdf")}
    log(f"Tracking {len(seen_files)} existing PDFs in WIRING - MASTER")

def sync_all_from_downloads():
    """Recursively scan ALL Downloads subfolders for PDFs and copy to MASTER."""
    copied = 0
    if not DOWNLOADS.exists():
        return 0
    master_lower = {n.lower() for n in seen_files}
    for pdf in DOWNLOADS.rglob("*.pdf"):
        if pdf.is_dir():
            continue
        fname = pdf.name
        if fname.lower() in master_lower:
            continue
        if any(p in fname.lower() for p in SKIP_PATTERNS):
            continue
        if not fname.lower().startswith("cpp-rfi"):
            continue
        dest = MASTER / fname
        if not dest.exists():
            try:
                shutil.copy2(str(pdf), str(dest))
                log(f"COPIED [{pdf.parent.name}]: {fname}")
                seen_files.add(fname)
                master_lower.add(fname.lower())
                copied += 1
            except Exception as e:
                log(f"ERROR copying {fname}: {e}")
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
        capture_output=True, text=True, timeout=600
    )
    for line in (result.stdout + result.stderr).splitlines():
        if "DONE" in line:
            log(line.strip())
    log("Rebuild complete.")

def main():
    log("=== auto_sync started ===")
    log(f"Watching Downloads: {DOWNLOADS} (all subfolders)")
    log(f"Watching MASTER: {MASTER}")
    log(f"Rebuild every: {REBUILD_INTERVAL // 3600} hours")

    init_seen()
    last_rebuild = time.time()
    rebuild_needed = False

    while True:
        try:
            c1 = sync_all_from_downloads()
            c2 = check_master_for_new()
            total = c1 + c2
            if total > 0:
                rebuild_needed = True
                log(f"{total} new file(s) detected")

            elapsed = time.time() - last_rebuild
            if (rebuild_needed and total == 0) or elapsed >= REBUILD_INTERVAL:
                time.sleep(5)
                extra = sync_all_from_downloads() + check_master_for_new()
                if extra > 0:
                    continue
                rebuild_data()
                rebuild_needed = False
                last_rebuild = time.time()
        except Exception as e:
            log(f"ERROR: {e}")

        time.sleep(30)

if __name__ == "__main__":
    main()
