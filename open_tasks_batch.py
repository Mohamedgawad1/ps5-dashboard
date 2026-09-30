"""Open every task-completion record for the assets in the RFI PDFs, all in one browser run.

One tab per task, nothing is typed, nothing is saved, nothing is completed.
"""

import argparse
import json
import os
import subprocess
import time
import urllib.parse
import urllib.request
from pathlib import Path

from playwright.sync_api import sync_playwright

from fill_rfi import (
    EDIT_URL,
    PROFILE_DIR,
    SWITCHBOARD,
    TASKS_GRID,
    TASK_FIELD,
    log,
    open_tasks_app,
    type_into,
    wait_products,
)
from itr_upload import find_tasks, read_pdf, task_index
import cdp_query
import sc_browser

PDF_DIR = r"C:\Users\mylap\Downloads\rfi"
SKIP_FILE = "skip_tasks.txt"
RECORD_CACHE = Path(__file__).parent / "task_record_cache.json"


def load_records():
    try:
        return json.loads(RECORD_CACHE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_records(data):
    try:
        RECORD_CACHE.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception as exc:
        log("  (cache save failed: %s)" % str(exc)[:80])


CDP = "http://127.0.0.1:9222"


def http_new_tab(url, timeout=20):
    """Open a tab through the DevTools HTTP endpoint - no CDP attach needed."""
    target = "%s/json/new?%s" % (CDP, urllib.parse.quote(url, safe=""))
    req = urllib.request.Request(target, method="PUT")
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8", "ignore"))


def cdp_tabs(timeout=15):
    with urllib.request.urlopen(CDP + "/json/list", timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8", "ignore"))


def ensure_browser(seconds=120):
    """Wait for the SmartCloud Chrome; start it if it is not running."""
    end = time.time() + seconds
    while time.time() < end:
        try:
            with urllib.request.urlopen(CDP + "/json/version", timeout=4) as r:
                name = json.loads(r.read().decode()).get("Browser", "")
            if name:
                return name
        except Exception:
            pass
        if time.time() > end - seconds + 3:
            log("المتصفح مش شغال - بيشتغله دلوقتي ...")
            bat = Path(__file__).parent / "Start_SmartCloud.bat"
            try:
                subprocess.Popen(["cmd", "/c", str(bat)], creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            except Exception as exc:
                log("  start failed: %s" % str(exc)[:90])
        time.sleep(3)
    return ""


def open_tabs_fast(plan_rows):
    """Every record id is known, so skip the browser attach and just fire the tabs."""
    log("كل الـ records معروفة - هفتح الـ tabs مباشرة من غير ما أربط بالمتصفح\n")
    opened_ids = []
    for n, (tid, row) in enumerate(plan_rows, 1):
        url = EDIT_URL % row["record"]
        try:
            http_new_tab(url)
            opened_ids.append(tid)
            log("   [tab %d/%d] %s -> record %s" % (n, len(plan_rows), row["task"], row["record"]))
        except Exception as exc:
            log("   !! %s: %s" % (row["task"], str(exc)[:90]))
        time.sleep(0.4)
    return opened_ids


def verify_cached(known, tasks, live=True, stream=None):
    """Re-check every cached task id against the live grid, then return the rows.

    The cache is only a shortcut, never the authority: if the record id changed
    we take the live one, and if the task is gone we refuse to open anything.
    """
    rows, failed2 = [], []
    if not live:
        return [(t, known[t]) for t in tasks if known.get(t)], []
    app = None
    try:
        app = cdp_query.open_app_tab()
    except Exception as exc:
        log("  (live check skipped: %s)" % str(exc)[:80])
    if app is None:
        log("  live check unavailable - falling back to the cache")
        return [(t, known[t]) for t in tasks if known.get(t)], []
    changed = 0
    for n, tid in enumerate(tasks, 1):
        hit = None
        for _ in (1, 2):
            try:
                hit = cdp_query.find_task(app, tid)
            except Exception as exc:
                log("   !! %s" % str(exc)[:80])
                try:
                    app.close()
                    app = cdp_query.open_app_tab()
                except Exception:
                    app = None
            if hit:
                break
        if not hit:
            log("   !! %s مش موجود في الـ grid دلوقتي" % tid)
            failed2.append(tid)
            continue
        cached = known.get(tid) or {}
        if cached.get("record") and cached["record"] != hit["record"]:
            log("   %s: الـ record اتغير %s -> %s (هحدّث الـ cache)" % (tid, cached["record"], hit["record"]))
            changed += 1
        known[tid] = hit
        rows.append((tid, hit))
        log("[%d/%d] %s -> %s" % (n, len(tasks), tid, hit["record"]))
        if stream:
            try:
                http_new_tab(EDIT_URL % hit["record"])
                stream.append(tid)
                log("      [tab %d/%d] اتفتح" % (len(stream), len(tasks)))
            except Exception as exc:
                log("      !! %s" % str(exc)[:60])
            time.sleep(0.3)
    if app is not None:
        try:
            app.close()
        except Exception:
            pass
    if changed:
        save_records(known)
    return rows, failed2


def wait_browser_open(count, minutes=240):
    """Stay alive with the browser, never close it."""
    log("\nفتحت %d صفحة. المتصفح سيبه مفتوح - اقفل الـ tabs لما تخلص." % count)
    end = time.time() + minutes * 60
    while time.time() < end:
        time.sleep(30)
        try:
            cdp_tabs(timeout=10)
        except Exception:
            log("  (المتصفح مش راد - مستنيك ترجّعه)")
    return


def is_closed(row):
    """A closed task cannot take the RFI file any more, so never open it."""
    return "closed" in (row.get("state") or "").lower()


def load_skip():
    """Task IDs written in skip_tasks.txt next to the PDFs, one per line."""
    out = set()
    for path in (Path(PDF_DIR) / SKIP_FILE, Path(__file__).parent / SKIP_FILE):
        try:
            for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
                line = line.split("#", 1)[0].strip().upper()
                if line:
                    out.add(line)
        except Exception:
            pass
    return out


def add_skip(task_id):
    path = Path(PDF_DIR) / SKIP_FILE
    try:
        current = path.read_text(encoding="utf-8", errors="ignore").splitlines()
    except Exception:
        current = []
    if (task_id or "").strip().upper() not in {l.strip().upper() for l in current}:
        with path.open("a", encoding="utf-8") as f:
            f.write(task_id.strip().upper() + "\n")
    return path


def ensure_session(ctx):
    for extra in list(ctx.pages[1:]):
        try:
            extra.close()
        except Exception:
            pass
    home = ctx.pages[0] if ctx.pages else ctx.new_page()
    home.bring_to_front()
    try:
        home.goto(SWITCHBOARD, wait_until="domcontentloaded", timeout=60000)
    except Exception as exc:
        log("goto note: %s" % str(exc)[:120])
    if wait_products(home, 45):
        log("session OK")
        return home
    log("الجلسة مش حية - سجّل دخولك في النافذة (مستني 45 دقيقة)")
    end = time.time() + 2700
    attempt = 0
    told = False
    while time.time() < end:
        attempt += 1
        if wait_products(home, 20):
            break
        try:
            if "401.aspx" in home.url or "Not Authorized" in (home.title() or ""):
                clicked = home.evaluate("""() => { for (const el of document.querySelectorAll('*')) {
                    const t = (el.innerText || '').trim();
                    if (t === 'OK' && el.children.length === 0) { el.click(); return true; } } return false; }""")
                log("  401: OK pressed = %s" % clicked)
                if attempt >= 2 and "401.aspx" in (home.url or ""):
                    home.goto("https://wly04-sc.intergraphsmartcloud.com/ISC/Login.aspx", wait_until="domcontentloaded", timeout=60000)
                    log("  فتحت صفحة الدخول مباشرة")
            if not told and "login" in (home.url or "").lower():
                told = True
                try:
                    fields = home.evaluate("""() => [...document.querySelectorAll('input')].map(e => (e.type||'text') + ':' + (e.id || e.name)).slice(0, 8)""")
                    log("  حقول صفحة الدخول: %s" % fields)
                except Exception:
                    pass
        except Exception as exc:
            log("  (401/login note: %s)" % str(exc)[:80])
        if attempt % 6 == 1:
            log("  مستني login ... %s" % home.url[:110])
        time.sleep(3)
    if not wait_products(home, 20):
        log("NO SESSION - مفيش دخول")
        return None
    log("session OK")
    return home


GRID_ROWS = r"""(s) => { try {
    const g = document.querySelector(s); const kg = $(g).data('kendoGrid');
    if (!kg) return null;
    const rows = kg.dataSource.view().map(r => {
        const tr = kg.tbody.find('tr[data-uid="' + r.uid + '"]');
        const a = tr && tr.length ? tr[0].querySelector('a[onclick*="vTasks_TestsCompletion"]') : null;
        const m = a ? (a.getAttribute('onclick') || '').match(/vTasks_TestsCompletion",\s*(\d+)/) : null;
        return {record: m ? m[1] : null, task: (r.TaskName || '').trim(), asset: r.AssetTag, type: r.TaskType, state: r.WeekOverdue};
    });
    return {total: kg.dataSource.total(), page: 1, pages: 1, rows: rows};
  } catch (e) { return {error: String(e)}; } }"""


GRID_BUSY = r"""s => { const g = document.querySelector(s); if (!g) return true;
    const kg = $(g).data('kendoGrid'); if (!kg) return true;
    const ds = kg.dataSource;
    return !!(ds._requestStart && ds._requestStart !== ds._requestEnd); }"""


class AppGone(Exception):
    """The tasks app tab closed - the caller must re-open it and retry."""


def _alive(page):
    try:
        return page is not None and not page.is_closed()
    except Exception:
        return False


def _grid_snapshot(app):
    """One read of the grid, no waiting."""
    if not _alive(app):
        raise AppGone("app tab closed")
    try:
        snap = app.evaluate(GRID_ROWS, TASKS_GRID)
    except Exception as exc:
        if "closed" in str(exc).lower() or "target" in str(exc).lower():
            raise AppGone(str(exc)[:120])
        log("   grid read error: %s" % str(exc)[:100])
        return None
    if snap and snap.get("error"):
        log("   grid js error: %s" % str(snap["error"])[:140])
        return None
    return snap


def _wait_grid(app, want, seconds=25):
    """Wait until the grid finished the search and (if any) holds the wanted task."""
    want = (want or "").strip().upper()
    end = time.time() + seconds
    snap = None
    while time.time() < end:
        snap = _grid_snapshot(app)
        if snap:
            for r in snap["rows"]:
                if (r["task"] or "").strip().upper() == want:
                    return snap, r
        try:
            if not app.evaluate(GRID_BUSY, TASKS_GRID) and snap is not None and snap.get("rows"):
                return snap, None
        except Exception as exc:
            if "closed" in str(exc).lower() or "target" in str(exc).lower():
                raise AppGone(str(exc)[:120])
        time.sleep(0.5)
    return snap, None


def record_id_for(app, task_id):
    """Find the row whose TaskName is exactly task_id, walking the grid pages.

    Never fall back to "the first row with a link" - that opens the wrong record.
    """
    want = (task_id or "").strip().upper()
    if not want:
        return None
    if not type_into(app, TASK_FIELD, task_id):
        return None
    app.evaluate("s => { const b = document.querySelector(s); if (b) b.click(); }", "#btn_vTasks_TestsPlanned_Search_Search")

    snap, hit = _wait_grid(app, want, seconds=25)
    if not snap:
        return None
    if hit:
        return hit

    pages = int(snap.get("pages") or 0) or 1
    if snap.get("total", 0) > 0 and pages > 1:
        for p in range(2, min(pages, 30) + 1):
            try:
                app.evaluate("(s, p) => { const kg = $(document.querySelector(s)).data('kendoGrid'); kg.pager.page(p); }", TASKS_GRID, p)
            except Exception as exc:
                log("   pager note: %s" % str(exc)[:80])
                break
            snap = _grid_snapshot(app)
            if not snap:
                break
            for r in snap["rows"]:
                if (r["task"] or "").strip().upper() == want:
                    log("   (found on page %d)" % p)
                    return r

    # last resort: a partial match, so a slightly different name still opens
    for r in (snap.get("rows") or []):
        t = (r["task"] or "").strip().upper()
        if t and want and (want in t or t in want):
            log("   %s: matched loosely to %s" % (task_id, r["task"]))
            return r

    log("   %s: مش موجود في الـ grid (%d نتيجة / %d صفحة) - هتخطاه"
        % (task_id, snap.get("total", 0), pages))
    return None


def open_record(ctx, record_id, label, expect_task=None):
    page = ctx.new_page()
    page.set_default_timeout(25000)
    try:
        page.goto(EDIT_URL % record_id, wait_until="commit", timeout=60000)
    except Exception as exc:
        log("   goto note: %s" % str(exc)[:100])
    ok = False
    for _ in range(25):
        try:
            page.wait_for_timeout(3000)
            body = page.evaluate("() => document.body ? document.body.innerText : ''")
        except Exception:
            log("   the tab closed itself")
            try:
                page.close()
            except Exception:
                pass
            return None
        if len(body or "") > 300:
            ok = True
            break
    if not ok:
        log("   !! record %s did not render" % record_id)
        try:
            page.close()
        except Exception:
            pass
        return None
    if expect_task:
        if expect_task.strip().upper() not in (body or "").upper():
            log("   !! record %s does not show %s - closing, this is the wrong record"
                % (record_id, expect_task))
            try:
                page.close()
            except Exception:
                pass
            return None
    log("   opened: %s (record %s)" % (label, record_id))
    return page


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--assets", nargs="*", default=None)
    ap.add_argument("--tasks", nargs="*", default=None)
    ap.add_argument("--pdf", nargs="*", default=None)
    ap.add_argument("--type", nargs="*", default=["Static Test"], help="fallback task types when the PDF states none")
    ap.add_argument("--allow-other-type", action="store_true", help="open another type if the wanted one is missing")
    ap.add_argument("--no-other-type", action="store_true", help="only open the exact detected type, never fall back")
    ap.add_argument("--dry-run", action="store_true", help="print the plan only, do not open any tab")
    ap.add_argument("--skip-missing", action="store_true", help="write tasks missing on the platform to skip_tasks.txt")
    ap.add_argument("--close-old", action="store_true", help="close the task tabs of the previous run (default: leave them for you)")
    ap.add_argument("--no-fast", action="store_true", help="always look the record ids up on the platform")
    ap.add_argument("--no-verify", action="store_true", help="trust the cached record ids without re-checking the task id live")
    ap.add_argument("--no-cdp", action="store_true", help="skip the raw cdp path, use Playwright")
    ap.add_argument("--keep-open", type=int, default=0, help="0 = wait until you close the browser")
    args = ap.parse_args()

    tasks = list(args.tasks or [])
    if not tasks:
        # one (asset, type) pair per PDF - never share a single type across PDFs
        pairs = []
        seen_pairs = set()

        def add_pair(tag, want, origin):
            key = (tag, (want or "").lower())
            if tag and want and key not in seen_pairs:
                seen_pairs.add(key)
                pairs.append((tag, want, origin))

        for tag in args.assets or []:
            for want in args.type:
                add_pair(tag, want, "--assets")

        pdfs = args.pdf or sorted(str(p) for p in Path(PDF_DIR).glob("*.pdf"))
        for pdf in pdfs:
            try:
                info = read_pdf(pdf)
            except Exception as exc:
                log("  could not read %s: %s" % (Path(pdf).name, str(exc)[:110]))
                continue
            own = info.get("detected_type") or info.get("mapped_type")
            log("%s" % Path(pdf).name)
            log("   RFI No: %s | النوع: %s" % (info.get("rfi_no"), own))
            log("   الأصول (%d): %s" % (len(info.get("asset_tags") or []), ", ".join(info.get("asset_tags") or [])))
            if not own:
                log("   ! مفيش نوع واضح في الـ PDF - هستخدم %s" % "/".join(args.type))
                for want in args.type:
                    for tag in info.get("asset_tags") or []:
                        add_pair(tag, want, Path(pdf).name)
            else:
                for tag in info.get("asset_tags") or []:
                    add_pair(tag, own, Path(pdf).name)

        log("\n%d (asset, type) pair(s)\n" % len(pairs))
        skip = load_skip()
        if skip:
            log("skip list (%d): %s" % (len(skip), ", ".join(sorted(skip))))
        log_types = set()
        by_origin = {}
        for tag, want, origin in pairs:
            by_origin.setdefault(origin, []).append((tag, want))
        for origin, items in by_origin.items():
            picked, fallback, closed = [], [], []
            for tag, want in items:
                rows = find_tasks(tag) or []
                hit = [r for r in rows if (r.get("type") or "").lower() == want.lower()]
                open_hit = [r for r in hit if not is_closed(r)]
                if open_hit:
                    picked.extend((tag, r) for r in open_hit)
                    closed.extend((tag, r) for r in hit if is_closed(r))
                    continue
                if hit:  # the wanted type exists but every one of them is closed
                    closed.extend((tag, r) for r in hit)
                others = ", ".join("%s/%s" % (r.get("task_id"), r.get("type")) for r in rows) or "مفيش"
                log("  - %s [%s]: مفيش %s مفتوح | الموجود: %s" % (tag, want, want, others))
                fallback.extend((tag, r) for r in rows if not is_closed(r))
            fallback = [(t, r) for t, r in fallback if (t, r) not in picked]
            if not picked and fallback and not args.no_other_type:
                kinds = sorted({(r.get("type") or "?") for _, r in fallback})
                log("  ! الـ RFI طلب %s، لكن الأصول على المنصة: %s"
                    % ("/".join(sorted({w for _, w in items})), ", ".join(kinds)))
                log("  >> النوع اللي هيتفتح فعليًا: %s" % ", ".join(kinds))
                log_types.update(kinds)
                picked = fallback
            else:
                log_types.update((r.get("type") or "?") for _, r in picked)
            if closed:
                log("  x %d task مقفول (Closed) - مش هيتفتح" % len(closed))
            for tag, row in picked:
                tid = row.get("task_id")
                if tid and tid in skip:
                    log("  = %s | %s | %s | %s  (في قائمة التخطي)"
                        % (tid, row.get("type"), row.get("state"), tag))
                    continue
                if tid and tid not in tasks:
                    tasks.append(tid)
                    log("  + %s | %s | %s | %s" % (tid, row.get("type"), row.get("state"), tag))

    if not tasks:
        log("!! no tasks to open")
        return 2
    log("\n%d task(s) to open" % len(tasks))
    if log_types:
        log("النوع اللي هيتفتح: %s" % ", ".join(sorted(log_types)))
    if args.dry_run:
        log("DRY RUN - مفيش تاب اتفتح. راجع القائمة فوق قبل التشغيل.")
        return 0

    opened, failed = [], []
    if not ensure_browser():
        log("!! المتصفح لسه مش راد على 9222 - اقفل أي Chrome تاني و شغّل SmartCloud - Open Tasks")
        return 1

    # every record id is in the cache, but we still re-check the task id live so a
    # stale record can never open the wrong page (the check costs ~0s per task)
    known = load_records()
    if not args.no_fast and all(known.get(t) and known[t].get("record") for t in tasks):
        streamed = []
        rows, failed2 = verify_cached(known, tasks, live=not args.no_verify, stream=streamed)
        if rows:
            if args.close_old:
                for info in cdp_tabs():
                    u = info.get("url") or ""
                    if "vTasks_TestsCompletion" in u:
                        try:
                            urllib.request.urlopen(urllib.request.Request(CDP + "/json/close/" + info["id"], method="GET"), timeout=8)
                        except Exception:
                            pass
            done = streamed
            if len(done) < len(rows):
                done = open_tabs_fast([r for r in rows if r[0] not in done]) + done
            log("\n==================================================")
            log("فتحت %d/%d صفحة | مفيش حاجة اتكتبت ومفيش حاجة اتحفظت" % (len(done), len(rows)))
            if failed2:
                log("اللي فشلوا: %s" % ", ".join(failed2))
            log("في كل tab: تاب Files -> دوس أيقونة الرفع في عمود Files")
            for p in sorted(Path(PDF_DIR).glob("*.pdf")):
                log("الملف: %s" % p)
            log("==================================================")
            wait_browser_open(len(done), minutes=args.keep_open or 240)
            return 0

    # raw CDP path: same as before but without Playwright, so it works while the
    # background sync holds its own connection to the browser
    if not args.no_fast and not args.no_cdp:
        rows, failed2, done = [], [], []
        try:
            app = cdp_query.open_app_tab()
        except Exception as exc:
            log("  (raw cdp app failed: %s)" % str(exc)[:90])
            app = None
        if app is not None:
            fresh = 0
            for n, tid in enumerate(tasks, 1):
                log("[%d/%d] %s" % (n, len(tasks), tid), )
                hit = None
                for attempt in (1, 2):
                    try:
                        hit = cdp_query.find_task(app, tid)
                    except Exception as exc:
                        log("   !! %s" % str(exc)[:90])
                        try:
                            app.close()
                            app = cdp_query.open_app_tab()
                        except Exception:
                            app = None
                    if hit:
                        break
                if not hit:
                    log("   !! %s مش موجود في الـ grid" % tid)
                    failed2.append(tid)
                    continue
                rows.append((tid, hit))
                known[tid] = hit
                fresh += 1
                save_records(known)
                # open it right away so the pages start showing while we keep going
                try:
                    http_new_tab(EDIT_URL % hit["record"])
                    done.append(tid)
                    log("   %s -> record %s  [tab %d/%d]" % (hit["task"], hit["record"], len(done), len(tasks)))
                except Exception as exc:
                    log("   %s -> record %s  !! %s" % (hit["task"], hit["record"], str(exc)[:60]))
                time.sleep(0.3)
            app.close()
            if fresh:
                log("(حفظت %d record id في الـ cache)" % fresh)
            if rows:
                log("\n==================================================")
                log("فتحت %d/%d صفحة | مفيش حاجة اتكتبت ومفيش حاجة اتحفظت" % (len(done), len(rows)))
                if failed2:
                    log("اللي فشلوا: %s" % ", ".join(failed2))
                log("في كل tab: تاب Files -> دوس أيقونة الرفع في عمود Files")
                for p in sorted(Path(PDF_DIR).glob("*.pdf")):
                    log("الملف: %s" % p)
                log("==================================================")
                wait_browser_open(len(done), minutes=args.keep_open or 240)
                return 0

    pw = sync_playwright().start()
    browser, ctx = sc_browser.connect(pw)
    try:
        home = sc_browser.session_page(ctx)
        if not sc_browser.products_ready(home, 30):
            log("الجلسة مش حية - سجّل دخولك في نافذة Chrome (مستني 45 دقيقة)")
            end = time.time() + 2700
            attempt = 0
            told = False
            while time.time() < end:
                attempt += 1
                if sc_browser.products_ready(home, 20):
                    break
                try:
                    if "401.aspx" in (home.url or "") or "Not Authorized" in (home.title() or ""):
                        home.evaluate("""() => { for (const el of document.querySelectorAll('*')) {
                            const t = (el.innerText || '').trim();
                            if (t === 'OK' && el.children.length === 0) { el.click(); return true; } } return false; }""")
                    if not told and "login" in (home.url or "").lower():
                        told = True
                        log("  اكتب اسم المستخدم والباسورد في النافذة")
                except Exception:
                    pass
                if attempt % 6 == 1:
                    log("  مستني login ... %s" % home.url[:110])
                time.sleep(3)
        if not sc_browser.products_ready(home, 20):
            log("NO SESSION - مفيش دخول")
            return 1
        log("session OK")

        app_box = {"page": None}

        def get_app():
            p = app_box["page"]
            if _alive(p):
                return p
            try:
                if p is not None:
                    p.close()
            except Exception:
                pass
            app_box["page"] = open_tasks_app(ctx)
            log("  (re-opened the tasks app)")
            return app_box["page"]

        # ---- phase 1: resolve every record id (cache first, then the grid)
        known = load_records()
        plan_rows = []
        need_search = []
        for tid in tasks:
            hit = known.get(tid)
            if hit and hit.get("record"):
                plan_rows.append((tid, hit))
                log("  cached: %s -> record %s" % (hit.get("task", tid), hit["record"]))
            else:
                need_search.append(tid)

        fresh = {}
        for n, tid in enumerate(need_search, 1):
            log("[%d/%d] بجيب رقم الريكورد بتاع %s ..." % (n, len(need_search), tid))
            row = None
            for attempt in (1, 2, 3, 4):
                try:
                    row = record_id_for(get_app(), tid)
                    break
                except AppGone as exc:
                    log("   ~ تب التطبيق قفل - هفتحه تاني وأعيد")
                    app_box["page"] = None
                    time.sleep(2)
                except Exception as exc:
                    log("   !! attempt %d failed: %s" % (attempt, str(exc)[:110]))
                    app_box["page"] = None
                    time.sleep(3)
            if not row:
                log("   !! %s مش موجود في الـ grid - هتخطاه" % tid)
                failed.append(tid)
                if args.skip_missing:
                    log("      (أُضيف لقائمة التخطي: %s)" % add_skip(tid))
                continue
            plan_rows.append((tid, row))
            fresh[tid] = row
            log("   %s -> record %s" % (row["task"], row["record"]))

        if fresh:
            known.update(fresh)
            save_records(known)
            log("(حفظت %d record id في الـ cache)" % len(fresh))

        # ---- phase 2: fire all tabs, no waiting for the render
        if args.close_old:
            old = 0
            for p in list(ctx.pages):
                try:
                    if p.url and "vTasks_TestsCompletion" in p.url:
                        p.close()
                        old += 1
                except Exception:
                    pass
            if old:
                log("قفلت %d صفحة مهمة قديمة\n" % old)

        log("بفتح %d صفحة دلوقتي ...\n" % len(plan_rows))
        tabs = []
        for tid, row in plan_rows:
            page = ctx.new_page()
            page.set_default_timeout(25000)
            url = EDIT_URL % row["record"]
            # short timeout on purpose: fire the tab and move on, the page keeps
            # loading on its own instead of blocking the next one
            try:
                page.goto(url, wait_until="commit", timeout=8000)
            except Exception:
                pass
            tabs.append((tid, row, page))
            log("   [tab %d/%d] %s -> record %s" % (len(tabs), len(plan_rows), row["task"], row["record"]))

        # ---- phase 3: one single pass, the tabs already load on their own
        log("\nالصفحات بتحمّل دلوقتي - استنى 30 ثانية وأتأكد مرة واحدة بس...")
        time.sleep(30)
        bad = 0
        for tid, row, page in tabs:
            body = ""
            try:
                body = page.evaluate("() => document.body ? document.body.innerText : ''")
            except Exception:
                body = ""
            if len(body or "") <= 300:
                log("   ~ record %s لسه بيحمّل" % row["record"])
                bad += 1
                continue
            if (row["task"] or tid).strip().upper() not in body.upper():
                log("   ~ record %s مظهرش فيه %s - راجعه بنفسك" % (row["record"], row["task"]))
                bad += 1
                continue
            opened.append(page)

        log("\n==================================================")
        log("فتحت %d صفحة، %d محتاجة مراجعة، %d فشل" % (len(opened), bad, len(failed)))
        if failed:
            log("اللي فشلوا: %s" % ", ".join(failed))
        log("مفيش حاجة اتكتبت، ومفيش حاجة اتحفظت.")
        log("في كل tab: تاب Files -> دوس أيقونة الرفع في عمود Files")
        pdfs = sorted(Path(PDF_DIR).glob("*.pdf"))
        for p in pdfs:
            log("الملف: %s" % p)
        log("==================================================")
        sc_browser.keep_open(ctx)
    finally:
        try:
            pw.stop()
        except Exception:
            pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
