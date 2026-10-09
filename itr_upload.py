r"""
Attach a PDF to a SmartCloud Test Form record (vAssets_TestForms).

Workflow
    1. read the PDF:  the Asset Tag comes from page 1
    2. detect the PDF type from the pages  -> "Static Test" / "Conformity Check" / ...
    3. that type maps to the platform Task Type, which gives the Task ID (T-xxxxx-xxxx)
    4. open vAssets_TestForms, search by Asset Tag + Task ID, select the row
    5. open "Executed Forms For Selected Assets" and attach the PDF
    6. STOP - nothing is saved until you press Save yourself in the browser

The browser is a normal visible Chrome window with a saved profile, so the session
looks exactly like your own: you log in by hand, once.

Usage
    python itr_upload.py scan
    python itr_upload.py attach "C:\Users\mylap\Downloads\rfi\file.pdf"
    python itr_upload.py attach "C:\path\file.pdf" --task T-00069-4968
    python itr_upload.py inspect
"""

import argparse
import json
import os
import re
import sys
import time

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(os.environ.get("LOCALAPPDATA", BASE_DIR), "sc_itr_profile")
SHOTS_DIR = os.path.join(BASE_DIR, "_itr_shots")
WATCH_DIR = r"C:\Users\mylap\Downloads\rfi"
TASK_CACHE = os.path.join(BASE_DIR, "sc_pull", "data", "cpp_agi_tasks_full.json")
LIVE_TASK_INDEX = os.path.join(BASE_DIR, "sc_pull", "data", "tasks_by_tag_live.json")
# The platform only has these three task types, so map the work description onto them.
# Order matters: the first rule that matches the description wins.
TYPE_MAP_FILE = os.path.join(BASE_DIR, "rfi_type_map.json")

SWITCHBOARD_URL = (
    "https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm"
)
TESTFORMS_URL = "https://wly04-sc.intergraphsmartcloud.com/ISC/tools/vAssets_TestForms/index.htm"
VIEW = "vAssets_TestForms"

BOX_ASSET = "#textbox_Name_search_%s" % VIEW
BOX_TASK = "#textbox_TaskName_search_%s" % VIEW
BTN_SEARCH = "#btn_%s_Search_Search" % VIEW
BTN_EXECUTED = "#%s_ActionButtonbtnScannedPDFs" % VIEW

ASSET_TAG_RE = re.compile(r"\b(?:PS5|PS4|PR1|PR2)-[A-Z0-9]+(?:-[A-Z0-9]+){2,}\b")
# the same tags typed with spaces instead of dashes: "PS5 60 BI 0001G CC01"
ASSET_TAG_SPACED_RE = re.compile(
    r"\b(PS5|PS4|PR1|PR2)[\s-]+(\d{1,3})[\s-]+([A-Z]{2,4})[\s-]+(\d{3,4}[A-Z]?)[\s-]+([A-Z]{2}\d{2,3})\b"
)

TYPE_MAP_FILE = os.path.join(BASE_DIR, "rfi_type_map.json")
TYPE_RULES_FILE = os.path.join(BASE_DIR, "rfi_type_rules.json")


def _load_work_type_rules():
    """The work-description -> task-type table lives in rfi_type_rules.json so it can
    grow without touching the code. Falls back to the built-in list."""
    builtin = [
        ("Piping Test Preparations", ["piping test", "pressure test", "hydro test", "hydrotest",
                                      "pressure testing", "test preparation"]),
        ("Static Test", ["request to witness", "witness the", "witness testing", "static test",
                         "static testing", "functional test", "electrical cable testing",
                         "cable testing after installation", "cable testing",
                         "testing after installation", "cable test"]),
        ("Conformity Check", ["inspection", "glanding", "gladding", "termination", "installation",
                              "install", "device", "cable"]),
    ]
    try:
        with open(TYPE_RULES_FILE, encoding="utf-8") as f:
            data = json.load(f)
        rules = [(r["type"], [str(p).lower() for p in r.get("phrases", [])])
                 for r in data.get("rules", []) if r.get("type")]
        return rules or builtin
    except Exception:
        return builtin


TYPE_KEYWORDS = [
    # The RFI forms only say "Request for Inspection", so the type comes from the
    # "Inspection activity description" line:
    #   "Visual Inspection of ..."          -> Conformity Check
    #   "Request to Witness ... Testing"    -> Static Test
    ("Conformity Check", ["visual inspection", "inspection of", "dimensional inspection"]),
    ("Static Test", ["request to witness", "witness the", "witness testing", "static test", "static testing"]),
    ("Conformity Check", ["conformity check", "conformity-check", "conformitycheck"]),
    ("Piping Test Preparations", ["piping test", "pressure test", "hydrotest", "hydro test"]),
    ("Preservation", ["preservation", "preserve"]),
    ("Certificates", ["certificate"]),
]


ACCOUNT_RULES = """
Account safety (read this)
------------------------
* The tool NEVER types your password. Login is always done by you, by hand.
* It never kills the browser. Killing Chrome mid-session invalidates the SSO
  session, which is what causes the "Authorization Failed / session has expired"
  loop and the repeated reauthorizations that can get an account restricted.
* One login wait per run, no retry loops against the login page.
* If the session expired, the run stops and waits for you. Nothing is submitted.
"""


def profile_locked():
    """True if another Chrome is already using our profile (do NOT kill it)."""
    import subprocess

    ps = (
        "Get-CimInstance Win32_Process -Filter \"Name='chrome.exe'\" | "
        "Where-Object { $_.CommandLine -like '*sc_itr_profile*' } | "
        "Select-Object -ExpandProperty ProcessId"
    )
    try:
        out = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps],
            capture_output=True,
            text=True,
            timeout=20,
        )
        return [x.strip() for x in (out.stdout or "").splitlines() if x.strip()]
    except Exception:  # noqa: BLE001
        return []


def log(msg):
    print(msg, flush=True)


def shot(page, name):
    os.makedirs(SHOTS_DIR, exist_ok=True)
    path = os.path.join(SHOTS_DIR, "%s_%s.png" % (name, int(time.time())))
    try:
        page.screenshot(path=path, full_page=True)
        log("  [screenshot] %s" % path)
    except Exception as exc:  # noqa: BLE001
        log("  [screenshot failed] %s" % str(exc)[:120])
    return path


def interactive():
    try:
        return bool(sys.stdin) and sys.stdin.isatty()
    except Exception:  # noqa: BLE001
        return False


def pause(msg, fallback_wait=180):
    if interactive():
        try:
            input(msg)
            return
        except (EOFError, KeyboardInterrupt):
            pass
    log(msg)
    log("  (مش console تفاعلي - مستني %d ثانية)" % fallback_wait)
    time.sleep(fallback_wait)


def page_alive(page):
    try:
        return not page.is_closed()
    except Exception:  # noqa: BLE001
        return False


def safe_close(ctx):
    try:
        ctx.close()
    except Exception:  # noqa: BLE001
        pass


def keep_open(ctx, page):
    log("\n=== المتصفح مفتوح - سيبه زي ما هو، وقفله انت لما تخلص ===")
    try:
        while page_alive(page):
            page.wait_for_timeout(3000)
    except KeyboardInterrupt:
        log("\n(Ctrl+C)")
    safe_close(ctx)


def pick_channel():
    for ch, exe in (
        ("chrome", r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
        ("chrome", r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
        ("msedge", r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
        ("msedge", r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
    ):
        if os.path.isfile(exe):
            return ch
    return None


# ---------------------------------------------------------------- PDF reading
def pdf_pages_text(path, limit=None):
    from pypdf import PdfReader

    reader = PdfReader(path)
    pages = []
    for i, page in enumerate(reader.pages):
        if limit is not None and i >= limit:
            break
        try:
            pages.append(page.extract_text() or "")
        except Exception:  # noqa: BLE001
            pages.append("")
    return pages


def type_from_map(filename):
    """Business rule: PDF name pattern -> SmartCloud Task Type."""
    if not os.path.isfile(TYPE_MAP_FILE):
        return None
    try:
        with open(TYPE_MAP_FILE, encoding="utf-8") as f:
            rules = json.load(f)
    except Exception:  # noqa: BLE001
        return None
    name = os.path.basename(filename)
    for pattern, ttype in rules.items():
        try:
            if re.match(pattern, name, re.IGNORECASE):
                return ttype
        except re.error:
            if pattern.lower() in name.lower():
                return ttype
    return None


RFI_NO_RE = re.compile(r"RFI\s*(?:No\.?|Number|#)?\s*[:\-]?\s*([A-Z]{2,5}-RFI-[A-Z0-9\-/]+)", re.I)


def read_rfi_no(path):
    """The RFI number printed on the form, e.g. CPP-RFI-68-64-11-0336.

    This is the value the platform stores in the record Caption, so it is the
    only reliable way to tell which task page a PDF belongs to.
    """
    from pypdf import PdfReader

    reader = PdfReader(path)
    for page in reader.pages[:2]:
        m = RFI_NO_RE.search(page.extract_text() or "")
        if m:
            return m.group(1).strip().rstrip(".,")
    return None


def _repair_split_tags(text):
    """The forms wrap long tag lists across lines, which cuts tags in half. The breaks
    show up in three shapes, so repair all of them:
      '...0003B-\\nCA04'            -> the hyphen is at the end of the line
      '...CC01,\\n08-85BT-1901-CN01' -> the fragment lost its own 'PS5-' prefix
      '...,\\nPS5\\n85BT-1904-CJ01'  -> the 'PS5' got separated from the rest
    """
    flat = re.sub(r"[ \t]*\r?\n[ \t]*", " ", text or "")
    # a bare 'PS5' that lost the rest of the tag to the next line
    flat = re.sub(r"\bPS5\s+(?=\d{2}-[A-Z0-9])", "PS5-", flat)
    flat = re.sub(r"\bPS5\s+(?=[A-Z]{2,4}-)", "PS5-", flat)
    flat = re.sub(r"\bPS5\s+(?=\d{2}[A-Z0-9]{2}-)", "PS5-", flat)
    # a fragment that lost its own 'PS5-' prefix
    flat = re.sub(r"(?<=[\s,;])(\d{2}-[A-Z0-9]+(?:-[A-Z0-9]+){2,})\b", r"PS5-\1", flat)
    # a tag cut after a hyphen: 'PS5-60-IE-0003B- CA04' has to become 'PS5-60-IE-0003B-CA04'
    glued = re.sub(r"-\s+(?=[A-Z0-9]{2,}\b)", "-", flat)
    glued = re.sub(r"(?<=[A-Z0-9])\s+(?=[A-Z0-9]{2,}-)", "-", glued)
    # a break right after the prefix can leave it doubled: PS5-PS5-60-...
    glued = re.sub(r"\b(PS\d|PR\d)-(?:PS\d-|PR\d-)+", r"\1-", glued)
    return _split_glued_tags(glued)


TAG_RE = re.compile(r"\b(?:PS|PR)\d-[A-Z0-9]+(?:-[A-Z0-9]+){3,}\b")


def _split_glued_tags(text):
    """A line break between two tags welds them together with a single hyphen
    ('PS5-71-BI-0005-CC02-PS5-71-BI-0003-CC01'). Split every run of tag characters
    that contains more than one 'PSn-' / 'PRn-' prefix."""
    def fix(m):
        s = m.group(0)
        starts = [x.start() for x in re.finditer(r"(?:PS|PR)\d-", s)]
        if len(starts) < 2:
            return s
        parts, prev = [], 0
        for st in starts[1:]:
            parts.append(s[prev:st].rstrip("-"))
            prev = st
        parts.append(s[prev:])
        return ", ".join(p for p in parts if p)
    return re.sub(r"(?:PS|PR)\d-[A-Z0-9]+(?:-[A-Z0-9]+)+", fix, text or "")


def _description_block(text):
    """Just the 'Inspection activity description' part of the form. The rest of the
    page is the tick-box list, which contains words like 'Test crew not available'
    that must not be read as the type of work. In some forms the description, the tag
    list and the tick-boxes are printed on top of each other, so stop at the first
    asset tag instead of relying on line breaks."""
    m = re.search(r"inspection\s+activity\s+description\s*:?\s*(.+?)$", text or "", re.S | re.I)
    if m:
        block = m.group(1)
    else:
        # Not every form prints the label: on some of them the description is the line
        # that follows the ITP item, on top of the tag list.
        tail = re.split(r"\bITP\s*item\b\s*:?\s*[0-9A-Za-z\-\.]+", text or "", maxsplit=1, flags=re.I)
        block = tail[1] if len(tail) > 1 else ""
    if not block.strip():
        return ""
    cut = re.search(r"(?:PS|PR)\d-[A-Z0-9]+(?:-[A-Z0-9]+)+", block)
    if cut:
        block = block[:cut.start()]
    block = re.split(r"\bcontractor\b|\bname\b\s+f\b|\btotal\b|\d+\s*tags?\b", block, flags=re.I)[0]
    return " ".join(ln.strip() for ln in block.splitlines() if ln.strip())


def _complete_short_tags(tags):
    """A tag cut at the line break can lose its numeric segment: 'PS5-85BT-1904-CJ01'.
    The rest of the list still carries it, so borrow it from the majority."""
    full = [t for t in tags if re.match(r"^[A-Z0-9]+-\d{2}-", t)]
    if not full:
        return tags
    seg = re.match(r"^[A-Z0-9]+-(\d{2})-", full[0]).group(1)
    out = []
    for t in tags:
        if re.match(r"^[A-Z0-9]+-[A-Z]{2,4}-", t):
            t = re.sub(r"^([A-Z0-9]+)-", r"\1-%s-" % seg, t)
        out.append(t)
    return out


def read_pdf(path):
    """Asset tag from page 1, document type from the pages or the mapping file."""
    pages = pdf_pages_text(path)
    first = _repair_split_tags(pages[0] if pages else "")
    tags = []
    for tag in ASSET_TAG_RE.findall(first):
        if tag not in tags:
            tags.append(tag)
    for parts in ASSET_TAG_SPACED_RE.findall(first):
        tag = "-".join(parts)
        if tag not in tags:
            tags.append(tag)
    tags = _complete_short_tags(tags)
    seen, uniq = set(), []
    for t in tags:
        k = t.upper()
        if k not in seen:
            seen.add(k)
            uniq.append(t)
    tags = uniq
    # A tag that is still cut in half after the repairs must never be opened silently.
    suspect = [t for t in tags if not TAG_RE.fullmatch(t)]
    full = "\n".join(pages).lower()
    first_low = first.lower()
    # Only the "Inspection activity description" block decides the type. The rest of
    # page 1 is the tick-box list ("Test crew not available", "Test equipment not
    # available" ...) and would otherwise force every RFI to Static Test.
    desc_low = _description_block(first).lower()
    hits = []
    for name, keys in TYPE_KEYWORDS:
        found = [k for k in keys if k in full]
        if found:
            hits.append((name, found))
    # The work description decides. The form title is only a last resort, because
    # every one of these forms is titled "Request for Inspection" no matter the job.
    # 0) the first code of the DESCRIPTION decides, e.g. "TPX13 - Motor LV" is a test
    #    (Static Test) and "CPX13 - ..." is a conformance check. Anything else falls
    #    back to the wording rules below.
    ptype, why = None, None
    desc = re.sub(r"^[^a-z0-9]+", "", desc_low)
    first_code = (desc.split()[0] if desc.split() else "")
    if re.match(r"^t[a-z]*\d", first_code):
        ptype, why = "Static Test", [first_code]
    elif re.match(r"^c[a-z]*\d", first_code):
        ptype, why = "Conformity Check", [first_code]
    if ptype is None and re.search(r"\btesting\b|\btest\b", desc_low):
        ptype, why = "Static Test", ["testing / test"]
    # glanding / termination is physical installation work, so it is a Conformity
    # Check even when the sentence starts with "Request to Witness".
    if ptype is None and re.search(r"\bgland\w*|\bterminat\w*", desc_low):
        ptype, why = "Conformity Check", ["glanding / termination"]
    if ptype is None and re.search(r"\bvisual\s+inspection\b|\binspection\s+of\b", desc_low):
        ptype, why = "Conformity Check", ["visual inspection / inspection of"]
    if ptype is None:
        for name, keys in _load_work_type_rules():
            found = [k for k in keys if k in desc_low]
            if found:
                ptype, why = name, found
                break
    if ptype is None:
        for keys, name in (("request to witness", "Static Test"),
                           ("request for test", "Static Test"),
                           ("visual inspection", "Conformity Check"),
                           ("request for inspection", "Conformity Check"),
                           ("inspection request", "Conformity Check")):
            if any(k in desc_low for k in keys.split("|")):
                ptype, why = name, [k for k in keys.split("|") if k in desc_low]
                break
    if ptype is None:
        for name, found in hits:
            ptype, why = name, found
            break
    hits.insert(0, ("work description", why or []))
    mapped = type_from_map(path)
    return {
        "pdf": os.path.basename(path),
        "path": os.path.abspath(path),
        "pages": len(pages),
        "rfi_no": read_rfi_no(path),
        "asset_tags": tags,
        "suspect_tags": suspect,
        "detected_type": ptype,
        "mapped_type": mapped,
        "type_hits": hits,
    }


# ------------------------------------------------------------- task lookup
_task_index = None


def task_index():
    """TaskName/TaskType per Asset Tag. The live index the sync refreshes every cycle
    always wins, so the plan never runs on the old offline snapshot."""
    global _task_index
    if _task_index is not None:
        return _task_index
    src = TASK_CACHE
    if os.path.isfile(LIVE_TASK_INDEX) and (
        not os.path.isfile(TASK_CACHE)
        or os.path.getmtime(LIVE_TASK_INDEX) > os.path.getmtime(TASK_CACHE)
    ):
        src = LIVE_TASK_INDEX
    if not os.path.isfile(src):
        raise SystemExit(
            "مش لاقي %s\nشغّل التحديث الأول (PS5 ITR Update) عشان يجيب الـ tasks من المنصة." % src
        )
    with open(src, encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict) and "tags" in data:
        _task_index = {k: list(v) for k, v in (data.get("tags") or {}).items()}
        log("  (task index: %s - %d assets, محدّث %s)"
            % (os.path.basename(src), len(_task_index), data.get("updated", "?")))
        return _task_index
    rows = data.get("rows") if isinstance(data, dict) else data
    idx = {}
    for r in rows or []:
        tag = (r.get("AssetTag") or "").strip()
        if not tag:
            continue
        idx.setdefault(tag, []).append(
            {
                "task_id": r.get("TaskName"),
                "type": r.get("TaskType"),
                "state": r.get("TaskState"),
                "category": r.get("TaskCategorySummary"),
                "id": r.get("ID"),
            }
        )
    _task_index = idx
    log("  (task cache: %s - %d assets, آخر تحديث %s)"
        % (os.path.basename(src), len(idx),
           time.strftime("%Y-%m-%d", time.localtime(os.path.getmtime(src)))))
    return idx


def find_tasks(asset_tag, want_type=None):
    rows = task_index().get(asset_tag, [])
    if want_type:
        rows = [r for r in rows if (r.get("type") or "").lower() == want_type.lower()]
    return rows


def choose_task(rows, explicit=None):
    if explicit:
        for r in rows:
            if r.get("task_id") == explicit:
                return r
        return None
    if not rows:
        return None
    for r in rows:
        if str(r.get("state") or "").lower() not in ("closed", "completed"):
            return r
    return rows[0]


# ------------------------------------------------------------------ browser
def wait_products(page, sec=60):
    t0 = time.time()
    while time.time() - t0 < sec:
        try:
            n = page.evaluate(
                """() => (typeof switchboardConfiguration!=="undefined" && switchboardConfiguration.Products
                && switchboardConfiguration.Products.length) ? switchboardConfiguration.Products.length : -1"""
            )
            if n >= 0:
                return True
        except Exception:  # noqa: BLE001
            pass
        try:
            page.wait_for_timeout(2000)
        except Exception:  # noqa: BLE001
            time.sleep(2)
    return False


LOGIN_HINT = re.compile(r"(sign\s*in|log\s*in|password|saml|username)", re.I)


def login_form_visible(page):
    try:
        pw = page.locator('input[type="password"]')
        if pw.count() and pw.first.is_visible():
            return True
    except Exception:  # noqa: BLE001
        return False
    try:
        return bool(LOGIN_HINT.search(page.title() or ""))
    except Exception:  # noqa: BLE001
        return False


def ensure_logged_in(page, timeout=600):
    if not login_form_visible(page):
        return True
    log("\n=== الصفحة دي تسجيل دخول. سجّل دخولك في المتصفح بنفسك ===")
    log("    مستنيك... (مش هعمل حاجة غير كده)")
    t0 = time.time()
    tick = 0
    while time.time() - t0 < timeout:
        if not page_alive(page):
            log("  [!] المتصفح اتقفل - أعد تشغيل الأمر")
            return False
        try:
            page.wait_for_timeout(2000)
        except Exception:  # noqa: BLE001
            time.sleep(2)
        if not login_form_visible(page):
            log("  OK - دخلت.")
            return True
        tick += 1
        if tick % 15 == 0:
            log("    لسه مستني تسجيل الدخول... %ds" % int(time.time() - t0))
    log("  [!] انتهت المهلة")
    return False


def start(args):
    channel = pick_channel() if args.browser == "auto" else None
    if args.browser in ("chrome", "msedge"):
        channel = args.browser
    if args.browser == "chromium":
        channel = None
    busy = profile_locked()
    if busy:
        log("[!] في Chrome تاني شغال بنفس البروفايل (PID %s)." % ", ".join(busy))
        log("    اقفله بإيدك أو خلّيه شغال - متقفلشSolving force، عشان الـ session متيرويش.")
        raise SystemExit(1)
    os.makedirs(PROFILE_DIR, exist_ok=True)
    pw = sync_playwright().start()
    kw = dict(
        headless=args.headless,
        args=["--start-maximized"],
        accept_downloads=True,
        viewport=None if not args.headless else {"width": 1600, "height": 950},
    )
    if channel:
        kw["channel"] = channel
    log("المتصفح: %s" % (channel or "chromium (playwright)"))
    ctx = pw.chromium.launch_persistent_context(PROFILE_DIR, **kw)
    home = ctx.pages[0] if ctx.pages else ctx.new_page()
    home.set_default_timeout(25000)
    home.goto(SWITCHBOARD_URL, wait_until="domcontentloaded", timeout=60000)
    if not wait_products(home, 45):
        log("\n=== الجلسة منتهية. سجّل دخولك بإيدك في المتصفح (مرة واحدة بس) ===")
        home.bring_to_front()
        if not ensure_logged_in(home) or not wait_products(home, 180):
            log("NO SESSION -/login مانفعش. جرّب تاني بعد شوية (متكررش).")
            safe_close(ctx)
            return None
    log("الجلسة حية.")
    return pw, ctx, home


def open_testforms(ctx, home, url=TESTFORMS_URL):
    app = ctx.new_page()
    app.set_default_timeout(25000)
    app.goto(url, wait_until="commit", timeout=60000)
    for _ in range(40):
        app.wait_for_timeout(3000)
        try:
            if app.evaluate(
                """() => { const g = document.querySelector('#%s_PrimaryList_grid');
                    if (!g) return false; const kg = $(g).data('kendoGrid'); return !!kg; }"""
                % VIEW
            ):
                return app
        except Exception:  # noqa: BLE001
            pass
    log("  [!] جريد الصفحة مش ظهر - هكمّل على أي حال")
    return app


def fill_box(app, selector, value):
    app.evaluate(
        """(A) => {
            const el = document.querySelector(A.sel);
            if (!el) throw new Error('not found: ' + A.sel);
            el.scrollIntoView({block: 'center'});
            el.focus();
            el.value = A.val;
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
            el.blur();
        }""",
        {"sel": selector, "val": value},
    )
    log("  fill %s = %s" % (selector, value))


def do_search(app, timeout=120):
    app.evaluate("(s) => { const b = document.querySelector(s); if (b) b.click(); }", BTN_SEARCH)
    log("  SEARCH")
    t0 = time.time()
    last = -1
    while time.time() - t0 < timeout:
        app.wait_for_timeout(3000)
        try:
            total = app.evaluate(
                """() => { const g = document.querySelector('#%s_PrimaryList_grid');
                    if (!g) return -2; const kg = $(g).data('kendoGrid');
                    if (!kg) return -3; return kg.dataSource.total(); }"""
                % VIEW
            )
        except Exception:  # noqa: BLE001
            continue
        if total != last:
            log("    results: %s" % total)
            last = total
        if total and total > 0:
            return total
    return last


def select_first_row(app):
    return app.evaluate(
        """() => {
            const g = document.querySelector('#%s_PrimaryList_grid');
            if (!g) return 'no grid';
            const kg = $(g).data('kendoGrid');
            if (!kg) return 'no kendo';
            const view = kg.dataSource.view();
            const row = view.at(0) || view[0];
            if (!row) return 'no row';
            const tr = kg.tbody.find('tr[data-uid="' + row.uid + '"]');
            if (!tr || !tr.length) return 'row not rendered';
            kg.select(tr);
            const sel = kg.select().length;
            return sel > 0 ? 'ok:' + sel : 'select failed';
        }"""
        % VIEW
    )


def attach_file(app, pdf, timeout=60):
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            fi = app.locator('input[type="file"]')
            if fi.count():
                fi.first.set_input_files(pdf)
                log("  PDF attached: %s" % os.path.basename(pdf))
                return True
        except Exception:  # noqa: BLE001
            pass
        app.wait_for_timeout(2000)
    return False


def dump_grid(app, limit=10):
    return app.evaluate(
        """(N) => { const g = document.querySelector('#%s_PrimaryList_grid');
            if (!g) return []; const kg = $(g).data('kendoGrid'); if (!kg) return [];
            return kg.dataSource.data().slice(0, N).map(r => { const o = {};
                for (const k of Object.keys(r)) { const v = r[k]; if (v === null || typeof v !== 'object') o[k] = v; }
                return o; }); }"""
        % VIEW,
        limit,
    )


# ------------------------------------------------------------------- modes
def resolve_type(info, override=None):
    if override:
        return override, "hand"
    if info.get("mapped_type"):
        return info["mapped_type"], "map"
    return info.get("detected_type"), "text"


def cmd_scan(args):
    folder = args.dir
    files = sorted(
        os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".pdf")
    )
    log("=== %d PDF في %s ===" % (len(files), folder))
    found = False
    for path in files:
        info = read_pdf(path)
        found = True
        ttype, src = resolve_type(info, args.type)
        log("\n%s" % info["pdf"])
        log("  asset tag (صفحة 1): %s" % (", ".join(info["asset_tags"][:4]) or "!! مفيش"))
        log("  نوع الـ PDF       : %s (%s)" % (ttype or "!! مش متعرف", src))
        if info["type_hits"]:
            log("  كلمات لقيتها      : %s"
                % ", ".join("%s <- %s" % (n, ", ".join(k)) for n, k in info["type_hits"]))
        tag = info["asset_tags"][0] if info["asset_tags"] else None
        if not tag:
            continue
        log("  الأصل المستخدم    : %s" % tag)
        rows = find_tasks(tag, ttype)
        if not rows:
            log("  tasks: مفيش '%s' للأصل ده" % (ttype or "أي نوع"))
            for r in find_tasks(tag)[:8]:
                log("      متاح: %s | %s | %s" % (r["task_id"], r["type"], r["state"]))
            continue
        for r in rows[:12]:
            log("    %s | %s | %s" % (r["task_id"], r["type"], r["state"]))
        if len(rows) > 12:
            log("    ... و%d كمان" % (len(rows) - 12))
    if not found:
        log("مفيش PDF في الفولدر.")


def cmd_attach(args):
    path = os.path.abspath(args.pdf)
    if not os.path.isfile(path):
        raise SystemExit("الملف مش موجود: %s" % path)
    info = read_pdf(path)
    log("=== %s ===" % info["pdf"])
    log("  asset tag (صفحة 1): %s" % (", ".join(info["asset_tags"][:4]) or "!! مفيش"))
    ttype, src = resolve_type(info, args.type)
    log("  نوع الـ PDF       : %s (%s)" % (ttype or "!! مش متعرف", src))
    if not ttype:
        raise SystemExit("مش عارف نوع الـ PDF - استخدم --type أو ضيف قاعدة في rfi_type_map.json")
    tag = args.asset or info["asset_tags"][0]
    log("  هنستخدم الأصل    : %s" % tag)

    rows = find_tasks(tag, ttype)
    if not rows:
        log("  [!] مفيش Task ID من نوع '%s' للأصل %s" % (ttype, tag))
        for r in find_tasks(tag)[:12]:
            log("      متاح: %s | %s | %s" % (r["task_id"], r["type"], r["state"]))
        raise SystemExit(1)
    task = choose_task(rows, args.task)
    if task is None:
        log("  [!] الـ Task ID '%s' مش موجود لنفس النوع. المتاح:" % args.task)
        for r in rows[:12]:
            log("      %s | %s" % (r["task_id"], r["state"]))
        raise SystemExit(1)
    log("  Task ID           : %s  (%s / %s)" % (task["task_id"], task["type"], task["state"]))
    if len(rows) > 1 and not args.task:
        log("  (فيه %d task من نفس النوع - لو غلط، استخدم --task):" % len(rows))
        for r in rows[:12]:
            log("      %s | %s" % (r["task_id"], r["state"]))

    started = start(args)
    if not started:
        raise SystemExit(1)
    pw, ctx, home = started
    try:
        app = open_testforms(ctx, home, args.url)
        shot(app, "01_testforms")

        log("\n=== تعبئة الفورم ===")
        fill_box(app, BOX_ASSET, tag)
        fill_box(app, BOX_TASK, task["task_id"])
        app.wait_for_timeout(1500)
        total = do_search(app)
        if not total or total <= 0:
            shot(app, "02_no_results")
            log("  [!] البحث رجع صفر نتايج. بص على الصورة / جرب asset tag تاني.")
            if args.keep_open:
                keep_open(ctx, app)
            return 1
        rows_shown = dump_grid(app)
        log("  نتائج:")
        for r in rows_shown:
            log("    %s | %s | ExecutedForm=%s" % (r.get("Name"), r.get("AssetTypeSummary"), r.get("FileID")))
        for r in rows_shown:
            if r.get("FileID"):
                log("  [!] الصف ده فيه Executed Form أصلاً (FileID=%s) - الرفع الجديد هيغيبه."
                    % r.get("FileID"))
        shot(app, "03_search_results")

        sel = select_first_row(app)
        log("  selected row: %s" % sel)
        if not str(sel).startswith("ok"):
            log("  [!] مش قادر أختار الصف - اختره انت من المتصفح")

        log("\n=== Executed Forms ===")
        app.evaluate("(s) => { const b = document.querySelector(s); if (b) b.click(); }", BTN_EXECUTED)
        app.wait_for_timeout(5000)
        shot(app, "04_executed_forms")

        ok = attach_file(app, path)
        app.wait_for_timeout(2500)
        shot(app, "05_attached" if ok else "05_no_file_input")
        log("\n=== خلصنا ===")
        log("  Task ID : %s" % task["task_id"])
        log("  PDF     : %s" % ("اتلحق" if ok else "!! مفيش input[type=file] - شوف الصورة"))
        log("  المتصفح مفتوح: بص، وبعدين انت اللي تدوس Save.")
        if args.keep_open:
            keep_open(ctx, app)
        else:
            pause("\nلما تخلص، Enter عشان أقفل... ")
    finally:
        safe_close(ctx)
        try:
            pw.stop()
        except Exception:  # noqa: BLE001
            pass
    return 0


def cmd_inspect(args):
    started = start(args)
    if not started:
        raise SystemExit(1)
    pw, ctx, home = started
    try:
        app = open_testforms(ctx, home, args.url)
        shot(app, "inspect")
        cols = app.evaluate(
            """() => { const g = document.querySelector('#%s_PrimaryList_grid');
                if (!g) return []; const kg = $(g).data('kendoGrid'); if (!kg) return [];
                return kg.columns.filter(c => !c.hidden).map(c => c.field + ' = ' + (c.title||'').replace(/<[^>]+>/g,'')); }"""
            % VIEW
        )
        log("\ncolumns: " + ", ".join(cols))
        if args.keep_open:
            keep_open(ctx, app)
        else:
            pause("\nEnter عشان أقفل... ")
    finally:
        safe_close(ctx)
        try:
            pw.stop()
        except Exception:  # noqa: BLE001
            pass


def main():
    ap = argparse.ArgumentParser(description="SmartCloud Test Form PDF attach")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("scan", help="اعرض الـ asset tag والنوع والـ Task IDs لكل PDF في الفولدر")
    p.add_argument("--dir", default=WATCH_DIR)
    p.add_argument("--type", help="نوع ثابت يتخطى الكشف التلقائي")
    p.set_defaults(func=cmd_scan)

    p = sub.add_parser("attach", help="املأ الفورم واعمل attach للـ PDF (من غير حفظ)")
    p.add_argument("pdf")
    p.add_argument("--asset", help="asset tag يدوي (لو مش عايز أول واحد في الـ PDF)")
    p.add_argument("--task", help="Task ID يدوي (لو فيه أكتر من واحد من نفس النوع)")
    p.add_argument("--type", help="نوع الـ PDF يدوي: Static Test / Conformity Check / ...")
    p.add_argument("--url", default=TESTFORMS_URL)
    p.add_argument("--keep-open", action="store_true")
    p.add_argument("--headless", action="store_true")
    p.add_argument("--browser", default="auto", choices=["auto", "chrome", "msedge", "chromium"])
    p.set_defaults(func=cmd_attach)

    p = sub.add_parser("inspect", help="اعرض أعمدة الجريد")
    p.add_argument("--url", default=TESTFORMS_URL)
    p.add_argument("--keep-open", action="store_true")
    p.add_argument("--headless", action="store_true")
    p.add_argument("--browser", default="auto", choices=["auto", "chrome", "msedge", "chromium"])
    p.set_defaults(func=cmd_inspect)

    args = ap.parse_args()
    sys.exit(args.func(args) or 0)


if __name__ == "__main__":
    main()
