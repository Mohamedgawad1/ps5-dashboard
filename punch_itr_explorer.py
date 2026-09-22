"""
PS5 SUBSYSTEM EXPLORER - Punch & ITR by Subsystem
Reads the latest 'PS-5 COMPLETIONS DPR SUMMERY -*.xlsx' and builds:
  PS5 SUBSYSTEM EXPLORER.xlsx
    Sheet 1 'SUBSYSTEM SEARCH' : dropdown -> all punches + all ITRs (FILTER formulas)
    Sheet 2 'PUNCH LIST'       : full detailed punch list (colored, autofilter)
    Sheet 3 'ITR LIST'         : full detailed ITR list (autofilter)
    Sheet 4 'SUBSYSTEM KEYS'   : master subsystem list + live counts (the key)
  subsystem_explorer.html      : standalone offline platform (same search)
Every daily rebuild carries over USER NOTES + any custom colors from the previous file.
"""
import os, re, glob, shutil, datetime, json, html as html_mod
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from dpr_new_summary import (extract_itr as sum_itr, extract_punch as sum_punch,
                             extract_milestones)

BASE = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(BASE, '_temp')
HOME_DL = os.path.join(os.path.expanduser('~'), 'Downloads')
DL_SUB = os.path.join(HOME_DL, 'PS5 - CPP AGI Completion Progress Dashboard_files')
OUT_XLSX = os.path.join(DL_SUB, 'PS5 SUBSYSTEM EXPLORER.xlsx')
OUT_HTML = os.path.join(DL_SUB, 'subsystem_explorer.html')
KEY_FILE = os.path.join(DL_SUB, '_explorer_state.json')
PCOLORS = ['', 'FFFFEB9C', 'FFC6EFCE', 'FFFFC7CE', 'FFBDD7EE']


def load_platform_state():
    """Shared state written by the web platform (explorer_server.py)."""
    st = {'notes': {}, 'colors': {}, 'cells': {}, 'rmk': {}}
    try:
        with open(KEY_FILE, encoding='utf-8') as f:
            raw = json.load(f)
        st['notes'] = raw.get('notes', {}) or {}
        st['colors'] = raw.get('colors', {}) or {}
        st['cells'] = raw.get('cells', {}) or {}
        st['rmk'] = raw.get('rmk', {}) or {}
    except Exception:
        pass
    return st


def save_platform_state(st):
    tmp = KEY_FILE + '.tmp'
    try:
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(st, f, ensure_ascii=False, indent=1)
        os.replace(tmp, KEY_FILE)
    except Exception:
        pass

NAVY, BLUE, ORANGE = 'FF003366', 'FF4472C4', 'FFED7D31'
WHITE, GRAY = 'FFFFFFFF', 'FFF2F2F2'
GREEN, RED, YELLOW, DGREEN = 'FFC6EFCE', 'FFFFC7CE', 'FFFFEB9C', 'FFD9D9D9'
SEL_FILL = 'FFFFE699'
OWN_COLORS = {NAVY, BLUE, ORANGE, WHITE, GRAY, GREEN, RED, YELLOW, DGREEN,
              SEL_FILL, 'FFDDEBF7', '00000000'}

thin = Side(style='thin', color='FFBFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

PUNCH_COLS = ['ID', 'TAG', 'CAT', 'DISC', 'DESCRIPTION', 'RESP COMPANY',
              'SUB SYSTEM', 'CLOSING DATE', 'STATUS', 'ORIGINATED BY',
              'PRIORITY', 'USER NOTES']
ITR_COLS = ['ID', 'TAG', 'CATEGORY', 'DISC', 'TASK TYPE', 'ASSET DESCRIPTION',
            'RESP COMPANY', 'SUB SYSTEM', 'SUB ID', 'CLOSING DATE', 'STATE',
            'M&P', 'EIT', 'RFC', 'PRIORITY', 'TEST FORM', 'USER NOTES']


def n(v):
    if v is None:
        return ''
    s = str(v).strip()
    if s in ('#N/A', '#VALUE!', 'nan', 'None'):
        return ''
    return s


def dstr(v):
    if isinstance(v, datetime.datetime):
        return v.strftime('%d-%b-%y')
    return n(v)


def find_latest_dpr():
    cands = []
    seen = set()
    for folder in (DL_SUB, HOME_DL, BASE):
        if not os.path.isdir(folder):
            continue
        for f in glob.glob(os.path.join(folder, '*.xlsx')):
            b = os.path.basename(f)
            if b.startswith('~$') or b in seen:
                continue
            if 'completions dpr summery' not in b.lower():
                continue
            if 'backup' in b.lower():
                continue
            seen.add(b)
            m = re.search(r'-(\d{2}-\d{2}-\d{2})', b)
            key = None
            if m:
                try:
                    key = datetime.datetime.strptime(m.group(1), '%d-%m-%y')
                except ValueError:
                    pass
            if key is None:
                key = datetime.datetime.fromtimestamp(os.path.getmtime(f))
            cands.append((key, f))
    return max(cands)[1] if cands else None


def open_source(path):
    dst = os.path.join(TMP, os.path.basename(path))
    last_err = None
    for i in range(6):
        try:
            shutil.copy2(path, dst)
            break
        except PermissionError as e:
            last_err = e
            if os.path.exists(dst) and os.path.getsize(dst) == os.path.getsize(path):
                break
            import time
            time.sleep(1.5 * (i + 1))
    else:
        if not os.path.exists(dst):
            raise last_err or PermissionError(path)
    return openpyxl.load_workbook(dst, read_only=True, data_only=True)


# ---------------- extraction ----------------

def extract_punch(ws):
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        rid = n(row[0])
        if not rid or rid == 'Punchlist ID':
            continue
        out.append({
            'id': rid, 'tag': n(row[1]), 'cat': n(row[3]), 'disc': n(row[4]),
            'desc': n(row[5]), 'resp': n(row[6]), 'sub': n(row[7]),
            'close': dstr(row[8]), 'status': n(row[9]),
            'orig': n(row[10]), 'prio': n(row[11]) or n(row[12]),
        })
    return out


def extract_itr(ws):
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        rid = n(row[0])
        if not rid or rid == 'Task ID':
            continue
        out.append({
            'id': rid, 'tag': n(row[1]), 'category': n(row[7]), 'disc': n(row[8]),
            'ttype': n(row[9]), 'adesc': n(row[11]), 'resp': n(row[13]),
            'sub': n(row[21]), 'close': dstr(row[22]), 'subid': n(row[23]),
            'testform': n(row[26]), 'state': n(row[28]),
            'mp': dstr(row[30]), 'eit': dstr(row[31]), 'rfc': dstr(row[32]),
            'prio': n(row[38]),
        })
    return out


def build_keys(punch, itr):
    names, by_id = [], {}
    for r in punch:
        s = r['sub']
        if not s:
            continue
        sid = s.split(' - ')[0]
        if s not in by_id:
            by_id[s] = {'name': s, 'sid': sid, 'p': 0, 'i': 0}
            names.append(s)
        by_id[s]['p'] += 1
    for r in itr:
        s = r['sub']
        if not s:
            continue
        sid = r['subid'] or s.split(' - ')[0]
        if s not in by_id:
            by_id[s] = {'name': s, 'sid': sid, 'p': 0, 'i': 0}
            names.append(s)
        by_id[s]['i'] += 1
    names.sort(key=lambda s: (by_id[s]['sid'], s))
    return [by_id[s] for s in names]


def is_done_punch(rec):
    return rec['status'].lower() == 'closed'


RFC_DISCS = ['B', 'E', 'H', 'I', 'M', 'P', 'S', 'T']


def extract_rfc(ws):
    """RFC PROGRESS sheet -> per-subsystem dict incl. blocking ITR count."""
    out = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        name = str(row[0]).strip() if row[0] else ''
        if not name.startswith('PS5'):
            continue

        def num(i, _r=row):
            try:
                return float(_r[i])
            except (TypeError, ValueError):
                return None

        def dtok(i, _r=row):
            v = _r[i]
            if isinstance(v, datetime.datetime):
                return v.strftime('%d-%b-%y')
            return v if v is not None else ''

        d = {}
        for bi, key in enumerate(RFC_DISCS):
            base = 9 + bi * 4
            d[key] = {'t': num(base), 'c': num(base + 1),
                      'b': num(base + 2), 'p': num(base + 3)}
        out.append({
            'name': name, 'sid': name.split(' - ')[0],
            'prio': n(row[1]), 'rfc_bhm': dtok(2), 'rfc_eit': dtok(3),
            'baseline': dtok(4), 'recovery': dtok(5), 'milestone': n(row[7]),
            'tot_pct': num(8), 'disc': d,
            'itrs': num(41), 'closed': num(42), 'balance': num(43),
            'itr_pct': num(44),
            'blk_cpp1': num(49), 'blk_eit': num(50), 'blk_eacop': num(51),
            'rem_eacop': n(row[52]), 'rem_cpeit': n(row[53]), 'rem_cpp1': n(row[54]),
            'walkdown': n(row[55]), 'signed': dtok(6),
        })
    try:
        cells = load_platform_state().get('cells', {}).get('RFC PROGRESS', {})
        for rec in out:
            o = cells.get(rec['sid'])
            if not o:
                continue
            for ci, key, lab in RFC_COLMAP:
                if not lab or lab not in o:
                    continue
                v = o[lab]
                if isinstance(key, tuple):
                    try:
                        rec['disc'][key[1]][key[2]] = float(v)
                    except (TypeError, ValueError):
                        pass
                elif key.startswith('blk_') or key in ('itrs', 'closed', 'balance'):
                    try:
                        rec[key] = float(v)
                    except (TypeError, ValueError):
                        pass
                else:
                    rec[key] = v
    except Exception:
        pass
    return out


def find_ov_files():
    """Latest UNATRAC exports: ovPunchlist_*.xlsx + ovTasks_TestsPlanned_*.xlsx"""
    out = {'punch': None, 'itr': None}
    seen = set()
    for folder in (HOME_DL, DL_SUB, BASE):
        if not os.path.isdir(folder):
            continue
        for f in glob.glob(os.path.join(folder, 'ov*.xlsx')):
            b = os.path.basename(f)
            if b.startswith('~$') or b in seen:
                continue
            seen.add(b)
            low = b.lower()
            mt = os.path.getmtime(f)
            if 'punchlist' in low:
                if out['punch'] is None or mt > os.path.getmtime(out['punch']):
                    out['punch'] = f
            elif 'tasks' in low:
                if out['itr'] is None or mt > os.path.getmtime(out['itr']):
                    out['itr'] = f
    return out


def _odate(v):
    s = n(v)
    if not s:
        return ''
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', s)
    if m:
        try:
            return datetime.datetime.strptime(m.group(0), '%Y-%m-%d').strftime('%d-%b-%y')
        except ValueError:
            return ''
    return s


def extract_ov_punch(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    hdr = [n(h) for h in next(it)]
    idx = {h: i for i, h in enumerate(hdr)}
    g = lambda r, k: r[idx[k]] if k in idx and idx[k] < len(r) else None
    out = {}
    for r in it:
        rid = n(g(r, 'Punchlist ID'))
        if not rid:
            continue
        out[rid] = {
            'id': rid,
            'tag': n(g(r, 'Asset (Name/Tag)')),
            'cat': n(g(r, 'Punchlist Category (Name)')),
            'disc': n(g(r, 'Discipline (Name)')),
            'desc': n(g(r, 'Description')),
            'resp': n(g(r, 'Scheduling - Responsible Company (Name)')),
            'sub': n(g(r, 'Systemization - Subsystem (Summary)')),
            'close': _odate(g(r, 'Completed Date')) or _odate(g(r, 'Workflow - Closing Date')),
            'status': n(g(r, 'Status')),
            'orig': n(g(r, 'Originated By')),
            'prio': n(g(r, 'Priority')),
        }
    wb.close()
    return out


def extract_ov_itr(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    hdr = [n(h) for h in next(it)]
    idx = {h: i for i, h in enumerate(hdr)}

    def g(r, k):
        return r[idx[k]] if k in idx and idx[k] < len(r) else None
    out = {}
    for r in it:
        rid = n(g(r, 'Task ID'))
        if not rid:
            continue
        out[rid] = {
            'id': rid,
            'tag': n(g(r, 'Asset - Tag')),
            'category': n(g(r, 'Category (Summary)')),
            'disc': n(g(r, 'Discipline')) or n(g(r, 'Discipline (Summary)')),
            'ttype': n(g(r, 'Task Type (Name)')),
            'adesc': n(g(r, 'Asset - Description')),
            'resp': n(g(r, 'Responsible Company (Summary)')),
            'sub': n(g(r, 'Systemization - Subsystem (Summary)')),
            'subid': n(g(r, 'Subsystem ID')),
            'close': _odate(g(r, 'Closing Date')),
            'state': n(g(r, 'Task State')),
            'mp': '', 'eit': '', 'rfc': '',
            'prio': n(g(r, 'Task Priority')) or n(g(r, 'Priority')),
            'testform': n(g(r, 'Test Form - Description')),
        }
    wb.close()
    return out


def merge_with_ov(rows, ov):
    """Overlay fresher UNATRAC values onto DPR rows + append rows missing."""
    if not ov:
        return rows, 0
    upd = added = 0
    have = {r.get('id'): r for r in rows}
    for rid, o in ov.items():
        if rid in have:
            tgt = have[rid]
            for k, v in o.items():
                if v and str(v).strip() and str(tgt.get(k, '')) != str(v):
                    tgt[k] = v
                    upd += 1
        else:
            rows.append(dict(o))
            have[rid] = rows[-1]
            added += 1
    return rows, upd + added


def is_done_itr(rec):
    s = rec['state'].lower()
    return ('complete' in s) or (s == 'closed')


# ---------------- carry-over (notes + custom colors) ----------------

def load_previous(path=None):
    path = path or OUT_XLSX
    state = {'notes': {}, 'fills': {}, 'pcolors': {}}
    plat = load_platform_state()
    state['pcolors'] = plat.get('colors', {})
    state['cells'] = plat.get('cells', {})
    if not os.path.exists(path):
        return state
    try:
        wb = openpyxl.load_workbook(path)
    except Exception:
        return state
    for sh in ('PUNCH LIST', 'ITR LIST'):
        if sh not in wb.sheetnames:
            continue
        ws = wb[sh]
        note_col = None
        for c in range(1, ws.max_column + 1):
            if n(ws.cell(1, c).value).upper() == 'USER NOTES':
                note_col = c
                break
        if note_col is None:
            continue
        notes, fills = {}, {}
        for r in range(2, ws.max_row + 1):
            rid = n(ws.cell(r, 1).value)
            if not rid:
                continue
            nv = n(ws.cell(r, note_col).value)
            if nv:
                notes[rid] = nv
            row_fills = {}
            for c in range(1, ws.max_column + 1):
                f = ws.cell(r, c).fill
                if f is None or f.patternType != 'solid':
                    continue
                rgb = getattr(f.fgColor, 'rgb', None)
                if isinstance(rgb, str) and rgb not in OWN_COLORS and len(rgb) == 8:
                    row_fills[c] = rgb
            if row_fills:
                fills[rid] = row_fills
        state['notes'][sh] = notes
        state['fills'][sh] = fills
    wb.close()
    for sh in ('PUNCH LIST', 'ITR LIST'):
        pn = plat['notes'].get(sh)
        if pn:
            merged = dict(state['notes'].get(sh) or {})
            merged.update(pn)
            state['notes'][sh] = merged
    return state


# ---------------- styling helpers ----------------

def style_title(ws, row, ncols, text, size=14):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row, 1, text)
    c.font = Font(bold=True, color=WHITE, size=size)
    c.fill = PatternFill('solid', start_color=NAVY, end_color=NAVY)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[row].height = 24 if size >= 14 else 18


def hdr_cell(c, fill=BLUE):
    c.font = Font(bold=True, color=WHITE, size=10)
    c.fill = PatternFill('solid', start_color=fill, end_color=fill)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = BORDER


def cat_fill(cat):
    u = cat.upper()
    return {'A': RED, 'B': YELLOW, 'C': GREEN}.get(u, DGREEN if u == '0' else None)


# ---------------- sheet builders ----------------

def build_search(wb, keys, rep_date, src_name, punch, itr, max_p, max_i):
    ws = wb.create_sheet('SUBSYSTEM SEARCH')
    style_title(ws, 1, 12, 'PS5  PUNCH & ITR  -  SUBSYSTEM EXPLORER')
    sub = ws.cell(2, 1, f'Report Date: {rep_date}    |    Source: {src_name}')
    sub.font = Font(italic=True, size=9, color='FF404040')
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=12)
    sub.alignment = Alignment(horizontal='center')

    lbl = ws.cell(4, 1, 'SELECT SUB SYSTEM:')
    lbl.font = Font(bold=True, size=12, color=NAVY)
    ws.merge_cells(start_row=4, start_column=2, end_row=4, end_column=5)
    sel = ws.cell(4, 2)
    sel.fill = PatternFill('solid', start_color=SEL_FILL, end_color=SEL_FILL)
    sel.font = Font(bold=True, size=12)
    sel.alignment = Alignment(horizontal='center', vertical='center')
    sel.border = Border(*[Side(style='medium', color=ORANGE)] * 4)
    ws.row_dimensions[4].height = 24
    tip = ws.cell(5, 1, 'Excel 365: results auto-fill below. Older Excel: use the arrows (AutoFilter) on PUNCH LIST / ITR LIST sheets.')
    tip.font = Font(italic=True, size=8, color='FF808080')
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=12)

    pl_last, il_last = len(punch) + 1, len(itr) + 1
    stats = [
        ('PUNCH TOTAL', f"=COUNTIF('PUNCH LIST'!$G:$G,$B$4)", NAVY),
        ('CLOSED', f"=COUNTIFS('PUNCH LIST'!$G:$G,$B$4,'PUNCH LIST'!$I:$I,\"Closed\")", GREEN),
        ('OPEN', f"=B6-D6", ORANGE),
        ('ITR TOTAL', f"=COUNTIF('ITR LIST'!$H:$H,$B$4)", NAVY),
        ('DONE', f"=COUNTIFS('ITR LIST'!$H:$H,$B$4,'ITR LIST'!$K:$K,\"Completed\")+COUNTIFS('ITR LIST'!$H:$H,$B$4,'ITR LIST'!$K:$K,\"Closed\")", GREEN),
        ('REMAINING', '=I6-K6', ORANGE),
    ]
    col = 1
    for lab, formula, fill in stats:
        a = ws.cell(6, col, lab)
        hdr_cell(a, fill=fill)
        b = ws.cell(6, col + 1, formula)
        hdr_cell(b, fill=fill)
        b.font = Font(bold=True, size=11, color=WHITE)
        col += 2
    ws.row_dimensions[6].height = 20

    r = 8
    style_title(ws, r, 11, 'PUNCH ITEMS FOR SELECTED SUBSYSTEM', size=11)
    ws['A' + str(r)].fill = PatternFill('solid', start_color=ORANGE, end_color=ORANGE)
    r += 1
    p_hdr = r
    for ci, h in enumerate(PUNCH_COLS[:-1], 1):
        hdr_cell(ws.cell(r, ci), fill=NAVY)
        ws.cell(r, ci).value = h
    r += 1
    p_spill = r
    ws.cell(r, 1).value = (
        f"=IFERROR(FILTER('PUNCH LIST'!$A$2:$K${pl_last},"
        f"'PUNCH LIST'!$G$2:$G${pl_last}=$B$4),\"-> Select a Sub System above\")"
    )

    r = p_spill + max_p + 2
    style_title(ws, r, 11, 'ITR TASKS FOR SELECTED SUBSYSTEM', size=11)
    ws['A' + str(r)].fill = PatternFill('solid', start_color=ORANGE, end_color=ORANGE)
    r += 1
    for ci, h in enumerate(ITR_COLS[:-1], 1):
        hdr_cell(ws.cell(r, ci), fill=NAVY)
        ws.cell(r, ci).value = h
    i_spill = r + 1
    ws.cell(i_spill, 1).value = (
        f"=IFERROR(FILTER('ITR LIST'!$A$2:$P${il_last},"
        f"'ITR LIST'!$H$2:$H${il_last}=$B$4),\"-> Select a Sub System above\")"
    )
    end_anchor = i_spill + max_i + 2

    dv = DataValidation(type='list',
                        formula1=f"='SUBSYSTEM KEYS'!$A$2:$A${len(keys)+1}",
                        allow_blank=True, showDropDown=False)
    dv.prompt = 'Pick a subsystem'; dv.promptTitle = 'Sub System'
    ws.add_data_validation(dv)
    dv.add('B4')

    widths = [13, 22, 10, 14, 46, 15, 34, 12, 12, 18, 10]
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = 'A7'
    return end_anchor


def build_detail(wb, name, cols, rows, get_sub, state, done_fn, color_mode):
    ws = wb.create_sheet(name)
    style_title(ws, 1, len(cols), name + '  -  ALL DATA (use filter arrows to search)')
    for ci, h in enumerate(cols, 1):
        hdr_cell(ws.cell(2, ci), fill=NAVY)
        ws.cell(2, ci).value = h
    notes = state['notes'].get(name, {})
    fills = state['fills'].get(name, {})
    ovr = (state.get('cells', {}) or {}).get(name, {})
    r = 3
    for rec in rows:
        rid = rec.get('id', '')
        ov = ovr.get(rid, {})
        vals = []
        for h, k in zip(cols, REC_KEYS[name]):
            if k == 'id':
                vals.append(rec.get(k, ''))
            elif h == 'USER NOTES':
                vals.append(notes.get(rid, ''))
            elif h in ov:
                vals.append(ov[h])
            else:
                vals.append(rec.get(k, ''))
        done = done_fn(rec)
        band = GREEN if done else (GRAY if r % 2 else WHITE)
        for ci, v in enumerate(vals, 1):
            c = ws.cell(r, ci, v if v != '' else None)
            c.font = Font(size=9)
            c.alignment = Alignment(vertical='center',
                                    horizontal='center' if ci in CENTERED[name] else 'left',
                                    wrap_text=(ci in WRAPPED[name]))
            if done:
                c.fill = PatternFill('solid', start_color=GREEN, end_color=GREEN)
            else:
                c.fill = PatternFill('solid', start_color=band, end_color=band)
        if not done and color_mode == 'cat':
            cf = cat_fill(rec.get('cat', ''))
            if cf:
                ws.cell(r, 3).fill = PatternFill('solid', start_color=cf, end_color=cf)
        pcol = state.get('pcolors', {}).get(name, {}).get(rid)
        if pcol:
            for ci in range(1, len(cols) + 1):
                ws.cell(r, ci).fill = PatternFill('solid', start_color=pcol, end_color=pcol)
        for ci, rgb in fills.get(rid, {}).items():
            if 1 <= ci <= len(cols):
                ws.cell(r, ci).fill = PatternFill('solid', start_color=rgb, end_color=rgb)
        r += 1
    last = r - 1
    ws.auto_filter.ref = f'A2:{get_column_letter(len(cols))}{last}'
    widths = WIDTHS[name]
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = 'A3'
    return ws


def build_key_sheet(wb, keys):
    ws = wb.create_sheet('SUBSYSTEM KEYS')
    heads = ['SUB SYSTEM (KEY)', 'ID', 'PUNCH TOTAL', 'PUNCH CLOSED', 'PUNCH OPEN', 'ITR TOTAL']
    style_title(ws, 1, len(heads), 'SUBSYSTEM KEYS  -  MASTER LIST (dropdown source)')
    for ci, h in enumerate(heads, 1):
        hdr_cell(ws.cell(2, ci), fill=NAVY)
        ws.cell(2, ci).value = h
    r = 3
    for k in keys:
        vals = [k['name'], k['sid'],
                f"=COUNTIF('PUNCH LIST'!$G:$G,$A{r})",
                f"=COUNTIFS('PUNCH LIST'!$G:$G,$A{r},'PUNCH LIST'!$I:$I,\"Closed\")",
                f"=C{r}-D{r}",
                f"=COUNTIF('ITR LIST'!$H:$H,$A{r})"]
        band = GRAY if r % 2 else WHITE
        for ci, v in enumerate(vals, 1):
            c = ws.cell(r, ci, v)
            c.border = BORDER
            c.font = Font(size=9, bold=(ci == 1))
            c.fill = PatternFill('solid', start_color=band, end_color=band)
            if ci > 2:
                c.alignment = Alignment(horizontal='center')
        r += 1
    for ci, w in enumerate([40, 12, 13, 14, 12, 11], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = 'A3'
    ws.auto_filter.ref = f'A2:F{r-1}'
    return ws


REC_KEYS = {
    'PUNCH LIST': ['id', 'tag', 'cat', 'disc', 'desc', 'resp', 'sub',
                   'close', 'status', 'orig', 'prio', 'id'],
    'ITR LIST': ['id', 'tag', 'category', 'disc', 'ttype', 'adesc', 'resp',
                 'sub', 'subid', 'close', 'state', 'mp', 'eit', 'rfc',
                 'prio', 'testform', 'id'],
}

# ---- full-editable registry -------------------------------------------------
# (column_index_in_output_sheet, record_key_or_(disc,sub), unique_label)
RFC_DISCS_L = ['B', 'E', 'H', 'I', 'M', 'P', 'S', 'T']
RFC_COLMAP = [(2, 'prio', 'Priority'), (3, 'rfc_bhm', 'RFC BHMPS'),
              (4, 'rfc_eit', 'RFC EIT'), (5, 'baseline', 'Baseline'),
              (6, 'recovery', 'Recovery'), (7, 'signed', 'SIGNED'),
              (8, 'milestone', 'Milestone')]
for _bi, _k in enumerate(RFC_DISCS_L):
    _b = 10 + _bi * 4
    RFC_COLMAP += [(_b, ('disc', _k, 't'), _k + ' TOTAL'),
                   (_b + 1, ('disc', _k, 'c'), _k + ' CLOSED'),
                   (_b + 2, ('disc', _k, 'b'), _k + ' BAL')]
RFC_COLMAP += [(42, 'itrs', 'ITRs'), (43, 'closed', 'CLOSED'),
               (44, 'balance', 'BALANCE'),
               (50, 'blk_cpp1', 'CPP-1'), (51, 'blk_eit', 'EIT'),
               (52, 'blk_eacop', 'EACOP'),
               (53, 'rem_eacop', 'REMARK EACOP'),
               (54, 'rem_cpeit', 'REMARK CPP-EIT'),
               (55, 'rem_cpp1', 'REMARK CPP-1'),
               (56, 'walkdown', 'STATUS')]
DETAIL_COLMAP = {
    'PUNCH LIST': list(zip(range(2, 12), REC_KEYS['PUNCH LIST'][1:-1],
                           PUNCH_COLS[1:-1])),
    'ITR LIST': list(zip(range(2, 17), REC_KEYS['ITR LIST'][1:-1],
                         ITR_COLS[1:-1])),
}


def capture_cells(path, specs):
    """Detect manual Excel edits vs fresh source values and pin them into the
    shared platform state so they survive every rebuild."""
    st = load_platform_state()
    cells = st.setdefault('cells', {})
    if not os.path.exists(path):
        return st
    try:
        wb = openpyxl.load_workbook(path)
    except Exception:
        return st
    changed = False
    try:
        for name, spec in specs.items():
            if name not in wb.sheetnames:
                continue
            ws = wb[name]
            hrow = spec['hrow']
            fresh = {spec['idfn'](rec): rec for rec in spec['rows']}
            cc = cells.setdefault(name, {})
            for r in range(hrow + 1, ws.max_row + 1):
                rid = spec['rowid'](ws, r)
                if not rid or rid not in fresh:
                    continue
                rec = fresh[rid]
                o = cc.setdefault(rid, {})
                for ci, key, lab in spec['colmap']:
                    if not lab:
                        continue
                    pv = n(ws.cell(r, ci).value)
                    if isinstance(key, tuple):
                        sv = n((rec.get('disc') or {}).get(key[1], {}).get(key[2], ''))
                    else:
                        sv = n(rec.get(key, ''))
                    if pv != sv and pv:
                        if o.get(lab) != pv:
                            o[lab] = pv
                            changed = True
                    elif not pv and lab in o:
                        o.pop(lab, None)
                        changed = True
            for rid in [k for k, v in cc.items() if not v]:
                cc.pop(rid, None)
    finally:
        wb.close()
    if changed:
        save_platform_state(st)
    return st


def rfc_row_sid(ws, r):
    a = ws.cell(r, 1).value
    s = str(a or '')
    m = re.match(r'=HYPERLINK\("[^"]*#([^"]+)",".*"\)', s)
    if m:
        return m.group(1)
    m = re.match(r'\s*(PS5-\d{2}-\d{2}[A-Z0-9-]*)\s+-', s)
    return m.group(1) if m else ''
CENTERED = {
    'PUNCH LIST': {1, 3, 4, 8, 9, 11},
    'ITR LIST': {1, 4, 8, 9, 10, 11, 12, 13, 14, 15},
}
WRAPPED = {
    'PUNCH LIST': {5, 7},
    'ITR LIST': {6, 8},
}
WIDTHS = {
    'PUNCH LIST': [13, 22, 7, 12, 50, 16, 36, 12, 12, 20, 10, 28],
    'ITR LIST': [15, 22, 20, 8, 18, 32, 14, 34, 11, 12, 16, 11, 11, 11, 10, 20, 28],
}


def save_with_retry(wb):
    try:
        wb.save(OUT_XLSX)
        return OUT_XLSX
    except PermissionError:
        pass
    alt = os.path.join(TMP, os.path.basename(OUT_XLSX))
    wb.save(alt)
    for i in range(4):
        try:
            shutil.copy2(alt, OUT_XLSX)
            return OUT_XLSX
        except PermissionError:
            import time
            time.sleep(2)
    stamp = datetime.datetime.now().strftime('%H%M%S')
    final = OUT_XLSX.replace('.xlsx', f' {stamp}.xlsx')
    shutil.copy2(alt, final)
    return final


# ---------------- HTML platform ----------------


def apply_extras(page):
    """Platform toolbar + login + print/save + source-wins override rule."""
    if '#xlsbtn{' in page or 'id="lgmodal"' in page:
        return page
    old_ovv = 'function ovv(sh,id,col,def){const o=OVR[sh];const v=(o&&o[id])?o[id][col]:undefined;return (v===undefined||v===null)?def:v;}'
    new_ovv = 'function ovv(sh,id,col,def){const o=OVR[sh];const v=(o&&o[id])?o[id][col]:undefined;if(v===undefined||v===null)return def;if(def!==""&&def!==null&&def!==undefined)return def;return v;}'
    if old_ovv in page:
        page = page.replace(old_ovv, new_ovv)
    hdr = ' <div id="syncbadge">&#9679; SYNCED TO EXCEL</div>'
    btns = hdr + '\n <button id="xlsbtn" title="DOWNLOAD EXCEL">&#11015;</button><button id="prnbtn" title="PRINT">&#128424;</button><button id="svbtn" title="SAVE">&#128190;</button><button id="rfbtn" title="REFRESH / UPDATE">&#8635;</button><button id="lgbtn" title="SIGN IN TO EDIT">&#128274;</button>'
    if hdr in page:
        page = page.replace(hdr, btns)
    i = page.rfind('</body>')
    page = page[:i] + EXTRAS_BLOCK + '\n' + page[i:]
    return page


EXTRAS_BLOCK = r'''<style>
#xlsbtn,#prnbtn,#svbtn,#rfbtn,#lgbtn{width:32px;height:32px;display:inline-flex;align-items:center;justify-content:center;border:none;border-radius:50%;font-size:14px;cursor:pointer;margin-left:8px;flex:0 0 auto;transition:transform .15s ease,filter .15s ease}
#xlsbtn{background:linear-gradient(135deg,#ff9a4d,#ed7d31 50%,#d95b17);box-shadow:0 3px 10px -3px rgba(237,125,49,.8),inset 0 1px 0 rgba(255,255,255,.4);color:#fff}
#prnbtn{background:linear-gradient(135deg,#5b8dd6,#39679f 50%,#274a77);box-shadow:0 3px 10px -3px rgba(57,103,159,.8),inset 0 1px 0 rgba(255,255,255,.35);color:#fff}
#svbtn{background:linear-gradient(135deg,#2fbe6e,#1f9d55 50%,#136b3a);box-shadow:0 3px 10px -3px rgba(31,157,85,.8),inset 0 1px 0 rgba(255,255,255,.35);color:#fff}
#rfbtn{background:linear-gradient(135deg,#ffe27a,#f0a500 55%,#c47f00);box-shadow:0 3px 10px -3px rgba(240,165,0,.9),inset 0 1px 0 rgba(255,255,255,.5);color:#141414;font-weight:900;transition:transform .3s ease}
#rfbtn:hover{transform:rotate(120deg)}
#lgbtn{background:linear-gradient(135deg,#27405c,#16283c);color:#ffc000;box-shadow:0 3px 10px -3px rgba(0,0,0,.6),inset 0 1px 0 rgba(255,255,255,.15)}
#xlsbtn:hover,#prnbtn:hover,#svbtn:hover,#lgbtn:hover{transform:translateY(-1px);filter:brightness(1.15)}
#xlsbtn:active,#prnbtn:active,#svbtn:active,#rfbtn:active,#lgbtn:active{transform:scale(.93)}
#lgmodal{position:fixed;inset:0;background:rgba(3,10,18,.78);z-index:3000;display:none;align-items:center;justify-content:center}
#lgmodal.on{display:flex}
.lgcard{background:#0f2438;border:1px solid #2a4a6b;border-radius:14px;padding:24px;width:320px;max-width:92vw;box-shadow:0 24px 60px -12px rgba(0,0,0,.85);text-align:center}
.lgcard h3{margin:0 0 16px;color:#ffc000;font-size:13px;letter-spacing:1.6px}
.lgcard select,.lgcard input{width:100%;box-sizing:border-box;margin-bottom:10px;padding:10px 12px;border-radius:9px;border:1px solid #2a4a6b;background:#0a1a2c;color:#eaf1f8;font-size:12.5px;outline:none}
.lgcard select:focus,.lgcard input:focus{border-color:#ed7d31}
.lgo{width:100%;padding:11px;border:none;border-radius:9px;background:linear-gradient(135deg,#ff9a4d,#d95b17);color:#141414;font-weight:900;font-size:12px;letter-spacing:1.2px;cursor:pointer;margin-top:2px}
.lgo:hover{filter:brightness(1.1)}
.lgerr{color:#ff8577;font-size:11px;text-align:center;min-height:15px;margin-top:8px;font-weight:700}
@media print{
body::before,#xlsbtn,#prnbtn,#svbtn,#rfbtn,#lgbtn,header,.flow,footer,.tools,.hint,#syncbadge,#clock,#lgmodal{display:none!important}
body{background:#fff!important}
.view{display:none!important}
.view.on{display:block!important}
.panel{background:#fff!important;border:1px solid #b9c2cd!important;box-shadow:none!important;margin:8px 0!important;break-inside:auto}
.kpis{padding-top:8px!important}
.tblwrap{max-height:none!important;overflow:visible!important;border:none!important}
th{position:static!important;background:#e8edf3!important;color:#000!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}
td{color:#000!important;background:#fff!important}
tbody tr:nth-child(even) td{background:#f4f6f9!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}
h1,h2,h3,p,span,b,div,td,th,li,label{color:#000!important;text-shadow:none!important}
}
</style>
<div id="lgmodal"><div class="lgcard"><h3>&#128274; SIGN IN TO EDIT</h3><select id="lguser"></select><input id="lgpass" type="password" placeholder="PASSWORD" autocomplete="off"><button class="lgo" id="lgo">ENTER</button><div class="lgerr" id="lgerr"></div></div></div>
<script>
(function(){"use strict";
var pb=document.getElementById('prnbtn');if(pb)pb.addEventListener('click',function(){try{window.print();}catch(e){}});
var sb=document.getElementById('svbtn');
function post(sh,id,patch){try{fetch('/api/state',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(Object.assign({sheet:sh,id:id},patch))}).catch(function(){});}catch(e){}}
function resave(){var n=0,i,k,st;try{for(i=0;i<localStorage.length;i++){k=localStorage.key(i);if(!k||k.indexOf('xstate_')!==0)continue;var sh=k.slice(7);try{st=JSON.parse(localStorage.getItem(k)||'{}');}catch(e){continue;}for(var id in st){var v=st[id];if(id==='_c'){for(var j in v){post(sh,j,{color:v[j]});n++;}}else if(id==='_e'){for(var j2 in v){var o=v[j2];for(var c in o){post(sh,j2,{col:c,value:o[c]});n++;}}}else{post(sh,id,{note:v});n++;}}}}catch(e){}
var oldt=sb.title;sb.textContent='OK'+n;sb.style.fontSize='9px';sb.style.fontWeight='900';sb.disabled=true;setTimeout(function(){sb.innerHTML='&#128190;';sb.style.fontSize='';sb.style.fontWeight='';sb.disabled=false;sb.title=oldt;},2200);}
if(sb)sb.addEventListener('click',resave);
})();
</script>
<script>
(function(){"use strict";
var USERS={'MOHAMED ABDELGAWAD':'m123@','JABESTEEN':'j123@','ZEESHAN ANWAR':'z123@'};
function get(k){try{return sessionStorage.getItem(k)}catch(e){return null}}
function setk(k,v){try{sessionStorage.setItem(k,v)}catch(e){}}
function delk(k){try{sessionStorage.removeItem(k)}catch(e){}}
function authed(){return !!get('ps5who')}
function applyAll(){var on=authed();var b=document.getElementById('lgbtn');if(b){b.innerHTML=on?'&#128275;':'&#128274;';b.style.background=on?'linear-gradient(135deg,#2fbe6e,#136b3a)':'linear-gradient(135deg,#27405c,#16283c)';b.style.color=on?'#fff':'#ffc000';b.title=on?('EDITING AS '+get('ps5who')+' - CLICK TO SIGN OUT'):'SIGN IN TO EDIT';}
var nodes=document.querySelectorAll('.edt');for(var i=0;i<nodes.length;i++){nodes[i].contentEditable=on?'true':'false';}
var m=document.getElementById('lgmodal');if(m)m.classList.remove('on');var p=document.getElementById('lgpass');if(p)p.value='';var er=document.getElementById('lgerr');if(er)er.textContent='';}
function openModal(){var m=document.getElementById('lgmodal');if(!m)return;m.classList.add('on');var s=document.getElementById('lguser');if(s&&!s.options.length){for(var k in USERS){var o=document.createElement('option');o.value=k;o.textContent=k;s.appendChild(o);}}setTimeout(function(){var p=document.getElementById('lgpass');if(p)p.focus();},80);}
function tryIn(){var s=document.getElementById('lguser'),p=document.getElementById('lgpass'),er=document.getElementById('lgerr');if(!s||!p||!er)return;var u=s.value;if(USERS[u]===p.value){setk('ps5who',u);applyAll();}else{er.textContent='WRONG PASSWORD - TRY AGAIN';p.value='';p.focus();}}
document.addEventListener('beforeinput',function(e){var t=e.target;if(t&&t.classList&&t.classList.contains('edt')&&!authed()){e.preventDefault();openModal();}},true);
document.addEventListener('click',function(e){if(e.target&&e.target.id==='lgmodal'){e.target.classList.remove('on');}},true);
document.addEventListener('keydown',function(e){if(e.key==='Escape'){var m=document.getElementById('lgmodal');if(m)m.classList.remove('on');}});
var b=document.getElementById('lgbtn');if(b)b.addEventListener('click',function(){if(authed()){if(window.confirm('SIGN OUT ('+get('ps5who')+') ?')){delk('ps5who');applyAll();}}else{openModal();}});
var go=document.getElementById('lgo');if(go)go.addEventListener('click',tryIn);
var pp=document.getElementById('lgpass');if(pp)pp.addEventListener('keydown',function(e){if(e.key==='Enter')tryIn();});
applyAll();
})();
</script>'''


def build_html(keys, punch, itr, rep_date, state=None, rfc=None,
               itr_sum=None, pun_sum=None, miles_n=0, miles=None):
    state = state or {'notes': {}, 'pcolors': {}}
    rfc = rfc or []
    subs = [{'name': k['name'], 'sid': k['sid']} for k in keys]
    sub_idx = {k['name']: i for i, k in enumerate(keys)}
    sub_sids = [k['sid'] for k in keys]

    def _pj(p):
        i = sub_idx.get(p['sub'], -1)
        return [p['id'], p['tag'], p['cat'], p['disc'], p['desc'][:120],
                p['status'], p['close'], i, p['sub'],
                sub_sids[i] if i >= 0 else '']

    def _ij(r):
        i = sub_idx.get(r['sub'], -1)
        return [r['id'], r['tag'], r['disc'], r['ttype'], r['adesc'][:100],
                r['state'], r['close'], i, r['sub'],
                sub_sids[i] if i >= 0 else '']

    pj = [_pj(p) for p in punch]
    ij = [_ij(r) for r in itr]
    def _pc(v):
        return '' if v is None else round(v * 100)
    rf = []
    for x in rfc:
        rf.append({
            'sid': x['sid'], 'name': x['name'], 'prio': x['prio'],
            'bhm': x['rfc_bhm'], 'eit': x['rfc_eit'], 'base': x['baseline'],
            'rec': x.get('recovery', ''), 'signed': x.get('signed', ''),
            'mile': x['milestone'], 'tot': _pc(x['tot_pct']),
            'itrp': _pc(x['itr_pct']),
            'b1': int(x['blk_cpp1'] or 0), 'b2': int(x['blk_eit'] or 0),
            'b3': int(x['blk_eacop'] or 0),
            're1': x['rem_eacop'], 're2': x['rem_cpeit'], 're3': x['rem_cpp1'],
            'wd': x['walkdown'], 'mile2': '',
            'd': [{'t': x['disc'][k]['t'], 'c': x['disc'][k]['c']}
                  for k in ('B', 'E', 'H', 'I', 'M', 'P', 'S', 'T')],
            'it': x['itrs'], 'ic': x['closed'], 'ib': x['balance'],
        })
    rf_done = [x for x in rf if x['tot'] != '' and x['tot'] >= 100]

    # dashboard summary data
    itrt, punt = [], []
    it_tot = it_clo = 0
    it_pct = None
    if itr_sum:
        _, _, irows = itr_sum
        for row in irows:
            nm = (row['name'] or '').lower()
            itrt.append([row['group'], row['name'], row['total'], row['closed'],
                         row['pct'], row['open']])
            if nm == 'total':
                it_tot = row['total']; it_clo = row['closed']; it_pct = row['pct']
    if pun_sum:
        order, pdata = pun_sum
        for name in order:
            e = pdata[name]
            punt.append([name] + list(e['A']) + list(e['B']) + list(e['C']))
            if name.lower().startswith('grand'):
                pu_tot = e['A'][0] + e['B'][0] + e['C'][0]
                pu_clo = e['A'][1] + e['B'][1] + e['C'][1]

    # authoritative KPI values computed straight from the live data rows
    pu_tot = len(punch)
    pu_clo = sum(1 for p in punch if (p['status'] or '').strip().lower() == 'closed')
    it_tot = len(itr)
    it_clo = sum(1 for r in itr
                 if 'complete' in (r['state'] or '').lower()
                 or (r['state'] or '').strip().lower() == 'closed')
    it_pct = (it_clo / it_tot) if it_tot else None

    page = """<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PS5 Completion Platform</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600&display=swap" rel="stylesheet">

<style>
#errbox{display:none;position:fixed;left:8px;bottom:8px;z-index:20000;background:#c00000;color:#fff;font-size:10.5px;padding:8px 12px;border-radius:8px;max-width:92vw;white-space:pre-wrap;font-family:Consolas,monospace}
*{margin:0;padding:0;box-sizing:border-box}
:root{--nav:#0b2239;--nav2:#12406e;--acc:#ed7d31;--gold:#ffc000;--line:#e6ebf2}
::selection{background:#ffe1c2}
body{font-family:'IBM Plex Sans','Segoe UI',Tahoma,Arial,sans-serif;background:
 radial-gradient(900px 320px at 90% -80px,rgba(21,90,158,.08),transparent),
 radial-gradient(700px 280px at -60px 30%,rgba(237,125,49,.07),transparent),
 #f3f5f9;
 color:#1a2a3a;font-size:13px;font-variant-numeric:tabular-nums}
header{background:linear-gradient(135deg,#071a30,#0d2f52 55%,#123f6b);color:#fff;padding:6px 18px;display:flex;align-items:center;gap:10px;position:sticky;top:0;z-index:70;min-height:50px;box-shadow:0 2px 10px rgba(4,18,36,.4)}
header::after{content:'';position:absolute;left:0;right:0;bottom:-3px;height:3px;background:linear-gradient(90deg,var(--acc),var(--gold) 40%,transparent 90%)}
.logo{width:32px;height:32px;border-radius:9px;background:linear-gradient(135deg,var(--acc),var(--gold));display:flex;align-items:center;justify-content:center;font-weight:800;font-size:11px;color:#081c30;box-shadow:0 0 0 3px rgba(255,192,0,.15)}
header h1{font-size:15px;font-weight:800;letter-spacing:1.4px}
header h1 span{color:#ffd27a;display:inline-block;background:linear-gradient(90deg,#ffd27a,#ffb055);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent}
header .logo svg{display:block}
header .meta{font-family:'IBM Plex Mono',Consolas,monospace;font-size:10px;color:#9db8d2;margin-top:2px;letter-spacing:.5px}
header .logo{width:34px;height:34px;border-radius:10px;background:linear-gradient(145deg,#16406e,#0b2239);display:flex;align-items:center;justify-content:center;box-shadow:inset 0 0 0 1px rgba(255,192,0,.4),0 8px 20px -6px rgba(0,0,0,.55)}
#syncbadge{margin-left:10px;font-size:10px;padding:4px 12px;border-radius:999px;background:#1f9d55;font-weight:700;white-space:nowrap;box-shadow:inset 0 0 0 1px rgba(255,255,255,.22)}
#clock{margin-left:auto;font-family:'IBM Plex Mono',Consolas,monospace;font-size:11px;font-weight:600;color:#ffd27a;letter-spacing:.8px;background:rgba(255,255,255,.07);border:1px solid rgba(255,210,122,.35);padding:6px 12px;border-radius:8px;white-space:nowrap}
.fbtn{float:right;margin-left:6px;cursor:pointer;font-size:8.5px;background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.38);border-radius:4px;padding:0 4px;line-height:13px;color:#fff}
.fbtn:hover{background:var(--acc);border-color:#fff}
th.act{box-shadow:inset 0 0 0 2px var(--gold)}
th.act .fbtn{background:var(--acc);border-color:#fff}
.fmenu{position:fixed;z-index:9999;width:235px;background:#fff;border:1px solid #c9d4e2;border-radius:10px;box-shadow:0 14px 34px -12px rgba(10,30,60,.45);overflow:hidden;font-size:11px}
.fh{display:flex;gap:6px;padding:7px;border-bottom:1px solid var(--line);background:#f5f8fc}
.fh input{flex:1;padding:4px 8px;font-size:11px;border:1.5px solid #ccd6e4;border-radius:6px;min-width:0}
.mini{cursor:pointer;font-weight:700;font-size:10px;color:#144a7c;background:#e3edf8;border:1px solid #c3d6ea;border-radius:6px;padding:4px 9px;white-space:nowrap}
.mini:hover{background:#d3e4f5}
.fl{max-height:240px;overflow:auto}
.fitem{display:flex;align-items:center;gap:7px;padding:5px 10px;border-bottom:1px solid #f0f3f7;cursor:pointer;margin:0}
.fitem:hover{background:#eef5fd}
.fitem span{flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;text-transform:none}
.fitem i{color:#8a97a5;font-style:normal;font-size:9.5px}
.pnlbtns{margin-left:auto;display:flex;gap:6px}
.pbt{cursor:pointer;font-family:inherit;font-size:9.5px;font-weight:700;letter-spacing:.6px;border:1.5px solid #ccd6e4;background:#fff;color:#44566c;border-radius:7px;padding:4px 10px;transition:.15s;white-space:nowrap}
.pbt:hover{border-color:var(--acc);color:var(--acc);box-shadow:0 2px 8px -3px rgba(237,125,49,.5)}
.pbt:active{transform:scale(.96)}
#loginbtn{margin-left:8px;cursor:pointer;font-family:inherit;font-size:10px;font-weight:800;letter-spacing:.7px;color:#0b2239;background:linear-gradient(135deg,#ffd27a,#ffb055);border:none;border-radius:999px;padding:6px 13px;white-space:nowrap;transition:.15s}
#loginbtn:hover{filter:brightness(1.05)}
#loginbtn.on{background:#1f9d55;color:#fff}
#lmodal{position:fixed;inset:0;background:rgba(4,18,36,.62);backdrop-filter:blur(3px);display:none;align-items:center;justify-content:center;z-index:10000}
#lmodal.on{display:flex}
#lmbox{background:#fff;border-radius:16px;padding:22px 24px;width:min(92vw,340px);box-shadow:0 30px 80px -20px rgba(0,0,0,.55);text-align:center}
#lmbox h3{font-size:14px;color:#0b2239;margin-bottom:12px;letter-spacing:1px}
#lmbox input,#lmbox select{width:100%;padding:9px 12px;font-size:13px;border:1.5px solid #ccd6e4;border-radius:9px;margin-bottom:9px;text-align:center;outline:none;font-family:inherit;background:#fff}
#lmbox input:focus,#lmbox select:focus{border-color:var(--acc);box-shadow:0 0 0 3px rgba(237,125,49,.15)}
.lmrow{display:flex;gap:8px;margin-top:4px}
.lmrow button{flex:1;cursor:pointer;font-family:inherit;font-weight:800;font-size:11px;letter-spacing:.8px;border:none;border-radius:9px;padding:9px 0}
#lok{background:linear-gradient(135deg,#0b2239,#155a9e);color:#fff}
#lcx{background:#eef2f7;color:#44566c}
.lmsg{font-size:10px;color:#8a97a5;margin-top:10px}
.shake{animation:shk .35s}
@keyframes shk{25%{transform:translateX(-7px)}75%{transform:translateX(7px)}}
.toast{position:fixed;left:50%;bottom:26px;transform:translateX(-50%) translateY(14px);background:#0b2239;color:#ffd27a;font-size:11.5px;font-weight:700;letter-spacing:.5px;padding:10px 18px;border-radius:999px;opacity:0;transition:.3s;z-index:10001;box-shadow:0 10px 30px -10px rgba(0,0,0,.5);max-width:90vw;text-align:center}
.toast.on{opacity:1;transform:translateX(-50%)}
body:not(.canedit) .edt{cursor:not-allowed}
@media print{
 header,.flow,footer,.tools,.hint,#syncbadge,#clock,body::before,.pnlbtns{display:none!important}
 body{background:#fff!important}
 .view{display:none!important}
 body:not(.print-one) .view.on{display:block!important}
 body.print-one .view.pv{display:block!important}
 body.print-one .panel{display:none!important}
 body.print-one .panel.pp{display:block!important}
 .kpis{padding-top:10px!important}
 .panel{box-shadow:none!important;border:1px solid #b7c1cd;margin:8px 0;break-inside:avoid}
 .tblwrap{max-height:none!important;overflow:visible!important;border:none}
 th{position:static!important;background:#dde5ee!important;color:#000!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}
 tbody tr:nth-child(even){background:#f4f6f9!important;-webkit-print-color-adjust:exact}
 tr{break-inside:avoid}
 a{text-decoration:none;color:#000}
}
#syncbadge.off{background:#c2610c}
.flow{display:flex;flex-wrap:wrap;gap:3px;background:rgba(255,255,255,.94);backdrop-filter:blur(6px) saturate(1.15);padding:4px 16px;border-bottom:1px solid var(--line);position:sticky;top:50px;z-index:60}
.step{border:none;cursor:pointer;font-family:inherit;font-weight:700;font-size:10.5px;color:#44566c;background:transparent;padding:6px 14px;border-radius:999px;transition:.15s;letter-spacing:.5px;text-transform:uppercase}
.step:hover{background:#e8edf4}
.step.on{background:linear-gradient(135deg,#0b2239,#155a9e);color:#fff;box-shadow:0 3px 9px rgba(11,34,57,.35)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:9px;padding:64px 16px 8px}
.kpi{border-radius:12px;padding:9px 13px 14px;color:#fff;display:flex;flex-direction:column;gap:1px;box-shadow:0 2px 6px rgba(10,30,60,.22);position:relative;overflow:hidden;min-height:66px}
.kpi::before{content:'';position:absolute;top:-30px;right:-20px;width:110px;height:110px;background:radial-gradient(circle,rgba(255,255,255,.28),transparent 65%)}
.kpi b{font-size:20px;line-height:1.05}.kpi span{font-size:10.5px;color:#fff;opacity:1;font-weight:600;letter-spacing:.9px;text-transform:uppercase;margin-top:3px;white-space:normal}
.kpi{padding-bottom:12px}
.kpi .pk{position:absolute;left:0;bottom:0;height:3px;background:linear-gradient(90deg,var(--gold),#fff);border-radius:999px;box-shadow:0 0 8px rgba(255,192,0,.75)}
.k-navy{background:linear-gradient(135deg,#0b2239,#155a9e)}.k-blue{background:linear-gradient(135deg,#3a62ad,#6a93d8)}
.k-green{background:linear-gradient(135deg,#256a29,#57a05b)}.k-orange{background:linear-gradient(135deg,#c2610c,#f0973f)}
.k-teal{background:linear-gradient(135deg,#00564c,#26a69a)}
.panel{background:#fff;border:1px solid var(--line);border-radius:14px;box-shadow:0 1px 2px rgba(16,42,67,.05),0 10px 26px -20px rgba(16,42,67,.3);padding:12px 14px;margin:12px 16px}
.panel h2{font-size:12.5px;color:#0b2239;border-left:3px solid var(--acc);padding-left:9px;margin-bottom:9px;text-transform:uppercase;letter-spacing:.7px;display:flex;align-items:center;gap:8px}
.tools{display:flex;flex-wrap:wrap;gap:7px;align-items:center;margin-bottom:9px}
input[type=search],select{padding:6px 11px;font-size:12px;border:1.5px solid #ccd6e4;border-radius:8px;background:#fbfcfe;outline:none;font-family:inherit;transition:.15s}
input[type=search]:focus,select:focus{border-color:var(--acc);box-shadow:0 0 0 3px rgba(237,125,49,.16)}
input[type=search]{min-width:230px}
select{max-width:400px;font-weight:600}
.chip{border:1.5px solid #ccd6e4;background:#fff;color:#44566c;border-radius:999px;padding:5px 13px;font-size:10.5px;font-weight:700;cursor:pointer;user-select:none;transition:.15s}
.chip:hover{border-color:var(--acc)}
.chip.on{background:var(--nav);color:#fff;border-color:var(--nav)}
.tblwrap{overflow:auto;max-height:calc(100vh - 208px);border-radius:12px;border:1px solid var(--line)}
table{border-collapse:separate;border-spacing:0;width:100%;font-size:11.5px;white-space:nowrap}
th{position:sticky;top:0;background:linear-gradient(180deg,#10365e,#0b2239);color:#fff;padding:7px 10px;text-align:left;font-size:10px;z-index:5;letter-spacing:.7px;text-transform:uppercase;border-bottom:2px solid var(--acc)}
td{padding:5px 10px;border-bottom:1px solid #eef1f6;background:#fff}
tbody tr:nth-child(even) td{background:#f8fafd}
tbody tr:hover td{background:#fff5e3}
.sublink{color:#0a58ca;font-weight:600;cursor:pointer;text-decoration:none;border-bottom:1px dashed #86abd6;transition:.12s}
.sublink:hover{color:var(--acc);border-bottom:1px solid var(--acc)}
.st-Closed,.st-Completed{color:#1b5e20;font-weight:700}
.st-Originated,.st-ToBe{color:#b35900}
.catA{background:#ffc7ce!important;font-weight:700;text-align:center}
.catB{background:#ffeb9c!important;font-weight:700;text-align:center}
.catC{background:#c6efce!important;font-weight:700;text-align:center}
.pct{color:#fff;font-weight:700;text-align:center;border-radius:4px;display:inline-block;padding:1px 8px;font-size:10.5px;box-shadow:0 1px 2px rgba(0,0,0,.18)}
.pb{height:11px;min-width:74px;background:#e8edf3;border-radius:999px;overflow:hidden;display:inline-block;vertical-align:middle}
.pb i{display:block;height:100%;background:linear-gradient(90deg,var(--acc),#57a05b);border-radius:999px}
.dcell{text-align:center;font-weight:600}
.blk{background:#c00000!important;color:#fff!important;font-weight:800;text-align:center}
.remcell{white-space:normal!important;min-width:180px;max-width:320px;font-size:11px;color:#444}
.src{white-space:normal!important;max-width:300px;font-size:11px;background:#fff8e1!important}
.cb{width:18px;height:18px;border-radius:50%;border:2px solid #cbd5e0;cursor:pointer;background:#fff;padding:0;transition:.15s}
.cb:hover{transform:scale(1.18);border-color:var(--acc)}
.ni{width:160px;border:1px dashed #b8c4d0;border-radius:6px;padding:3px 7px;font-size:11.5px;background:transparent;font-family:inherit;transition:.15s}
.ni:focus{border:1px solid var(--acc);background:#fffdf2;outline:none}
.saved{color:#2e7d32!important;font-weight:600}
.hint{font-size:11px;color:#789;padding:6px 2px}
.empty{padding:38px;text-align:center;color:#93a3b4;font-size:13px}
.view{display:none}.view.on{display:block}
.two{display:grid;grid-template-columns:repeat(auto-fit,minmax(430px,1fr));gap:0 12px;padding:0 6px}
.two .panel{margin:12px 0 0}
.hero{background:radial-gradient(900px 260px at 88% -60%,rgba(255,192,0,.22),transparent),linear-gradient(135deg,#081c30,#12406e);border-radius:14px;padding:14px 18px;margin:12px 16px;color:#fff;box-shadow:0 3px 14px rgba(4,18,36,.35);border:1px solid rgba(255,255,255,.08)}
.hero h2{font-size:16px;color:#fff;border:none;padding:0;margin:0 0 7px;letter-spacing:.5px}
.hero #sub-kpis{font-size:11.5px;color:#cfe0f2;letter-spacing:.3px}
.hero #sub-kpis b{color:var(--gold)}
.kpi .bar{height:7px;border-radius:999px;background:rgba(255,255,255,.25);overflow:hidden;margin-top:6px}
.kpi .bar i{display:block;height:100%;background:#fff;border-radius:999px}
.kpi .duo{display:flex;justify-content:space-between;font-size:9px;letter-spacing:.6px;opacity:.95;margin-top:4px;text-transform:uppercase}
.statgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(185px,1fr));gap:10px;margin-top:12px}
.stat{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);border-radius:10px;padding:10px 12px}
.stat h4{font-size:9.5px;letter-spacing:1.1px;text-transform:uppercase;color:#9fc3e8;margin-bottom:6px}
.stat .v{font-size:21px;font-weight:800;line-height:1}
.stat .v small{font-size:10.5px;color:#cfe0f2;font-weight:600}
.bar2{height:7px;border-radius:999px;background:rgba(0,0,0,.35);overflow:hidden;margin-top:7px}
.bar2 i{display:block;height:100%;border-radius:999px;background:linear-gradient(90deg,#57a05b,#8bd48d);transition:width .6s ease}
.bar2.open i{background:linear-gradient(90deg,#ed7d31,#ffb066)}
.segbar{display:flex;height:8px;border-radius:999px;overflow:hidden;background:rgba(0,0,0,.35);margin-top:7px}
.segbar i{display:block;height:100%;transition:width .6s ease}
.stat .v .sep{margin:0 9px;opacity:.35;font-weight:400}
.duo2{display:flex;justify-content:space-between;font-size:10px;color:#cfe0f2;margin-top:6px;letter-spacing:.6px;text-transform:uppercase}
.duo2 b{color:#fff}
@keyframes fadeUp{from{opacity:0;transform:translateY(9px)}to{opacity:1;transform:none}}
.view.on{animation:fadeUp .28s ease}
.sidchip{display:inline-block;background:rgba(255,255,255,.12);border:1px solid rgba(255,214,150,.5);color:#ffd9a0;border-radius:6px;padding:2px 9px;font-size:11.5px;font-weight:800;letter-spacing:1.2px;margin-right:8px;vertical-align:2px}
.panel{position:relative;transition:opacity .5s ease,transform .5s ease}
.panel.pre{opacity:0;transform:translateY(14px)}
.panel.in{opacity:1;transform:none}
.panel::before{content:'';position:absolute;inset:0;border-radius:inherit;padding:1px;background:linear-gradient(135deg,rgba(255,255,255,.85),rgba(226,232,240,.25) 45%,rgba(237,125,49,.30));-webkit-mask:linear-gradient(#fff 0 0) content-box,linear-gradient(#fff 0 0);-webkit-mask-composite:xor;mask-composite:exclude;pointer-events:none;opacity:.75}
header::after{content:'';position:absolute;top:0;left:-60%;width:38%;height:100%;background:linear-gradient(105deg,transparent,rgba(255,255,255,.07),transparent);animation:sheen 9s ease-in-out infinite;pointer-events:none}
@keyframes sheen{0%,55%{left:-60%}95%,100%{left:135%}}
nav .seg button.on{box-shadow:0 0 0 3px rgba(237,125,49,.16),var(--sh)}
button:active,.chip:active{transform:scale(.97)}
body{background-image:radial-gradient(900px 420px at 88% -10%,rgba(237,125,49,.08),transparent 60%),radial-gradient(900px 460px at -10% 112%,rgba(46,91,148,.12),transparent 60%)}
@media print{.panel{opacity:1!important;transform:none!important}}
footer{padding:14px 20px 24px;color:#8b99a9;font-size:10.5px;text-align:center}
::-webkit-scrollbar{width:10px;height:10px}
::-webkit-scrollbar-thumb{background:#c2cfdd;border-radius:999px;border:2px solid #f3f5f9}
::-webkit-scrollbar-thumb:hover{background:var(--acc)}
::-webkit-scrollbar-track{background:transparent}
button:focus-visible,input:focus-visible,select:focus-visible{outline:2px solid var(--acc);outline-offset:1px}
body::before{content:'';position:fixed;inset:0;background-image:linear-gradient(rgba(11,34,57,.04) 1px,transparent 1px),linear-gradient(90deg,rgba(11,34,57,.04) 1px,transparent 1px);background-size:26px 26px;pointer-events:none;z-index:0}
header,#wrap,.flow,footer{position:relative;z-index:1}
.panel table th{font-family:'IBM Plex Mono',Consolas,monospace}
.panel table tbody tr:nth-child(even){background:#fafbfd}
.panel table tbody tr:hover td{background:#eef5fd}
.panel table tbody tr:hover td:first-child{box-shadow:inset 3px 0 0 var(--acc)}
</style></head><body>
<div id="errbox"></div>
<header><div class="logo"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="#ffc000" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2l8.5 5v10L12 22l-8.5-5V7L12 2z"/><path d="M12 7.5v9M7.75 10v4M16.25 10v4"/></svg></div>
 <div><h1>PS5 <span>COMPLETION PLATFORM</span></h1><div class="meta">__META__</div></div>
 <div id="clock"></div>
 <div id="syncbadge">&#9679; SYNCED TO EXCEL</div>
 <button id="loginbtn">&#128274; LOGIN TO EDIT</button>
</header>
<div id="lmodal"><div id="lmbox">
 <h3>&#128274; EDITOR LOGIN</h3>
 <select id="lu">
  <option value="">-- SELECT NAME --</option>
  <option>MOHAMED ABDELGAWAD</option>
  <option>JABESTEEN</option>
  <option>ZEESHAN ANWAR</option>
 </select>
 <input id="lp" type="password" placeholder="PASSWORD">
 <div class="lmrow"><button id="lok">LOGIN</button><button id="lcx">CANCEL</button></div>
 <div class="lmsg" id="lmsg">Authorized engineers only &mdash; 3 accounts.</div>
</div></div>
<div class="flow">
 <button class="step on" data-v="dash">1 &middot; DASHBOARD</button>
 <button class="step" data-v="rfc">2 &middot; RFC PROGRESS</button>
 <button class="step" data-v="blk">3 &middot; BLOCKING / REMARKS</button>
 <button class="step" data-v="pu">4 &middot; PUNCH LIST</button>
 <button class="step" data-v="it">5 &middot; ITR LIST</button>
 <button class="step" data-v="sub">6 &middot; SUBSYSTEM</button>
 <button class="step" data-v="done">7 &middot; COMPLETE &#10003;</button>
</div>

<section class="view on" id="v-dash">
 <div class="kpis">__KPIS__</div>
  <div class="two">
   <div class="panel"><h2>ITR Progress by Discipline</h2><div class="tblwrap"><table id="t-itrt"></table></div></div>
   <div class="panel"><h2>Punch Summary (A / B / C)</h2><div class="tblwrap"><table id="t-punt"></table></div></div>
  </div>
  <div class="panel"><h2>MILESTONE &amp; DIS. &mdash; <span id="m-count"></span></h2>
   <div class="tools"><input type="search" id="q-ms" placeholder="Search subsystem / description / milestone / date..."><span class="hint" id="h-ms" style="margin:0"></span></div>
   <div class="tblwrap" style="max-height:460px"><table id="t-ms"></table></div>
  </div>
</section>

<section class="view" id="v-rfc">
 <div class="panel">
  <h2>RFC Progress &mdash; click a subsystem to open its Punch + ITR</h2>
  <div class="tools"><input type="search" id="q-rfc" placeholder="Search subsystem..."></div>
  <div class="tblwrap"><table id="t-rfc"></table></div><div class="hint" id="h-rfc"></div>
 </div>
</section>

<section class="view" id="v-blk">
 <div class="panel">
  <h2 style="border-color:#c00000">Blocking Points &amp; Remarks &mdash; CPP-1 / EIT / EACOP</h2>
   <div class="tools"><input type="search" id="q-blk" placeholder="Search subsystem..."><span class="chip" id="blk-all" style="display:none">&#8630; ALL SYSTEMS</span></div>
  <div class="tblwrap"><table id="t-blk"></table></div><div class="hint" id="h-blk"></div>
 </div>
</section>

<section class="view" id="v-pu">
 <div class="panel">
  <h2>Detailed Punch List</h2>
  <div class="tools">
   <select id="s-pu"></select>
   <input type="search" id="q-pu" placeholder="Search tag / ID / description...">
   <span class="chip on" data-cat="">ALL CATS</span><span class="chip" data-cat="A">A</span><span class="chip" data-cat="B">B</span><span class="chip" data-cat="C">C</span>
  </div>
  <div class="tblwrap"><table id="t-pu"></table></div><div class="hint" id="h-pu"></div>
 </div>
</section>

<section class="view" id="v-it">
 <div class="panel">
  <h2>Detailed ITR List</h2>
  <div class="tools">
   <select id="s-it"></select>
   <input type="search" id="q-it" placeholder="Search tag / ID / description...">
   <span class="chip on" data-st="">ALL STATES</span><span class="chip" data-st="done">DONE</span><span class="chip" data-st="open">OPEN</span>
  </div>
  <div class="tblwrap"><table id="t-it"></table></div><div class="hint" id="h-it"></div>
 </div>
</section>

<section class="view" id="v-sub">
 <div class="hero">
  <h2 id="sub-name">SELECT A SUBSYSTEM</h2>
  <div id="sub-kpis"></div>
 </div>
 <div class="two">
  <div class="panel"><h2>PUNCH LIST</h2><div class="tblwrap"><table id="t-sp"></table></div><div class="hint" id="h-sp"></div></div>
  <div class="panel"><h2>ITR LIST</h2><div class="tblwrap"><table id="t-si"></table></div><div class="hint" id="h-si"></div></div>
 </div>
</section>

<section class="view" id="v-done">
 <div class="panel">
  <h2 style="border-color:#2e7d32">RFC Complete &mdash; Subsystems at 100%</h2>
  <div class="tblwrap"><table id="t-done"></table></div><div class="hint" id="h-done"></div>
 </div>
</section>
<footer>PS5 &middot; CPP-AGI COMPLETION PLATFORM &nbsp;&middot;&nbsp; BUILD 24-Aug-2026 08:20 &nbsp;&middot;&nbsp; &mdash; CONFIDENTIAL &middot; FOR INTERNAL PROJECT USE ONLY</footer>
<script>
window.onerror=function(m,s,l){var b=document.getElementById('errbox');if(b){b.style.display='block';b.innerHTML+='\u26a0 '+String(m)+' @ln '+l+'<br>'}return false};
const SUBS=__SUBS__,PUNCH=__PUNCH__,ITR=__ITR__,RFC=__RFC__,ITRT=__ITRT__,PUNT=__PUNT__,MILES=__MILES__;
const NOTES=__NOTES__,PCOL=__PCOL__,OVR=__OVR__;
const RFK=['B','E','H','I','M','P','S','T'];
{const s=document.createElement('style');s.textContent='.edt:focus{outline:2px solid #ed7d31;background:#fffbe6}.edt.saved{background:#c6efce !important}.edt:hover{box-shadow:inset 0 0 0 1px #cbd5e0}';document.head.appendChild(s);}
function ovv(sh,id,col,def){const o=OVR[sh];const v=(o&&o[id])?o[id][col]:undefined;return (v===undefined||v===null)?def:v;}
function edt(sh,id,col,val,st,cls){return '<td class=\"edt '+(cls||'')+'\" contenteditable=\"true\" spellcheck=\"false\" data-sh=\"'+sh+'\" data-id=\"'+esc(id)+'\" data-col=\"'+esc(col)+'\" '+(st||'')+'>'+esc(val??'')+'</td>';}
const PALETTE=['','FFFFEB9C','FFC6EFCE','FFFFC7CE','FFBDD7EE'];
const HEX={'':'','#FFFFEB9C':'#ffeb9c','#FFC6EFCE':'#c6efce','#FFFFC7CE':'#ffc7ce','#FFBDD7EE':'#bdd7ee'};
let ONLINE=true,V='dash',F={cat:'',stt:''};
const $=id=>document.getElementById(id);
function sget(k){try{return sessionStorage.getItem(k)}catch(e){return null}}
function sset(k,v){try{sessionStorage.setItem(k,v)}catch(e){}}
function sdel(k){try{sessionStorage.removeItem(k)}catch(e){}}
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const doneI=s=>{s=String(s||'').trim();if(/to\s*be/i.test(s))return false;return /^comp/i.test(s)||/^closed$/i.test(s);};
const pcol=p=>p>=90?'#2e7d32':(p>=50?'#ed7d31':(p>0?'#c00000':'#8a97a5'));
if(!navigator.onLine||location.protocol==='file:'){$('syncbadge').textContent='\u25CF OFFLINE - saved locally';$('syncbadge').classList.add('off');ONLINE=false;}
function save(sh,id,patch){
 if(!CANEDIT){toast('VIEW ONLY - login to save');return;}
 try{
  const key='xstate_'+sh;let st={};try{st=JSON.parse(localStorage.getItem(key)||'{}')}catch(e){}
  if(patch.note!==undefined){st[id]=patch.note;}
  if(patch.color!==undefined){(st._c=st._c||{})[id]=patch.color;}
  if(patch.col!==undefined){((st._e=st._e||{})[id]=st._e[id]||{})[patch.col]=patch.value;}
  localStorage.setItem(key,JSON.stringify(st));
 }catch(e){}
 if(!ONLINE)return;
 fetch('/api/state',{method:'POST',headers:{'Content-Type':'application/json'},
  body:JSON.stringify(Object.assign({sheet:sh,id:id},patch))})
 .then(r=>{if(!r.ok)throw 0;$('syncbadge').classList.remove('off');$('syncbadge').innerHTML='\u25CF SYNCED TO EXCEL';})
 .catch(()=>{$('syncbadge').textContent='\u25CF OFFLINE - saved locally';$('syncbadge').classList.add('off');});
}
function fillSel(id){const s=$(id);s.innerHTML='<option value=\"\">ALL SUB SYSTEMS</option>';
 SUBS.forEach((x,i)=>{const o=document.createElement('option');o.value=x.sid;o.textContent=x.name;s.appendChild(o);});}
fillSel('s-pu');fillSel('s-it');
(function(){const f=document.querySelector('.flow'),h=document.querySelector('header');
 function fit(){if(f&&h)f.style.top=h.getBoundingClientRect().height.toFixed(0)+'px'}
 fit();window.addEventListener('resize',fit);})();
function bar(p){return '<span class=\"pb\"><i style=\"width:'+Math.max(2,p)+'%\"></i></span>';}
function renderDash(){
 let h='<tr><th>GROUP</th><th>DISCIPLINE</th><th>TOTAL</th><th>CLOSED</th><th>%</th><th>OPEN</th></tr>';
 for(const r of ITRT){
  const tot=String(r[1]||'').toLowerCase().indexOf('total')>=0;
  const p=(r[4]===null||r[4]==='')?null:Math.round(r[4]*100);
  h+='<tr'+(tot?' style=\"font-weight:700\"':'')+'><td>'+esc(r[0])+'</td><td>'+esc(r[1])+'</td><td>'+(r[2]??'')+'</td><td>'+(r[3]??'')+'</td>';
  h+=(p===null)?'<td></td>':'<td style=\"text-align:center\"><span class=\"pct\" style=\"background:'+pcol(p)+'\">'+p+'%</span></td>';
  h+='<td>'+(r[5]??'')+'</td></tr>';
 }
 $('t-itrt').innerHTML=h;
 let g='<tr><th rowspan=\"2\" style=\"z-index:6\">DISCIPLINE</th><th colspan=\"3\" style=\"background:#4472c4\">A</th><th colspan=\"3\" style=\"background:#ed7d31\">B</th><th colspan=\"3\" style=\"background:#57a05b\">C</th></tr>'+
 '<tr><th>T</th><th>CL</th><th>OP</th><th>T</th><th>CL</th><th>OP</th><th>T</th><th>CL</th><th>OP</th></tr>';
 for(const r of PUNT){const gr=String(r[0]).toLowerCase().startsWith('grand');
  g+='<tr'+(gr?' style=\"font-weight:700;background:#ffe9c9\"':'')+'><td>'+esc(r[0])+'</td>';
  for(let i=1;i<10;i++)g+='<td style=\"text-align:center\">'+(r[i]??'')+'</td>';
  g+='</tr>';}
 $('t-punt').innerHTML=g;
}
let SUBCUR='';
function goSub(sid){
 SUBCUR=sid;
 $('s-pu').value=sid;$('s-it').value=sid;$('q-pu').value='';$('q-it').value='';
 try{history.replaceState(null,'','#'+sid)}catch(e){}
 show('sub');renderSub();
}
function subRows(src,isP){
 const sh=isP?'PUNCH LIST':'ITR LIST';
 const N=NOTES[sh]||{},C=PCOL[sh]||{};
 let rows='';
 for(const r of src){
  const col=C[r[0]]||'';const bs=col?'style=\"background:'+HEX[col]+'\"':'';
  rows+='<tr data-id=\"'+esc(r[0])+'\" data-sh=\"'+sh+'\">';
  rows+='<td '+bs+'>'+esc(r[0])+'</td>';
  if(isP){
   rows+=edt(sh,r[0],'TAG',ovv(sh,r[0],'TAG',r[1]),bs);
   rows+=edt(sh,r[0],'CAT',ovv(sh,r[0],'CAT',r[2]),'',r[2]?'cat'+esc(r[2]):'');
   rows+=edt(sh,r[0],'DISC',ovv(sh,r[0],'DISC',r[3]),bs);
  }else{
   rows+=edt(sh,r[0],'TAG',ovv(sh,r[0],'TAG',r[1]),bs);
   rows+=edt(sh,r[0],'DISC',ovv(sh,r[0],'DISC',r[2]),bs);
  }
  const scol=isP?'STATUS':'STATE';
  rows+=edt(sh,r[0],'DESCRIPTION',ovv(sh,r[0],'DESCRIPTION',r[4]),'style=\"white-space:normal;min-width:260px\"');
  rows+=edt(sh,r[0],scol,ovv(sh,r[0],scol,r[5]),'','st-'+esc(String(r[5]).replace(/\\s/g,'')));
  rows+=edt(sh,r[0],'CLOSING DATE',ovv(sh,r[0],'CLOSING DATE',r[6]),bs);
  rows+='<td '+(col?'style=\"background:'+HEX[col]+'\"':'')+'><button class=\"cb\" style=\"border-color:'+(col?HEX[col]:'#cbd5e0')+';background:'+(col?HEX[col]:'#fff')+'\"></button></td>';
  rows+='<td><input class=\"ni\" placeholder=\"+ note...\" value=\"'+esc(N[r[0]]||'')+'\"></td></tr>';
 }
 return rows;
}
function renderSub(){
 const x=RFC.find(r=>r.sid===SUBCUR)||{name:SUBCUR};
 $('sub-name').innerHTML='<span class="sidchip">'+esc(SUBCUR||'-')+'</span>'+esc(x.name||'SELECT A SUBSYSTEM');
 const P=PUNCH.filter(r=>r[9]===SUBCUR);
 const I=ITR.filter(r=>r[9]===SUBCUR);
 const pc=P.filter(r=>r[5]==='Closed').length;
 const ic=I.filter(r=>doneI(r[5])).length;
 const pp=P.length?Math.round(pc/P.length*100):0;
 const ip=I.length?Math.round(ic/I.length*100):0;
 function st(lab,cp,op,c,o,n){
  return '<div class="stat"><h4>'+lab+'</h4><div class="v">'+cp+'% <small>CLOSED</small><span class="sep">|</span>'+op+'% <small>OPEN</small></div><div class="segbar"><i style="width:'+cp+'%;background:linear-gradient(90deg,#57a05b,#8bd48d)"></i><i style="width:'+op+'%;background:linear-gradient(90deg,#ed7d31,#ffb066)"></i></div><div class="duo2"><span><b>'+c.toLocaleString('en-US')+'</b> CLOSED</span><span><b>'+o.toLocaleString('en-US')+'</b> OPEN · '+n.toLocaleString('en-US')+' TOT</span></div></div>';
 }
 $('sub-kpis').innerHTML='<div class="statgrid" style="grid-template-columns:repeat(auto-fit,minmax(250px,1fr))">'
  +st('PUNCH',pp,100-pp,pc,P.length-pc,P.length)
  +st('ITR',ip,100-ip,ic,I.length-ic,I.length)
  +'</div><div style="margin-top:10px">TOTAL %: <b>'+(x.tot===''?'-':x.tot+'%')+'</b> &nbsp;|&nbsp; RFC SIGNED: <b>'+esc(x.signed||'-')+'</b></div>';
 const ph='<tr><th>ID</th><th>TAG</th><th>CAT</th><th>DISC</th><th>DESCRIPTION</th><th>STATUS</th><th>CLOSE</th><th style=\"width:26px\">&#9873;</th><th style=\"min-width:150px\">NOTES</th></tr>';
 const ih='<tr><th>ID</th><th>TAG</th><th>DISC</th><th>DESCRIPTION</th><th>STATE</th><th>CLOSE</th><th style=\"width:26px\">&#9873;</th><th style=\"min-width:150px\">NOTES</th></tr>';
 $('t-sp').innerHTML='<thead>'+ph+'</thead><tbody>'+(subRows(P,true)||'<tr><td colspan=\"9\"><div class=\"empty\">No punch items</div></td></tr>')+'</tbody>';
 $('t-si').innerHTML='<thead>'+ih+'</thead><tbody>'+(subRows(I,false)||'<tr><td colspan=\"8\"><div class=\"empty\">No ITR tasks</div></td></tr>')+'</tbody>';
 $('h-sp').textContent=P.length+' punch item(s)';
 $('h-si').textContent=I.length+' itr task(s) - all cells editable &amp; synced to Excel';
}
function renderRfc(){
 const SH='RFC PROGRESS';
 const q=$('q-rfc').value.trim().toLowerCase();
 let d=RFC.filter(x=>!q||x.name.toLowerCase().includes(q)||x.sid.toLowerCase().includes(q));
 let hh='<tr><th>SUB SYSTEM</th><th>PRIORITY</th><th>RFC DATE (B,H,M,P,S)</th><th>RFC DATES EIT</th><th>BASELINE FORCAST</th><th>RECOVERY/REVISED</th><th>RFC SIGNED</th><th>MILESTONE</th><th>TOTAL %</th>';
 for(const dd of RFK)hh+='<th title=\"'+dd+'\">'+dd+'</th>';
 hh+='<th>ITRs</th><th>CLOSED</th><th>BALANCE</th><th>ITR %</th><th>CPP-1</th><th>EIT</th><th>EACOP</th><th>EACOP</th><th>CPP-EIT</th><th>CPP-1</th><th>WALKDOWN STATUS</th><th style=\"width:26px\">&#9873;</th><th style=\"min-width:170px\">MY REMARKS</th></tr>';
 let rows='';
 for(const x of d){
  rows+='<tr data-id=\"'+esc(x.sid)+'\" data-sh=\"'+SH+'\">';
  rows+='<td class=\"sublink\" data-go=\"'+esc(x.sid)+'\" style=\"white-space:normal;min-width:240px\">'+esc(x.name)+'</td>';
  rows+=edt(SH,x.sid,'Priority',x.prio,'style=\"text-align:center;font-weight:700\"');
  rows+=edt(SH,x.sid,'RFC BHMPS',x.bhm,'style=\"text-align:center\"')+edt(SH,x.sid,'RFC EIT',x.eit,'style=\"text-align:center\"')+edt(SH,x.sid,'Baseline',x.base,'style=\"text-align:center\"')+edt(SH,x.sid,'Recovery',x.rec,'style=\"text-align:center\"');
  const scol=x.signed==='COMPLETE'?'color:#2e7d32':(x.signed==='PARTIAL'?'color:#ed7d31':'');
  rows+=edt(SH,x.sid,'SIGNED',x.signed||'-','style=\"text-align:center;font-weight:700;'+scol+'\"');
  rows+=edt(SH,x.sid,'Milestone',x.mile);
  const tp=x.tot===''?null:x.tot;
  rows+='<td><div style=\"display:flex;align-items:center;gap:6px\">'+(tp===null?'':bar(tp))+'<b style=\"font-size:11px\">'+(tp===null?'-':tp+'%')+'</b></div></td>';
  for(let k=0;k<8;k++){
   const dd=x.d[k];
   rows+=edt(SH,x.sid,RFK[k]+' TOTAL',dd.t,'style=\"text-align:center\"')+edt(SH,x.sid,RFK[k]+' CLOSED',dd.c,'style=\"text-align:center\"');
   const p=(dd.t!==''&&dd.c!=='')?Math.round(dd.c/dd.t*100):null;
   rows+='<td class=\"dcell\" style=\"text-align:center;color:'+(p===null?'inherit':pcol(p))+'\">'+(p===null?'-':p+'%')+'</td>';
  }
  rows+=edt(SH,x.sid,'ITRs',x.it,'style=\"text-align:center;font-weight:700\"')+edt(SH,x.sid,'CLOSED',x.ic,'style=\"text-align:center\"')+edt(SH,x.sid,'BALANCE',x.ib,'style=\"text-align:center\"');
  rows+='<td style=\"text-align:center\"><span class=\"pct\" style=\"background:'+pcol(x.itrp||0)+'\">'+(x.itrp===''?'-':x.itrp+'%')+'</span></td>';
  rows+=edt(SH,x.sid,'CPP-1',x.b1||'','style=\"text-align:center\"',x.b1>0?'blk':'')+edt(SH,x.sid,'EIT',x.b2||'','style=\"text-align:center\"',x.b2>0?'blk':'')+edt(SH,x.sid,'EACOP',x.b3||'','style=\"text-align:center\"',x.b3>0?'blk':'');
  rows+=edt(SH,x.sid,'REMARK EACOP',x.re1,'style=\"min-width:130px;background:#fff8e1\"')+edt(SH,x.sid,'REMARK CPP-EIT',x.re2,'style=\"min-width:130px;background:#fff8e1\"')+edt(SH,x.sid,'REMARK CPP-1',x.re3,'style=\"min-width:130px;background:#fff8e1\"');
  rows+=edt(SH,x.sid,'STATUS',x.wd,'style=\"text-align:center\"');
  rows+='<td style=\"width:26px\"><button class=\"cb\"></button></td>';
  rows+='<td><input class=\"ni\" style=\"width:200px\" placeholder=\"+ remark...\" value=\"'+esc((NOTES[SH]||{})[x.sid]||'')+'\"></td></tr>';
 }
 $('t-rfc').innerHTML='<thead>'+hh+'</thead><tbody>'+(rows||emptyRow())+'</tbody>';
 $('h-rfc').textContent=d.length+' subsystem(s) - every cell is editable & syncs to Excel';
}
function emptyRow(){return '<tr><td colspan=\"20\"><div class=\"empty\">No matching records</div></td></tr>';}
function renderBlk(){
  const SH='RFC PROGRESS';
  const q=$('q-blk').value.trim().toLowerCase();
  let d=RFC.filter(x=>!q||x.name.toLowerCase().includes(q)||x.sid.toLowerCase().includes(q));
  if(SUBCUR)d=d.filter(x=>x.sid===SUBCUR);
  d=d.slice().sort((a,b)=>(b.b1+b.b2+b.b3)-(a.b1+a.b2+a.b3));
 const N=NOTES[SH]||{},C=PCOL[SH]||{};
 let h='<tr><th>SUB SYSTEM</th><th>PRIO</th><th>TOTAL %</th><th>CPP-1</th><th>EIT</th><th>EACOP</th><th>EACOP</th><th>CPP-EIT</th><th>CPP-1</th><th>WALKDOWN</th><th>&#9873;</th><th style=\"min-width:170px\">MY REMARKS</th></tr>';
 let rows='';
 for(const x of d){
  const col=C[x.sid]||'';const bg=col?'style=\"background:'+HEX[col]+'\"':'';
  rows+='<tr data-id=\"'+esc(x.sid)+'\" data-sh=\"'+SH+'\">';
  rows+='<td '+bg+' class=\"sublink\" data-go=\"'+esc(x.sid)+'\" style=\"white-space:normal;min-width:240px\">'+esc(x.name)+'</td>';
  rows+='<td '+bg+' style=\"text-align:center\">'+esc(x.prio)+'</td><td '+bg+' style=\"text-align:center\">'+(x.tot===''?'-':x.tot+'%')+'</td>';
  rows+=edt(SH,x.sid,'CPP-1',x.b1||'','style=\"text-align:center\"',x.b1>0?'blk':'')+edt(SH,x.sid,'EIT',x.b2||'','style=\"text-align:center\"',x.b2>0?'blk':'')+edt(SH,x.sid,'EACOP',x.b3||'','style=\"text-align:center\"',x.b3>0?'blk':'');
   rows+=edt(SH,x.sid,'REMARK EACOP',x.re1,'style=\"background:#fff8e1\"','remcell');
   rows+=edt(SH,x.sid,'REMARK CPP-EIT',x.re2,'style=\"background:#fff8e1\"','remcell');
   rows+=edt(SH,x.sid,'REMARK CPP-1',x.re3,'style=\"background:#fff8e1\"','remcell');
  const wst=col?'style=\"background:'+HEX[col]+';text-align:center\"':'style=\"text-align:center\"';
  rows+=edt(SH,x.sid,'STATUS',ovv(SH,x.sid,'STATUS',x.wd),wst);
  rows+='<td '+(col?'style=\"background:'+HEX[col]+'\"':'')+'><button class=\"cb\" style=\"border-color:'+(col?HEX[col]:'#cbd5e0')+';background:'+(col?HEX[col]:'#fff')+'\"></button></td>';
  rows+='<td><input class=\"ni\" style=\"width:200px\" placeholder=\"+ remark...\" value=\"'+esc(N[x.sid]||'')+'\"></td></tr>';
 }
  $('t-blk').innerHTML='<thead>'+h+'</thead><tbody>'+(rows||emptyRow())+'</tbody>';
  $('blk-all').style.display=SUBCUR?'':'none';
  $('h-blk').textContent=d.length+' subsystem(s)'+(SUBCUR?' | SHOWING ONLY: '+SUBCUR:'')+' - sorted by blocking count. Every cell editable.';
}
function renderList(w){
 const isP=(w==='pu');
 const src=isP?PUNCH:ITR;
 const sid=$(isP?'s-pu':'s-it').value;
 const q=$(isP?'q-pu':'q-it').value.trim().toLowerCase();
 let d=src.filter(r=>{
  if(sid&&r[9]!==sid)return false;
  if(isP&&F.cat&&r[2]!==F.cat)return false;
  if(!isP&&F.stt==='done'&&!doneI(r[5]))return false;
  if(!isP&&F.stt==='open'&&doneI(r[5]))return false;
  if(q&&!r.slice(0,5).concat([r[8]]).some(v=>String(v).toLowerCase().includes(q)))return false;
  return true;
 });
 const LIMIT=2000,trunc=d.length>LIMIT;const view=trunc?d.slice(0,LIMIT):d;
 const sh=isP?'PUNCH LIST':'ITR LIST';
 const N=NOTES[sh]||{},C=PCOL[sh]||{};
 const heads=isP?['ID','TAG','CAT','DISC','DESCRIPTION','STATUS','CLOSE','SUB SYSTEM']:['ID','TAG','DISC','TASK TYPE','DESCRIPTION','STATE','CLOSE','SUB SYSTEM'];
 let h='<tr>'+heads.map(c=>'<th>'+c+'</th>').join('')+'<th style=\"width:26px\">&#9873;</th><th style=\"min-width:170px\">NOTES</th></tr>';
 let rows='';
 for(const r of view){
  const col=C[r[0]]||'';const bs=col?'style=\"background:'+HEX[col]+'\"':'';
  rows+='<tr data-id=\"'+esc(r[0])+'\" data-sh=\"'+sh+'\">';
  rows+='<td '+bs+'>'+esc(r[0])+'</td>';
  if(isP){
   rows+=edt(sh,r[0],'TAG',ovv(sh,r[0],'TAG',r[1]),bs);
   rows+=edt(sh,r[0],'CAT',ovv(sh,r[0],'CAT',r[2]),'',r[2]?'cat'+esc(r[2]):'');
   rows+=edt(sh,r[0],'DISC',ovv(sh,r[0],'DISC',r[3]),bs);
  }else{
   rows+=edt(sh,r[0],'TAG',ovv(sh,r[0],'TAG',r[1]),bs);
   rows+=edt(sh,r[0],'DISC',ovv(sh,r[0],'DISC',r[2]),bs);
   rows+=edt(sh,r[0],'TASK TYPE',ovv(sh,r[0],'TASK TYPE',''),bs);
  }
  const dcol=isP?'DESCRIPTION':'ASSET DESCRIPTION',scol=isP?'STATUS':'STATE';
  rows+=edt(sh,r[0],dcol,ovv(sh,r[0],dcol,r[4]),'style=\"white-space:normal;min-width:260px\"');
  rows+=edt(sh,r[0],scol,ovv(sh,r[0],scol,r[5]),'','st-'+esc(String(r[5]).replace(/\\s/g,'')));
  rows+=edt(sh,r[0],'CLOSING DATE',ovv(sh,r[0],'CLOSING DATE',r[6]),bs);
  rows+='<td class=\"sublink\" data-go=\"'+esc(r[9])+'\" '+bs+' style=\"white-space:normal;min-width:160px\">'+esc(r[8])+'</td>';
  rows+='<td '+(col?'style=\"background:'+HEX[col]+'\"':'')+'><button class=\"cb\" style=\"border-color:'+(col?HEX[col]:'#cbd5e0')+';background:'+(col?HEX[col]:'#fff')+'\"></button></td>';
  rows+='<td><input class=\"ni\" placeholder=\"+ note...\" value=\"'+esc(N[r[0]]||'')+'\"></td></tr>';
 }
 $(isP?'t-pu':'t-it').innerHTML='<thead>'+h+'</thead><tbody>'+(rows||emptyRow())+'</tbody>';
 $(isP?'h-pu':'h-it').textContent=d.length+' record(s)'+(trunc?' (first '+LIMIT+' shown)':'')+(sid?' | subsystem: '+sid:'')+(q?' | search: "'+q+'"':'');
}
function renderDone(){
 const SH='RFC PROGRESS';
 const d=RFC.filter(x=>(x.tot!==''&&x.tot>=100)||x.signed==='COMPLETE'||x.signed==='PARTIAL');
 const N=NOTES[SH]||{};
 let h='<tr><th>#</th><th>SUB SYSTEM</th><th>TOTAL %</th><th>MILESTONE</th><th>RFC SIGNED</th><th>WALKDOWN STATUS</th><th>MY REMARKS</th></tr>';
 let rows='';
 d.forEach((x,i)=>{
  rows+='<tr data-id=\"'+esc(x.sid)+'\" data-sh=\"'+SH+'\">';
  rows+='<td style=\"text-align:center;font-weight:700;background:'+(x.tot!==''&&x.tot>=100?'#c6efce':'#ffeb9c')+'\">'+(i+1)+'</td>';
  rows+='<td class=\"sublink\" data-go=\"'+esc(x.sid)+'\" style=\"white-space:normal;min-width:260px\">'+esc(x.name)+'</td>';
  rows+='<td style=\"text-align:center;font-weight:700\">'+(x.tot===''?'-':x.tot+'%')+'</td>';
  rows+=edt(SH,x.sid,'Milestone',x.mile);
  const scol=x.signed==='COMPLETE'?'color:#2e7d32':(x.signed==='PARTIAL'?'color:#ed7d31':'');
  rows+=edt(SH,x.sid,'SIGNED',x.signed||'-','style=\"text-align:center;font-weight:700;'+scol+'\"');
  rows+=edt(SH,x.sid,'STATUS',ovv(SH,x.sid,'STATUS',x.wd),'style=\"text-align:center\"');
  rows+='<td><input class=\"ni\" style=\"width:220px\" placeholder=\"+ remark...\" value=\"'+esc(N[x.sid]||'')+'\"></td></tr>';
 });
 $('t-done').innerHTML='<thead>'+h+'</thead><tbody>'+(rows||emptyRow())+'</tbody>';
 $('h-done').textContent=d.length+' of '+RFC.length+' subsystems at 100%';
}
const RENDER={dash:null,rfc:renderRfc,blk:renderBlk,pu:()=>renderList('pu'),it:()=>renderList('it'),sub:renderSub,done:renderDone};
function countUp(){
 document.querySelectorAll('#v-dash .kpi b').forEach(el=>{
  const t=el.textContent;
  if(!/^[0-9,]+$/.test(t))return;
  const n=parseInt(t.replace(/,/g,''),10);
  if(isNaN(n))return;
  const st=performance.now();
  function f(now){
   let p=Math.min(1,(now-st)/700);
   p=1-Math.pow(1-p,3);
   el.textContent=Math.round(n*p).toLocaleString('en-US');
   if(p<1)requestAnimationFrame(f);else el.textContent=t;
  }
  requestAnimationFrame(f);
 });
}
const REV_IO=('IntersectionObserver' in window)?new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.remove('pre');e.target.classList.add('in');REV_IO.unobserve(e.target)}})},{threshold:.05}):null;
function initReveal(){
 if(!REV_IO){document.querySelectorAll('.panel').forEach(function(p){p.classList.add('in')});return}
 document.querySelectorAll('.panel:not(.in)').forEach(function(p){
  if(p.classList.contains('pre'))return;
  p.classList.add('pre');
  requestAnimationFrame(function(){requestAnimationFrame(function(){REV_IO.observe(p)})});
 });
}
function show(v){
 V=v;
 document.querySelectorAll('.view').forEach(s=>s.classList.toggle('on',s.id==='v-'+v));
 document.querySelectorAll('.step').forEach(b=>b.classList.toggle('on',b.dataset.v===v));
 if(RENDER[v])RENDER[v]();
 window.scrollTo(0,0);
}
document.querySelectorAll('.step').forEach(b=>b.onclick=()=>show(b.dataset.v));
[['q-rfc',renderRfc],['q-blk',renderBlk],['q-pu',()=>renderList('pu')],['q-it',()=>renderList('it')]].forEach(([id,fn])=>$(id).addEventListener('input',fn));
$('s-pu').addEventListener('change',()=>renderList('pu'));
$('s-it').addEventListener('change',()=>renderList('it'));
$('blk-all').onclick=()=>{SUBCUR='';$('s-pu').value='';$('s-it').value='';renderBlk();};
document.querySelectorAll('.chip[data-cat]').forEach(ch=>ch.onclick=()=>{
 document.querySelectorAll('.chip[data-cat]').forEach(c=>c.classList.remove('on'));
 ch.classList.add('on');F.cat=ch.dataset.cat;renderList('pu');});
document.querySelectorAll('.chip[data-st]').forEach(ch=>ch.onclick=()=>{
 document.querySelectorAll('.chip[data-st]').forEach(c=>c.classList.remove('on'));
 ch.classList.add('on');F.stt=ch.dataset.st;renderList('it');});
document.addEventListener('click',e=>{
 const go=e.target.closest('[data-go]');
 if(go){goSub(go.dataset.go);return;}
  if(!e.target.classList.contains('cb'))return;
  if(!CANEDIT){toast('VIEW ONLY - login to color rows');return;}
 const tr=e.target.closest('tr'),id=tr.dataset.id,sh=tr.dataset.sh;
 const C=PCOL[sh]||(PCOL[sh]={});
 C[id]=PALETTE[(PALETTE.indexOf(C[id]||'')+1)%PALETTE.length];
 save(sh,id,{color:C[id]});
 if(RENDER[V])RENDER[V]();
});
document.addEventListener('focusout',e=>{
 if(!e.target.classList.contains('edt'))return;
 const id=e.target.dataset.id,sh=e.target.dataset.sh,col=e.target.dataset.col;
 const v=e.target.textContent.trim();
 ((OVR[sh]=OVR[sh]||{})[id]=OVR[sh][id]||{})[col]=v;
 e.target.classList.add('saved');setTimeout(()=>e.target.classList.remove('saved'),1200);
 save(sh,id,{col:col,value:v});
});
document.addEventListener('change',e=>{
 if(!e.target.classList.contains('ni'))return;
 const tr=e.target.closest('tr'),id=tr.dataset.id,sh=tr.dataset.sh;
 const N=NOTES[sh]||(NOTES[sh]={});
 N[id]=e.target.value.trim();
 e.target.classList.add('saved');setTimeout(()=>e.target.classList.remove('saved'),1200);
 save(sh,id,{note:N[id]});
});
function applyHash(){
 const hv=decodeURIComponent(location.hash.slice(1)).trim();
 if(!hv)return;
 if(hv.startsWith('view=')){show(hv.slice(5));return;}
 const s=SUBS.find(x=>x.sid===hv||x.name===hv||x.name.startsWith(hv));
 if(s)goSub(s.sid);
}
window.addEventListener('hashchange',applyHash);
renderDash();
function renderMs(){
 const q=($('q-ms').value||'').trim().toLowerCase();
 const d=MILES.filter(x=>!q||[x.sub,x.desc,x.milestone,x.month,x.mantrac,x.priority,x.pg,x.rfsu,x.mp_date,x.eit_date].some(v=>String(v??'').toLowerCase().includes(q)));
 let hh='<tr><th>SUB</th><th>DESCRIPTION</th><th>MILESTONE</th><th>MONTH</th><th>MANTRAC</th><th>PRIORITY</th><th>PG</th><th>RFSU</th><th>MP DATE</th><th>EIT DATE</th></tr>';
 let rows='';
 for(const x of d){
  rows+='<tr><td>'+esc(x.sub)+'</td><td style="white-space:normal;min-width:220px">'+esc(x.desc)+'</td><td>'+esc(x.milestone)+'</td><td>'+esc(x.month)+'</td><td>'+esc(x.mantrac)+'</td><td>'+esc(x.priority)+'</td><td>'+esc(x.pg)+'</td><td>'+esc(x.rfsu)+'</td><td>'+esc(x.mp_date)+'</td><td>'+esc(x.eit_date)+'</td></tr>';
 }
 $('t-ms').innerHTML='<thead>'+hh+'</thead><tbody>'+(rows||'<tr><td colspan="10"><div class="empty">No milestones</div></td></tr>')+'</tbody>';
 $('h-ms').textContent=d.length+' / '+MILES.length+' ROWS';
 $('m-count').textContent=MILES.length+' ROWS';
}
function tick(){
 const d=new Date(),p=n=>String(n).padStart(2,'0');
 const days=['SUN','MON','TUE','WED','THU','FRI','SAT'],mo=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
 $('clock').textContent=days[d.getDay()]+' '+p(d.getDate())+' '+mo[d.getMonth()]+' '+d.getFullYear()+' \u2013 '+p(d.getHours())+':'+p(d.getMinutes())+':'+p(d.getSeconds());
}
const FILTERS={};let FMENU=null;
function closeMenu(){if(FMENU){FMENU.remove();FMENU=null}}
document.addEventListener('click',function(e){if(FMENU&&!FMENU.contains(e.target)&&!e.target.closest('.fbtn'))closeMenu()});
function colVals(tb,ci){const s=new Map();tb.querySelectorAll('tbody tr').forEach(tr=>{const td=tr.children[ci];if(!td)return;const v=(td.textContent||'').trim()||'(Blank)';s.set(v,(s.get(v)||0)+1)});return [...s.entries()].sort((a,b)=>a[0].localeCompare(b[0],'en',{numeric:true}))}
function applyFilters(tid){const tb=document.getElementById(tid);if(!tb)return;const f=FILTERS[tid]||{};tb.querySelectorAll('tbody tr').forEach(tr=>{let ok=true;for(const ci in f){const set=f[ci];if(!set||!set.size)continue;const td=tr.children[ci];const v=((td&&td.textContent)||'').trim()||'(Blank)';if(!set.has(v)){ok=false;break}}tr.style.display=ok?'':'none'})}
function syncTh(tb){const f=FILTERS[tb.id]||{};tb.querySelectorAll('thead th').forEach((th,i)=>th.classList.toggle('act',!!(f[i]&&f[i].size)))}
function openMenu(th,tid,ci){closeMenu();const tb=document.getElementById(tid);if(!tb)return;
 const cur=new Set(((FILTERS[tid]||{})[ci]||[]));
 const m=document.createElement('div');m.className='fmenu';
 m.innerHTML='<div class="fh"><input type="search" placeholder="Search values..."><span class="mini" data-a="all">(All)</span><span class="mini" data-a="ok">OK</span></div><div class="fl"></div>';
 document.body.appendChild(m);
 const r=th.getBoundingClientRect();
 m.style.left=Math.max(6,Math.min(r.left,innerWidth-245))+'px';m.style.top=Math.min(r.bottom+4,innerHeight-300)+'px';
 const fl=m.querySelector('.fl');
 function draw(q){q=(q||'').toLowerCase();fl.innerHTML='';colVals(tb,ci).filter(v=>v[0].toLowerCase().includes(q)).forEach(v=>{const l=document.createElement('label');l.className='fitem';
  l.innerHTML='<input type="checkbox"'+(cur.has(v[0])?' checked':'')+'><span>'+esc(v[0])+'</span><i>'+v[1]+'</i>';
  l.querySelector('input').addEventListener('change',ev=>{ev.target.checked?cur.add(v[0]):cur.delete(v[0])});fl.appendChild(l)});
  if(!fl.children.length)fl.innerHTML='<div style="padding:10px;text-align:center;color:#8a97a5">No match</div>'}
 m.querySelector('.fh input').addEventListener('input',e=>draw(e.target.value));
 m.querySelector('[data-a=all]').addEventListener('click',()=>{if(FILTERS[tid])delete FILTERS[tid][ci];applyFilters(tid);syncTh(tb);closeMenu()});
 m.querySelector('[data-a=ok]').addEventListener('click',()=>{FILTERS[tid]=FILTERS[tid]||{};FILTERS[tid][ci]=cur;applyFilters(tid);syncTh(tb);closeMenu()});
 draw('');FMENU=m;m.querySelector('input').focus();}
try{document.querySelectorAll('table[id]').forEach(tb=>{
 function mk(){tb.querySelectorAll('thead th').forEach((th,i)=>{if(th.querySelector('.fbtn'))return;
  const b=document.createElement('span');b.className='fbtn';b.textContent='\u25BE';
  b.addEventListener('click',ev=>{ev.stopPropagation();openMenu(th,tb.id,i)});th.appendChild(b)})}
 mk();syncTh(tb);
 new MutationObserver(()=>{mk();applyFilters(tb.id);syncTh(tb)}).observe(tb,{childList:true});
});}catch(e){}
function toast(m){const t=document.createElement('div');t.className='toast';t.textContent=m;document.body.appendChild(t);setTimeout(()=>t.classList.add('on'),10);setTimeout(()=>{t.classList.remove('on');setTimeout(()=>t.remove(),350)},2600)}
CANEDIT=true;document.body.classList.add('canedit');
function tblName(h2){return ((h2.childNodes[0]&&h2.childNodes[0].textContent)||'SHEET').replace(/[^A-Za-z0-9 ]+/g,' ').trim().slice(0,28)||'SHEET'}
function exportPanel(pn,h2){
 const tables=[...pn.querySelectorAll('table[id]')];
 if(!tables.length){toast('Nothing to export here');return}
 const stamp=new Date().toISOString().slice(0,10);
  const fname=tblName(h2).replace(/ +/g,'_')+'_'+stamp;
 if(!window.XLSX){
  let csv='';
  tables.forEach(tb=>{tb.querySelectorAll('tr').forEach(tr=>{if(tr.parentElement.tagName==='TBODY'&&tr.style.display==='none')return;
   csv+=[...tr.children].map(td=>'"'+td.innerText.replace(/\u25BE/g,'').trim().replace(/"/g,'""')+'"').join(',')+'\r\n'});csv+='\r\n'});
  const b=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'});const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=fname+'.csv';a.click();
  toast('Saved as CSV (offline mode)');return}
 const wb=window.XLSX.utils.book_new();
 tables.forEach((tb,i)=>{
  const rows=[];
  tb.querySelectorAll('tr').forEach(tr=>{if(tr.parentElement.tagName==='TBODY'&&tr.style.display==='none')return;
   rows.push([...tr.children].map(td=>td.innerText.replace(/\u25BE/g,'').replace(/\s+/g,' ').trim()))});
  const ws=window.XLSX.utils.aoa_to_sheet(rows);
  window.XLSX.utils.book_append_sheet(wb,ws,(tables.length>1?'T'+(i+1):tblName(h2)).slice(0,31));
 });
 window.XLSX.writeFile(wb,fname+'.xlsx');toast('\u2705 Downloaded: '+fname+'.xlsx')}
function printPanel(pn){const v=pn.closest('.view');
 document.body.classList.add('print-one');pn.classList.add('pp');if(v)v.classList.add('pv');
 setTimeout(()=>{window.print();setTimeout(()=>{document.body.classList.remove('print-one');pn.classList.remove('pp');if(v)v.classList.remove('pv')},500)},60)}
try{document.querySelectorAll('.panel h2').forEach(h2=>{if(h2.querySelector('.pnlbtns'))return;
 const bar=document.createElement('span');bar.className='pnlbtns';
 const bx=document.createElement('button');bx.className='pbt';bx.innerHTML='&#128190; SAVE EXCEL';
 const bp=document.createElement('button');bp.className='pbt';bp.innerHTML='&#128424; PRINT';
 bx.addEventListener('click',()=>exportPanel(h2.closest('.panel'),h2));
 bp.addEventListener('click',()=>printPanel(h2.closest('.panel')));
 bar.append(bx,bp);h2.appendChild(bar)});}catch(e){}
try{renderMs()}catch(e){}
try{$('q-ms').addEventListener('input',renderMs)}catch(e){}
try{tick()}catch(e){}try{setInterval(tick,1000)}catch(e){}
setTimeout(function(){document.querySelectorAll('.panel.pre').forEach(function(p){p.classList.add('in')})},1500);
try{applyHash()||show('dash')}catch(e){show('dash')}
try{initReveal()}catch(e){document.querySelectorAll('.panel').forEach(function(p){p.classList.add('in')})}
try{countUp()}catch(e){}
</script></body></html>"""
    meta = f'Report Date: {rep_date}    |    Source: {html_mod.escape(os.path.basename(src_path_global))}'
    notes_js = state.get('notes', {})
    colors_js = {}
    for sh in ('PUNCH LIST', 'ITR LIST', 'RFC PROGRESS'):
        colors_js[sh] = state.get('pcolors', {}).get(sh, {})

    def _kpi(cls, val, lab, frac=None):
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            val = f'{val:,}'
        bar = '' if frac is None else (
            '<i class="pk" style="width:'
            + f'{max(0.0, min(100.0, float(frac))):.0f}%"></i>')
        return (f'<div class="kpi {cls}"><b>{val}</b><span>{lab}</span>'
                + bar + '</div>')
    it_pct_txt = '' if it_pct in (None, '') else str(round(it_pct * 100)) + '%'
    kpis = ''.join([
        _kpi('k-navy', it_tot, 'ITR TOTAL'),
        _kpi('k-green', it_clo, 'ITR CLOSED',
             (it_clo / it_tot * 100) if it_tot else None),
        _kpi('k-teal', it_pct_txt, 'ITR %'),
        _kpi('k-blue', pu_tot, 'PUNCH TOTAL'),
        _kpi('k-green', pu_clo, 'PUNCH CLOSED',
             (pu_clo / pu_tot * 100) if pu_tot else None),
        _kpi('k-orange', max(0, pu_tot - pu_clo), 'PUNCH OPEN',
             ((pu_tot - pu_clo) / pu_tot * 100) if pu_tot else None),
        _kpi('k-navy', miles_n, 'MILESTONES'),
        _kpi('k-green', len(rf_done), '100% COMPLETE'),
    ])
    page = page.replace('__META__', meta)
    page = page.replace('__SUBS__', json.dumps(subs, ensure_ascii=False))
    page = page.replace('__MILES__', json.dumps(miles or [], ensure_ascii=False))
    page = page.replace('__PUNCH__', json.dumps(pj, ensure_ascii=False))
    page = page.replace('__ITR__', json.dumps(ij, ensure_ascii=False))
    page = page.replace('__RFC__', json.dumps(rf, ensure_ascii=False))
    page = page.replace('__NOTES__', json.dumps(notes_js, ensure_ascii=False))
    page = page.replace('__PCOL__', json.dumps(colors_js, ensure_ascii=False))
    page = page.replace('__OVR__', json.dumps(state.get('cells', {}), ensure_ascii=False))
    page = page.replace('__ITRT__', json.dumps(itrt, ensure_ascii=False))
    page = page.replace('__PUNT__', json.dumps(punt, ensure_ascii=False))
    page = page.replace('__KPIS__', kpis)
    page = apply_extras(page)
    with open(OUT_HTML, 'w', encoding='utf-8') as f:
        f.write(page)
    return OUT_HTML


src_path_global = ''


def main():
    global src_path_global
    src = find_latest_dpr()
    if not src:
        print('ERROR: no DPR SUMMERY file found')
        raise SystemExit(1)
    src_path_global = src
    print(f'Source : {src}')
    m = re.search(r'-(\d{2}-\d{2}-\d{2})', os.path.basename(src))
    rep_date = m.group(1) if m else datetime.datetime.fromtimestamp(
        os.path.getmtime(src)).strftime('%d-%m-%y')

    wb_src = open_source(src)
    try:
        punch = extract_punch(wb_src['DETAILED PUNCH LIST'])
        itr = extract_itr(wb_src['DETAILED ITR LIST'])
        try:
            rfc = extract_rfc(wb_src['RFC PROGRESS'])
        except Exception:
            rfc = []
        itr_sum = sum_itr(wb_src['ITR SUMMARY'])
        pun_sum = sum_punch(wb_src['PUNCH SUMMERY'])
        ms_sheet = next((s for s in wb_src.sheetnames if s.upper().startswith('MILESTONE')),
                        wb_src.sheetnames[4])
        miles_list = extract_milestones(wb_src[ms_sheet])
        miles_n = len(miles_list)
    finally:
        wb_src.close()
    ov = find_ov_files()
    if ov.get('punch'):
        try:
            n0 = len(punch)
            punch, ch = merge_with_ov(punch, extract_ov_punch(ov['punch']))
            print(f"  OV punch : {os.path.basename(ov['punch'])} -> {n0}+{len(punch)-n0} rows, {ch} updates")
        except Exception as e:
            print(f'  OV punch FAILED: {e}')
    if ov.get('itr'):
        try:
            n0 = len(itr)
            itr, ch = merge_with_ov(itr, extract_ov_itr(ov['itr']))
            print(f"  OV itr   : {os.path.basename(ov['itr'])} -> {n0}+{len(itr)-n0} rows, {ch} updates")
        except Exception as e:
            print(f'  OV itr FAILED: {e}')
    print(f'  Punch rows: {len(punch)} | ITR rows: {len(itr)} | RFC: {len(rfc)} | Miles: {miles_n}')

    keys = build_keys(punch, itr)
    max_p = max((k['p'] for k in keys), default=1)
    max_i = max((k['i'] for k in keys), default=1)
    print(f'  Subsystems: {len(keys)} | max punch/sub: {max_p} | max itr/sub: {max_i}')

    state = load_previous()
    # platform edits persist via _explorer_state.json; source files always win on rebuild
    
    n_cells = sum(len(v) for v in (state.get('cells') or {}).values())
    n_notes = sum(len(v) for v in state['notes'].values())
    n_fills = sum(len(v) for v in state['fills'].values())
    print(f'  Carry-over: {n_notes} notes, {n_fills} colored rows, {n_cells} pinned cells')

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_search(wb, keys, rep_date, os.path.basename(src), punch, itr, max_p, max_i)
    build_detail(wb, 'PUNCH LIST', PUNCH_COLS, punch, None, state, is_done_punch, 'cat')
    build_detail(wb, 'ITR LIST', ITR_COLS, itr, None, state, is_done_itr, None)
    build_key_sheet(wb, keys)
    out = save_with_retry(wb)
    print(f'DONE -> {out}')
    hp = build_html(keys, punch, itr, rep_date, state, rfc,
                    itr_sum, pun_sum, miles_n, miles_list)
    print(f'DONE -> {hp}')


if __name__ == '__main__':
    main()
