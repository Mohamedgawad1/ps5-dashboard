"""
PS5 DPR - New Organized Summary Builder
Reads the latest 'PS-5 COMPLETIONS DPR SUMMERY -*.xlsx' (Downloads -> fallback project dir)
and builds a clean organized Excel:
  Sheet 1 'DAILY SUMMARY' : ITR Summary + Punch Summary combined
  Sheet 2 'MILESTONES'    : subsystem milestone data (source sheet 5)
"""
import os, re, glob, shutil, datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(BASE, '_temp')
DOWNLOADS = os.path.join(os.path.expanduser('~'), 'Downloads')
OUT_NAME = 'PS5 DPR - NEW SUMMARY.xlsx'

NAVY   = 'FF003366'
BLUE   = 'FF4472C4'
LBLUE  = 'FFD9E2F3'
ORANGE = 'FFED7D31'
WHITE  = 'FFFFFFFF'
LGREEN = 'FFC6EFCE'
LRED   = 'FFFFC7CE'
GRAY   = 'FFF2F2F2'

thin = Side(style='thin', color='FFBFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def n(v):
    if v is None: return ''
    s = str(v).strip()
    if s in ('#N/A', 'nan', 'None'): return ''
    return s

def num(v):
    v = n(v)
    try: return float(v)
    except ValueError: return None

def find_latest_dpr():
    cands = []
    for folder in (DOWNLOADS, BASE):
        for f in glob.glob(os.path.join(folder, '*.xlsx')):
            b = os.path.basename(f)
            if b.startswith('~$'): continue
            if 'completions dpr summery' in b.lower():
                cands.append((os.path.getmtime(f), f))
    return max(cands)[1] if cands else None

def src_report_date(path):
    m = re.search(r'-(\d{2}-\d{2}-\d{2})', os.path.basename(path))
    if m:
        try: return datetime.datetime.strptime(m.group(1), '%d-%m-%y').strftime('%d %b %Y')
        except ValueError: pass
    return datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime('%d %b %Y')

def open_source(path):
    dst = os.path.join(TMP, os.path.basename(path))
    try:
        shutil.copy2(path, dst)
    except PermissionError:
        if not os.path.exists(dst): raise
    return openpyxl.load_workbook(dst, read_only=True, data_only=True)

# ---------------- extraction ----------------

def extract_itr(ws):
    """Returns (date_cols[list of labels], week_total_label, rows[dict]).
    Column offsets are detected dynamically from the header row."""
    disc_row = None
    for i, row in enumerate(ws.iter_rows(min_col=1, values_only=True), 1):
        vals = [n(v) for v in row[:15]]
        if any(v.lower() == 'discipline' for v in vals):
            disc_row = i; break
    if disc_row is None: raise RuntimeError('ITR SUMMARY: Discipline header not found')

    grid = []
    for row in ws.iter_rows(min_row=disc_row, values_only=True):
        grid.append(list(row))

    base = next(i for i, v in enumerate(grid[0]) if n(v).lower() == 'discipline')
    tcol, ccol, pcol, ocol = base + 1, base + 2, base + 3, base + 4

    # date labels start after the Open column on the row below the header
    date_labels, week_lbl, j = [], 'Week Total', ocol + 1
    if len(grid) > 1:
        hdr = grid[1]
        while j < len(hdr):
            lab = n(hdr[j])
            if lab.lower() == 'total':
                week_lbl = lab.title(); break
            if lab: date_labels.append(lab)
            j += 1

    d0, d1 = ocol + 1, ocol + 1 + len(date_labels)
    rows, group = [], ''
    for g in grid[2:]:
        c3 = n(g[base]) if len(g) > base else ''
        c4 = num(g[tcol]) if len(g) > tcol else None
        c5 = num(g[ccol]) if len(g) > ccol else None
        c6 = num(g[pcol]) if len(g) > pcol else None
        c7 = num(g[ocol]) if len(g) > ocol else None
        dailies = [num(g[d0 + k]) if len(g) > d0 + k and g[d0 + k] is not None else None
                   for k in range(len(date_labels))]
        week = num(g[d1]) if len(g) > d1 else None
        if not c3:
            continue
        if c3.lower().startswith('*note'): break
        if c4 is None and c5 is None:
            group = c3.upper() if not c3.startswith('Grand') else group
            if c3.lower().startswith('grand'):
                rows.append({'group': '', 'name': c3, 'total': None, 'closed': None,
                             'pct': None, 'open': None, 'dailies': [], 'week': week})
            continue
        rows.append({'group': group, 'name': c3, 'total': c4, 'closed': c5,
                     'pct': c6, 'open': c7, 'dailies': dailies, 'week': week})
        if c3.lower().startswith('grand'): break
    return date_labels, week_lbl, rows

def extract_punch(ws):
    """Three blocks side by side: A(C..F) B(I..L) C(O..R). Returns dict disc -> {A/B/C: (t,c,o)}"""
    hdr_row = None
    rows = list(ws.iter_rows(min_col=3, max_col=18, values_only=True))
    for i, r in enumerate(rows):
        if n(r[0]).lower() == 'discipline':
            hdr_row = i; break
    if hdr_row is None: raise RuntimeError('PUNCH SUMMERY: header not found')
    blocks = {'A': 0, 'B': 6, 'C': 12}
    data = {}
    order = []
    for r in rows[hdr_row+1:]:
        name = n(r[0])
        if not name or name.lower().startswith('*note'): continue
        entry = {}
        for k, off in blocks.items():
            t, c, o = num(r[off+1]), num(r[off+2]), num(r[off+3])
            entry[k] = (int(t) if t is not None else 0,
                        int(c) if c is not None else 0,
                        int(o) if o is not None else 0)
        if name not in data:
            order.append(name)
        data[name] = entry
        if name.lower().startswith('grand'): break
    return order, data

def extract_milestones(ws):
    out = []
    for i, r in enumerate(ws.iter_rows(min_col=1, max_col=12, values_only=True), 1):
        if i == 1: continue
        sub = n(r[0])
        if not sub.startswith('PS5-'): continue
        desc = n(r[2]) or n(r[1])
        if 'responsible company' in desc.lower():
            desc = ''
        out.append({
            'sub': sub,
            'desc': desc,
            'milestone': n(r[3]),
            'month': n(r[4]),
            'mantrac': n(r[5]),
            'priority': n(r[6]),
            'pg': n(r[8]),
            'rfsu': n(r[9]),
            'mp_date': n(r[10]),
            'eit_date': n(r[11]),
        })
    return out

# ---------------- styling helpers ----------------

def style_title(ws, row, ncols, text, size=14, fill=NAVY):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row, 1); c.value = text
    c.font = Font(bold=True, color=WHITE, size=size)
    c.fill = PatternFill('solid', start_color=fill, end_color=fill)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[row].height = 24 if size >= 14 else 18

def style_section(ws, row, ncols, text, fill=ORANGE):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row, 1); c.value = text
    c.font = Font(bold=True, color=WHITE, size=11)
    c.fill = PatternFill('solid', start_color=fill, end_color=fill)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)

def hdr_cell(c, fill=NAVY):
    c.font = Font(bold=True, color=WHITE, size=10)
    c.fill = PatternFill('solid', start_color=fill, end_color=fill)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = BORDER

def fmt_date_token(s):
    try: return datetime.datetime.strptime(s, '%Y-%m-%d %H:%M:%S').strftime('%d-%b-%y')
    except ValueError: return s

# ---------------- builders ----------------

def build_daily_summary(wb, rep_date, src_name, itr, punch):
    date_labels, week_lbl, rows = itr
    order, pdata = punch
    ws = wb.create_sheet('DAILY SUMMARY')

    n_dates = max(1, len(date_labels))
    width = 7 + n_dates + 1          # Group|Disc|Total|Closed|%|Open|dates|week
    style_title(ws, 1, width, 'PS5  PRE-COMMISSIONING  DAILY  SUMMARY')
    sub = ws.cell(2, 1, f'Report Date: {rep_date}    |    Source: {src_name}')
    sub.font = Font(italic=True, size=9, color='FF404040')
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
    sub.alignment = Alignment(horizontal='center')

    r = 4
    style_section(ws, r, width, 'ITR  PROGRESS  BY  DISCIPLINE'); r += 1
    heads = ['Group', 'Discipline', 'Total', 'Closed', '%', 'Open'] + date_labels + [week_lbl]
    for ci, h in enumerate(heads, 1): hdr_cell(ws.cell(r, ci)); ws.cell(r, ci).value = h
    hr = r; r += 1

    grp_fill = {'CPP': LBLUE, 'EIT': GRAY}
    for row in rows:
        is_total = row['name'].lower() in ('total', 'grand total')
        vals = [row['group'], row['name'], row['total'], row['closed'], row['pct'], row['open']]
        vals += row['dailies'] + [row['week']]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(r, ci)
            if v is not None and v != '': c.value = v
            c.border = BORDER
            c.font = Font(bold=is_total, size=10)
            if ci == 5 and isinstance(row['pct'], float) and ci == 5:
                c.number_format = '0.0%'
            if ci > 6: c.alignment = Alignment(horizontal='center')
            if is_total:
                c.fill = PatternFill('solid', start_color='FFDDEBF7' if row['name'].lower() == 'total' else ORANGE,
                                     end_color='FFDDEBF7' if row['name'].lower() == 'total' else ORANGE)
                if row['name'].lower() == 'grand total':
                    c.font = Font(bold=True, size=10, color=WHITE)
            elif row['group'] in grp_fill and ci <= 2:
                c.fill = PatternFill('solid', start_color=grp_fill[row['group']], end_color=grp_fill[row['group']])
        r += 1

    r += 1
    style_section(ws, r, width, 'PUNCH  SUMMARY  ( A / B / C )'); r += 1
    heads = ['Discipline',
             'A Total', 'A Closed', 'A Open',
             'B Total', 'B Closed', 'B Open',
             'C Total', 'C Closed', 'C Open']
    for ci, h in enumerate(heads, 1):
        hdr_cell(ws.cell(r, ci),
                 fill=NAVY if ci == 1 else (BLUE if ci <= 4 else (ORANGE if ci <= 7 else 'FF70AD47')))
        ws.cell(r, ci).value = h
    r += 1
    for name in order:
        e = pdata[name]
        is_total = name.lower().startswith('grand')
        vals = [name] + list(e['A']) + list(e['B']) + list(e['C'])
        for ci, v in enumerate(vals, 1):
            c = ws.cell(r, ci); c.value = v; c.border = BORDER
            c.font = Font(bold=is_total, size=10)
            if ci > 1: c.alignment = Alignment(horizontal='center')
            if ci in (3, 6, 9) and isinstance(v, int) and v and not is_total:
                c.fill = PatternFill('solid', start_color=LGREEN, end_color=LGREEN)
            elif ci in (4, 7, 10) and isinstance(v, int) and v and not is_total:
                c.fill = PatternFill('solid', start_color=LRED, end_color=LRED)
            if is_total: c.fill = PatternFill('solid', start_color='FFDDEBF7', end_color='FFDDEBF7')
        r += 1

    widths = [12, 20, 9, 9, 8, 9] + [9]*n_dates + [11]
    for ci, w in enumerate(widths[:width], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = ws.cell(hr + 1, 1)
    return ws

def build_milestones(wb, milestones):
    ws = wb.create_sheet('MILESTONES')
    heads = ['Subsystem', 'Description', 'Milestone', 'Month', 'Mantrac',
             'Priority', 'PG/Common', 'RFSU', 'M&P Date', 'EIT Date']
    keys = ['sub', 'desc', 'milestone', 'month', 'mantrac',
            'priority', 'pg', 'rfsu', 'mp_date', 'eit_date']
    style_title(ws, 1, len(heads), 'PS5  SUBSYSTEM  MILESTONES')
    for ci, h in enumerate(heads, 1):
        hdr_cell(ws.cell(2, ci)); ws.cell(2, ci).value = h
    r = 3
    for m in milestones:
        band = GRAY if (r % 2 == 0) else WHITE
        for ci, k in enumerate(keys, 1):
            v = m[k]
            if k in ('mp_date', 'eit_date') and v: v = fmt_date_token(v)
            c = ws.cell(r, ci, v if v != '' else None)
            c.border = BORDER
            c.font = Font(size=9)
            c.fill = PatternFill('solid', start_color=band, end_color=band)
            if k in ('month', 'mantrac', 'priority', 'pg', 'rfsu'):
                c.alignment = Alignment(horizontal='center')
        pr = n(m['priority'])
        if pr.isdigit() or re.match(r'^\d+(\.\d+)?$', pr):
            p = float(pr)
            fill = LGREEN if p <= 1 else (GRAY if p < 2 else LRED)
            ws.cell(r, 6).fill = PatternFill('solid', start_color=fill, end_color=fill)
            ws.cell(r, 6).font = Font(size=9, bold=(p <= 1))
        r += 1
    widths = [13, 46, 17, 11, 10, 9, 11, 7, 12, 12]
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = 'A3'
    ws.auto_filter.ref = f'A2:{get_column_letter(len(heads))}{r-1}'
    return ws

def main():
    src = find_latest_dpr()
    if not src:
        print('ERROR: no DPR SUMMERY file found'); raise SystemExit(1)
    print(f'Source : {os.path.basename(src)}')
    rep_date = src_report_date(src)
    wb_src = open_source(src)
    try:
        itr = extract_itr(wb_src['ITR SUMMARY'])
        print(f"  ITR: {len(itr[2])} rows | dates: {itr[0]}")
        punch = extract_punch(wb_src['PUNCH SUMMERY'])
        print(f'  Punch: {len(punch[0])} disciplines')
        ms_sheet = next((s for s in wb_src.sheetnames if s.upper().startswith('MILESTONE')), wb_src.sheetnames[4])
        ms = extract_milestones(wb_src[ms_sheet])
        print(f'  Milestones: {len(ms)} subsystems ({ms_sheet})')
    finally:
        wb_src.close()

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)
    build_daily_summary(out_wb, rep_date, os.path.basename(src), itr, punch)
    build_milestones(out_wb, ms)
    out_path = os.path.join(BASE, OUT_NAME)
    try:
        out_wb.save(out_path)
    except PermissionError:
        alt = os.path.join(TMP, OUT_NAME)
        out_wb.save(alt); shutil.copy2(alt, out_path)
    print(f'DONE -> {out_path}')

if __name__ == '__main__':
    main()
