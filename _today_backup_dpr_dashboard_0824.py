"""
PS5 DPR ONE-SHEET DASHBOARD
Reads latest 'PS-5 COMPLETIONS DPR SUMMERY -*.xlsx' -> builds PS5 DPR DASHBOARD.xlsx:
  Sheet 1 'DASHBOARD'   : ITR SUMMARY (src 1) + PUNCH SUMMERY (src 2) + MILESTONES (src 5)
  Sheet 2 'RFC PROGRESS': modern tracker - click a subsystem in col A -> opens
                          subsystem_explorer.html#SID with ALL its Punch + ITR data
  Sheet 3 'PUNCH LIST'  : detailed punch (search/filter)
  Sheet 4 'ITR LIST'    : detailed ITR (search/filter)
User notes & custom colors are carried over every daily rebuild.
"""
import os, re, datetime, shutil
import urllib.parse
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.utils import get_column_letter, column_index_from_string

from dpr_new_summary import (extract_itr as sum_itr, extract_punch as sum_punch,
                             extract_milestones, src_report_date, fmt_date_token)
from punch_itr_explorer import (DL_SUB, TMP, NAVY, BLUE, ORANGE, WHITE, GRAY,
                                GREEN, RED, YELLOW, DGREEN, SEL_FILL, OWN_COLORS,
                                BORDER, hdr_cell, style_title, cat_fill,
                                find_latest_dpr, open_source,
                                extract_punch as det_punch, extract_itr as det_itr,
                                extract_rfc, load_platform_state,
                                build_keys, is_done_punch, is_done_itr,
                                build_detail, REC_KEYS, save_with_retry,
                                load_previous, capture_cells, DETAIL_COLMAP,
                                RFC_COLMAP, rfc_row_sid, RFC_DISCS_L,
                                find_ov_files, extract_ov_punch, extract_ov_itr,
                                merge_with_ov)
from copy import copy as _style_copy

OUT_XLSX = os.path.join(DL_SUB, 'PS5 DPR DASHBOARD.xlsx')
OUT_XLSM = os.path.join(DL_SUB, 'PS5 DPR DASHBOARD.xlsm')
MACRO_BIN = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'rfc_vba.bin')
EXPLORER_HTML = os.path.join(DL_SUB, 'subsystem_explorer.html')
PAGES_URL = 'https://mohamedgawad1.github.io/PS5-COMPLETION-PLATFORM/'

SEC_ITR = 'ITR  PROGRESS  BY  DISCIPLINE'
SEC_PUN = 'PUNCH  SUMMARY  ( A / B / C )'
SEC_MIL = 'SUBSYSTEM  MILESTONES'


def clean(v):
    return v


def n_(v):
    if v is None:
        return ''
    s = str(v).strip()
    return '' if s in ('#N/A', '#VALUE!', 'nan', 'None') else s


def pct_fill(p):
    if p is None:
        return None
    if p >= 0.9999:
        return GREEN
    if p >= 0.5:
        return YELLOW
    if p > 0:
        return RED
    return DGREEN


def prio_style(cell, pr):
    try:
        p = float(pr)
    except (TypeError, ValueError):
        return
    fill = GREEN if p <= 1 else (GRAY if p < 2 else RED)
    cell.fill = PatternFill('solid', start_color=fill, end_color=fill)
    cell.font = Font(size=10, bold=(p <= 1))


# ---------------- carry-over ----------------

def load_prev_dashboard():
    st = {'notes': {}, 'fills': {}}
    prev = OUT_XLSM if os.path.exists(OUT_XLSM) else OUT_XLSX
    if not os.path.exists(prev):
        return st
    try:
        wb = openpyxl.load_workbook(prev)
    except Exception:
        return st
    if 'DASHBOARD' not in wb.sheetnames:
        wb.close(); return st
    ws = wb['DASHBOARD']
    sec, note_col = None, None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        av = str(a).strip() if a else ''
        if av.startswith('ITR  PROGRESS'):
            sec, note_col = 'ITR', None; continue
        if av.startswith('PUNCH  SUMMARY'):
            sec, note_col = 'PUNCH', None; continue
        if av.startswith('SUBSYSTEM  MILESTONES'):
            sec = 'MIL'
            for c in range(1, ws.max_column + 1):
                h = ws.cell(r + 1, c).value
                if h and str(h).strip().upper() == 'USER NOTES':
                    note_col = c; break
            continue
        if not sec:
            continue
        keycol = 2 if sec in ('ITR', 'PUNCH') else 1
        kv = ws.cell(r, keycol).value
        if not kv or str(kv).startswith('='):
            continue
        kv = str(kv).strip()
        if note_col:
            nv = ws.cell(r, note_col).value
            if nv and str(nv).strip():
                st['notes'][kv] = str(nv).strip()
        row_fills = {}
        for c in range(1, ws.max_column + 1):
            f = ws.cell(r, c).fill
            if f is None or f.patternType != 'solid':
                continue
            rgb = getattr(f.fgColor, 'rgb', None)
            if isinstance(rgb, str) and len(rgb) == 8 and rgb not in OWN_COLORS:
                row_fills[c] = rgb
        if row_fills:
            st['fills'][kv] = row_fills
    wb.close()
    return st


def apply_user_colors(ws, r, key, st):
    for c, rgb in st['fills'].get(key, {}).items():
        ws.cell(r, c).fill = PatternFill('solid', start_color=rgb, end_color=rgb)


# ---------------- DASHBOARD sheet ----------------

def build_dashboard(wb, rep_date, src_name, itr, punch_sum, miles, st):
    ws = wb.create_sheet('DASHBOARD')
    date_labels, week_lbl, irows = itr
    order, pdata = punch_sum
    width = max(12, 7 + len(date_labels) + 1)
    style_title(ws, 1, width, 'PS5  PRE-COMMISSIONING  DAILY  DASHBOARD')
    sub = ws.cell(2, 1, f'Report Date: {rep_date}    |    Source: {src_name}')
    sub.font = Font(italic=True, size=9, color='FF404040')
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
    sub.alignment = Alignment(horizontal='center')

    pl = ws.cell(3, 1, f'=HYPERLINK("{PAGES_URL}","OPEN  PLATFORM  ONLINE   (works from any PC / phone)")')
    pl.font = Font(bold=True, size=11, color=WHITE, underline='single')
    pl.fill = PatternFill('solid', start_color='FF70AD47', end_color='FF70AD47')
    pl.alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=width)
    ws.row_dimensions[3].height = 20

    r = 6
    style_section(ws, r, width, SEC_ITR); a_itr = r; r += 1
    heads = ['Group', 'Discipline', 'Total', 'Closed', '%', 'Open'] + date_labels + [week_lbl]
    for ci, h in enumerate(heads, 1):
        hdr_cell(ws.cell(r, ci)); ws.cell(r, ci).value = h
    r += 1
    grp_fill = {'CPP': 'FFD9E2F3', 'EIT': GRAY}
    for row in irows:
        is_total = row['name'].lower() in ('total', 'grand total')
        vals = [row['group'], row['name'], row['total'], row['closed'],
                row['pct'], row['open']] + row['dailies'] + [row['week']]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(r, ci)
            if v is not None and v != '':
                c.value = v
            c.border = BORDER
            c.font = Font(bold=is_total, size=10)
            if ci == 5 and isinstance(row['pct'], float):
                c.number_format = '0.0%'
            if ci > 6:
                c.alignment = Alignment(horizontal='center')
            if is_total:
                c.fill = PatternFill('solid', start_color='FFDDEBF7' if row['name'].lower() == 'total' else ORANGE,
                                     end_color='FFDDEBF7' if row['name'].lower() == 'total' else ORANGE)
                if row['name'].lower() == 'grand total':
                    c.font = Font(bold=True, size=10, color=WHITE)
            elif row['group'] in grp_fill and ci <= 2:
                c.fill = PatternFill('solid', start_color=grp_fill[row['group']], end_color=grp_fill[row['group']])
        apply_user_colors(ws, r, row['name'], st)
        r += 1

    r += 1
    style_section(ws, r, width, SEC_PUN); a_pun = r; r += 1
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
            if ci > 1:
                c.alignment = Alignment(horizontal='center')
            if ci in (3, 6, 9) and isinstance(v, int) and v and not is_total:
                c.fill = PatternFill('solid', start_color=GREEN, end_color=GREEN)
            elif ci in (4, 7, 10) and isinstance(v, int) and v and not is_total:
                c.fill = PatternFill('solid', start_color=RED, end_color=RED)
            if is_total:
                c.fill = PatternFill('solid', start_color='FFDDEBF7', end_color='FFDDEBF7')
        apply_user_colors(ws, r, name, st)
        r += 1

    r += 1
    style_section(ws, r, width, SEC_MIL); a_mil = r; r += 1
    heads = ['Subsystem', 'Description', 'Milestone', 'Month', 'Mantrac',
             'Priority', 'PG/Common', 'RFSU', 'M&P Date', 'EIT Date', 'USER NOTES']
    keys = ['sub', 'desc', 'milestone', 'month', 'mantrac', 'priority',
            'pg', 'rfsu', 'mp_date', 'eit_date']
    mil_hdr = r
    for ci, h in enumerate(heads, 1):
        hdr_cell(ws.cell(r, ci)); ws.cell(r, ci).value = h
    r += 1
    for m in miles:
        band = GRAY if (r % 2 == 0) else WHITE
        for ci, k in enumerate(keys, 1):
            v = m[k]
            if k in ('mp_date', 'eit_date') and v:
                v = fmt_date_token(v)
            c = ws.cell(r, ci, v if v != '' else None)
            c.border = BORDER
            c.font = Font(size=9)
            c.fill = PatternFill('solid', start_color=band, end_color=band)
            if k in ('month', 'mantrac', 'priority', 'pg', 'rfsu'):
                c.alignment = Alignment(horizontal='center')
        pr = n_(m['priority'])
        if pr.replace('.', '').isdigit():
            prio_style(ws.cell(r, 6), pr)
        nc = ws.cell(r, len(heads), st['notes'].get(m['sub'], '') or None)
        nc.border = BORDER; nc.font = Font(size=9)
        apply_user_colors(ws, r, m['sub'], st)
        r += 1
    ws.auto_filter.ref = f'A{mil_hdr}:{get_column_letter(len(heads))}{r-1}'

    # quick navigation bar (row 4)
    nav = [('ITR PROGRESS', a_itr, BLUE), ('PUNCH SUMMARY', a_pun, 'FF70AD47'),
           ('MILESTONES', a_mil, ORANGE), ('BACK TO TOP', 1, NAVY)]
    col = 1
    for label, target, fill in nav:
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 2)
        c = ws.cell(4, col, f'=HYPERLINK("#DASHBOARD!A{target}","{label}")')
        c.font = Font(bold=True, color=WHITE, size=10, underline='single')
        c.fill = PatternFill('solid', start_color=fill, end_color=fill)
        c.alignment = Alignment(horizontal='center', vertical='center')
        col += 3
    ws.row_dimensions[4].height = 20

    widths = [13, 24, 9, 9, 8, 9] + [9] * len(date_labels) + [11]
    for ci, w in enumerate(widths[:width], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.column_dimensions['B'].width = 30
    ws.freeze_panes = 'A5'
    return ws


def style_section(ws, row, ncols, text, fill=ORANGE):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row, 1); c.value = text
    c.font = Font(bold=True, color=WHITE, size=11)
    c.fill = PatternFill('solid', start_color=fill, end_color=fill)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)


def load_prev_rfc_remarks():
    """MY REMARKS per sid: previous workbook + platform state file."""
    out = {}
    prev = OUT_XLSM if os.path.exists(OUT_XLSM) else OUT_XLSX
    if os.path.exists(prev):
        try:
            wb = openpyxl.load_workbook(prev, read_only=True)
            if 'RFC PROGRESS' in wb.sheetnames:
                ws = wb['RFC PROGRESS']
                rem_col = None
                for c in range(1, 70):
                    for hr in (1, 2, 3):
                        v = ws.cell(hr, c).value
                        if v and str(v).strip().upper() == 'MY REMARKS':
                            rem_col = c
                            break
                    if rem_col:
                        break
                if rem_col:
                    for r in range(4, ws.max_row + 1):
                        a = ws.cell(r, 1).value
                        if not a:
                            continue
                        s = str(a)
                        sid = ''
                        m = re.match(r'=HYPERLINK\("[^"]*#([^"]+)",".*"\)', s)
                        if m:
                            sid = m.group(1)
                        else:
                            m2 = re.match(r'\s*(PS5-\d{2}-\d{2}[A-Z0-9-]*)\s+-', s)
                            if m2:
                                sid = m2.group(1)
                        if not sid:
                            continue
                        v = ws.cell(r, rem_col).value
                        if v and str(v).strip():
                            out[sid] = str(v).strip()
            wb.close()
        except Exception:
            pass
    plat = load_platform_state()['notes'].get('RFC PROGRESS', {})
    out.update(plat)
    return out


# ---------------- RFC PROGRESS sheet ----------------

def snapshot_rfc_header(src_ws):
    """Clone the header block (rows 1-3) of the SOURCE RFC PROGRESS sheet."""
    vals, styles = {}, {}
    for r in (1, 2, 3):
        for c in range(1, 57):
            s = src_ws.cell(r, c)
            solid = s.fill is not None and s.fill.patternType == 'solid'
            if s.value is None and not solid:
                continue
            v = s.value
            if isinstance(v, str) and v.startswith('='):
                v = ''
            vals[(r, c)] = v
            styles[(r, c)] = _style_copy(s._style)
    merges = [str(m) for m in src_ws.merged_cells.ranges if m.max_row <= 3]
    widths = {}
    for k, dim in src_ws.column_dimensions.items():
        try:
            ci = column_index_from_string(k)
        except Exception:
            continue
        if ci < 57 and dim.width:
            widths[ci] = round(dim.width, 1)
    heights = {r: src_ws.row_dimensions[r].height for r in (1, 2, 3)
               if src_ws.row_dimensions[r].height}
    return {'vals': vals, 'styles': styles, 'merges': merges,
            'widths': widths, 'heights': heights}


def build_rfc(wb, rfc_rows, html_link_base, my_remarks=None):
    """Styled tracker (113133 look) with source-matching column set."""
    my_remarks = my_remarks or {}
    ws = wb.create_sheet('RFC PROGRESS')
    discs = RFC_DISCS_L
    dnames = {'B': 'BUILDING', 'E': 'ELECT', 'H': 'HVAC', 'I': 'INSTR',
              'M': 'MECH', 'P': 'PIPING', 'S': 'SAFETY', 'T': 'TELECOM'}
    LASTC = 2 + 4 + 1 + 1 + 1 + len(discs) * 4 + 4 + 4 + 3 + 3 + 1 + 1
    style_title(ws, 1, LASTC, 'RFC  PROGRESS  -  CLICK SUBSYSTEM IN COLUMN A FOR FULL PUNCH + ITR DETAILS')

    spans = [('SUBSYSTEM (CLICK) ->', 2, NAVY), ('RFC DATES', 4, BLUE),
             ('RFC SIGNED', 1, 'FF70AD47'), ('MILESTONE', 1, BLUE),
             ('TOTAL %', 1, ORANGE)]
    for dk in discs:
        spans.append((dnames[dk], 4, BLUE))
    spans.append(('ITR TOTALS', 4, ORANGE))
    spans.append(('', 4, WHITE))
    spans.append(('BLOCKING ITRs', 3, 'FFC000'))
    spans.append(('BLOCKING POINTS REMARKS', 3, NAVY))
    spans.append(('WALKDOWN', 1, 'FF70AD47'))
    spans.append(('MY REMARKS (SYNCED)', 1, RED))
    cc = 1
    for label, span, fill in spans:
        ws.merge_cells(start_row=2, start_column=cc, end_row=2, end_column=cc + span - 1)
        cell = ws.cell(2, cc, label)
        hdr_cell(cell, fill=fill)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cc += span
    hd = ['', 'Priority', 'RFC BHMPS', 'RFC EIT', 'Baseline', 'Recovery',
          'SIGNED', 'Milestone', 'TOTAL %']
    for dk in discs:
        hd += ['Total', 'Closed', 'Bal', '%']
    hd += ['ITRs', 'Closed', 'Balance', 'ITR %', '', '', '', '',
           'CPP-1', 'EIT', 'EACOP',
           'EACOP', 'CPP-EIT', 'CPP-1',
           'STATUS', 'MY REMARKS']
    for ci, h in enumerate(hd, 1):
        hdr_cell(ws.cell(3, ci), fill='FF335A82')
        if h:
            ws.cell(3, ci).value = h
    ws.row_dimensions[3].height = 22

    link = f'{html_link_base}#'
    r = 4
    for rec in rfc_rows:
        band = GRAY if r % 2 else WHITE

        def put(ci, v, numfmt=None, fill=None, bold=False, size=9, wrap=False):
            c = ws.cell(r, ci, v if v not in ('', None) else None)
            c.font = Font(size=size, bold=bold)
            c.border = BORDER
            c.alignment = Alignment(horizontal='center', vertical='center',
                                    wrap_text=wrap)
            if numfmt == 'pct':
                c.number_format = '0%'
            f2 = fill or band
            c.fill = PatternFill('solid', start_color=f2, end_color=f2)
            return c

        a = ws.cell(r, 1, rec['name'])
        a.hyperlink = Hyperlink(ref=a.coordinate, location="'PUNCH LIST'!A1",
                                tooltip='Click: show only this subsystem in PUNCH + ITR')
        a.font = Font(size=9, color='FF0563C1', underline='single')
        a.fill = PatternFill('solid', start_color=SEL_FILL, end_color=SEL_FILL)
        a.border = BORDER
        a.alignment = Alignment(vertical='center')

        pc = put(2, rec['prio'])
        prio_style(pc, rec['prio'])
        put(3, rec['rfc_bhm']); put(4, rec['rfc_eit'])
        put(5, rec['baseline']); put(6, rec.get('recovery', ''))
        sg = str(rec.get('signed', '') or '')
        sc = put(7, sg, bold=True)
        if sg.upper() == 'COMPLETE':
            sc.fill = PatternFill('solid', start_color=GREEN, end_color=GREEN)
        elif sg.upper() == 'PARTIAL':
            sc.fill = PatternFill('solid', start_color=YELLOW, end_color=YELLOW)
        put(8, rec['milestone'], bold=True)

        tp = put(9, rec['tot_pct'], numfmt='pct', size=10, bold=True)
        pf = pct_fill(rec['tot_pct']) or DGREEN
        tp.font = Font(size=10, bold=True, color=WHITE)
        tp.fill = PatternFill('solid', start_color=pf, end_color=pf)

        ci = 10
        for dk in RFC_DISCS_L:
            dd = rec['disc'][dk]
            put(ci, dd['t']); put(ci + 1, dd['c']); put(ci + 2, dd['b'])
            pp = put(ci + 3, dd['p'], numfmt='pct')
            ppf = pct_fill(dd['p'])
            if ppf:
                pp.fill = PatternFill('solid', start_color=ppf, end_color=ppf)
            elif dd['p'] == 0:
                pp.fill = PatternFill('solid', start_color=DGREEN, end_color=DGREEN)
            ci += 4
        put(42, rec['itrs'], bold=True); put(43, rec['closed'])
        put(44, rec['balance'])
        ip = put(45, rec['itr_pct'], numfmt='pct', size=10, bold=True)
        ipf = pct_fill(rec['itr_pct']) or DGREEN
        ip.font = Font(size=10, bold=True, color=WHITE)
        ip.fill = PatternFill('solid', start_color=ipf, end_color=ipf)

        for ci2, key in ((50, 'blk_cpp1'), (51, 'blk_eit'), (52, 'blk_eacop')):
            v = rec[key]
            c = put(ci2, int(v) if isinstance(v, float) and v == int(v) else v,
                    fill=RED if v else None, bold=bool(v))
            if v:
                c.font = Font(size=9, bold=True, color=WHITE)
        for ci2, key in ((53, 'rem_eacop'), (54, 'rem_cpeit'), (55, 'rem_cpp1')):
            put(ci2, rec[key], size=8, wrap=True, fill='FFFFF2CC')
        put(56, rec['walkdown'])

        mr = ws.cell(r, LASTC, my_remarks.get(rec['sid']) or None)
        mr.font = Font(size=9); mr.border = BORDER
        mr.alignment = Alignment(vertical='center', wrap_text=True)
        mr.fill = PatternFill('solid', start_color=SEL_FILL, end_color=SEL_FILL)
        r += 1

    last = r - 1
    ws.auto_filter.ref = f'A3:{get_column_letter(LASTC)}{last}'
    fallback = {1: 34, 2: 7, 9: 9, 42: 7, 43: 8, 44: 9, 45: 8,
                50: 7, 51: 7, 52: 8, 53: 24, 54: 24, 55: 24, 56: 12, 57: 28}
    for ci in range(1, LASTC + 1):
        ws.column_dimensions[get_column_letter(ci)].width = fallback.get(ci, 7)
    ws.freeze_panes = 'B4'
    return ws


def build_rfc_complete(wb, rfc_rows, html_link_base):
    ws = wb.create_sheet('RFC COMPLETE')
    done = [x for x in rfc_rows
            if (x['tot_pct'] is not None and x['tot_pct'] >= 0.9999)
            or str(x.get('signed', '') or '').upper() in ('COMPLETE', 'PARTIAL')]
    done.sort(key=lambda x: x['sid'])
    ncols = 5
    style_title(ws, 1, ncols, f'RFC  COMPLETE  -  100% OR SIGNED COMPLETE/PARTIAL  ({len(done)} of {len(rfc_rows)})')
    hd = ['#', 'SUB SYSTEM (CLICK)', 'TOTAL %', 'WALKDOWN STATUS', 'RFC SIGNED']
    for ci, h in enumerate(hd, 1):
        hdr_cell(ws.cell(2, ci), fill='FF70AD47' if ci == 4 else NAVY)
        ws.cell(2, ci).value = h
    link = f'{html_link_base}#'
    r = 3
    for i, rec in enumerate(done, 1):
        band = GRAY if r % 2 else WHITE
        c = ws.cell(r, 1, i)
        c.alignment = Alignment(horizontal='center')
        c.border = BORDER
        c.font = Font(size=9, bold=True)
        full = rec['tot_pct'] is not None and rec['tot_pct'] >= 0.9999
        c.fill = PatternFill('solid', start_color=GREEN if full else YELLOW,
                             end_color=GREEN if full else YELLOW)
        a = ws.cell(r, 2, rec['name'])
        a.hyperlink = Hyperlink(ref=a.coordinate, location="'PUNCH LIST'!A1",
                                tooltip='Click: show only this subsystem in PUNCH + ITR')
        a.font = Font(size=9, color='FF0563C1', underline='single')
        a.fill = PatternFill('solid', start_color=band, end_color=band)
        a.border = BORDER
        tp = ws.cell(r, 3, rec['tot_pct'])
        tp.number_format = '0%'
        tp.font = Font(size=9, bold=True)
        tp.border = BORDER
        tp.alignment = Alignment(horizontal='center')
        tpf = pct_fill(rec['tot_pct']) or band
        tp.fill = PatternFill('solid', start_color=tpf, end_color=tpf)
        wd = ws.cell(r, 4, rec['walkdown'] or None)
        wd.border = BORDER
        wd.font = Font(size=9)
        wd.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        if rec['walkdown']:
            wd.fill = PatternFill('solid', start_color='FFFFF2CC', end_color='FFFFF2CC')
        else:
            wd.fill = PatternFill('solid', start_color=band, end_color=band)
        sg = ws.cell(r, 4, rec['signed'] or None)
        sg.border = BORDER
        sg.font = Font(size=9)
        sg.alignment = Alignment(horizontal='center', vertical='center')
        if str(rec['signed']).upper() == 'COMPLETE':
            sg.fill = PatternFill('solid', start_color='FFC6EFCE', end_color='FFC6EFCE')
        elif str(rec['signed']).upper() == 'PARTIAL':
            sg.fill = PatternFill('solid', start_color='FFFFEB9C', end_color='FFFFEB9C')
        else:
            sg.fill = PatternFill('solid', start_color=band, end_color=band)
        r += 1
    ws.auto_filter.ref = f'A2:D{max(r-1, 2)}'
    for ci, w in enumerate([6, 44, 18, 14], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = 'A3'
    return ws


SHEET_CODENAMES = {'DASHBOARD': 'SH_DASH', 'RFC PROGRESS': 'SH_RFC',
                   'RFC COMPLETE': 'SH_RFCC', 'PUNCH LIST': 'SH_PUNCH',
                   'ITR LIST': 'SH_ITR'}
WB_CODENAME = 'ThisWorkbook'


def patch_codenames(path):
    """Pin stable VBA codeNames so the injected macro binds on every rebuild."""
    import zipfile
    zin = zipfile.ZipFile(path)
    wbx = zin.read('xl/workbook.xml').decode('utf-8')
    rels = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid_target = {}
    for m in re.finditer(r'<Relationship [^>]*>', rels):
        tag = m.group(0)
        if 'relationships/worksheet' not in tag:
            continue
        tm = re.search(r'Target="([^"]*)"', tag)
        im = re.search(r'Id="rId(\d+)"', tag)
        if tm and im:
            tg = tm.group(1).lstrip('/')
            if not tg.startswith('xl/'):
                tg = 'xl/' + tg
            rid_target[int(im.group(1))] = tg
    targets = {}
    for m in re.finditer(r'<sheet [^>]*>', wbx):
        tag = m.group(0)
        nm = re.search(r'name="([^"]+)"', tag)
        rid = re.search(r'r:id="rId(\d+)"', tag)
        if nm and rid and nm.group(1) in SHEET_CODENAMES:
            t = rid_target.get(int(rid.group(1)))
            if t:
                targets[t] = SHEET_CODENAMES[nm.group(1)]
    data = {}
    for it in zin.infolist():
        b = zin.read(it.filename)
        if it.filename == 'xl/workbook.xml':
            s = b.decode('utf-8')
            if 'codeName=' not in s:
                if '<workbookPr' in s:
                    s = re.sub(r'<workbookPr(\s|/|>)',
                               lambda mm: f'<workbookPr codeName="{WB_CODENAME}"{mm.group(1)}',
                               s, count=1)
                else:
                    j = s.index('>', s.index('<workbook')) + 1
                    s = s[:j] + f'<workbookPr codeName="{WB_CODENAME}"/>' + s[j:]
            b = s.encode('utf-8')
        elif it.filename in targets:
            s = b.decode('utf-8')
            cn = targets[it.filename]
            if 'codeName=' not in s.split('<dimension')[0]:
                if '<sheetPr' in s.split('<dimension')[0]:
                    s = re.sub(r'<sheetPr(\s|/|>)',
                               lambda mm: f'<sheetPr codeName="{cn}"{mm.group(1)}',
                               s, count=1)
                else:
                    j = s.index('>', s.index('<worksheet')) + 1
                    s = s[:j] + f'<sheetPr codeName="{cn}"/>' + s[j:]
            b = s.encode('utf-8')
        data[it.filename] = b
    zin.close()
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for fn, bb in data.items():
            zout.writestr(fn, bb)


def make_xlsm(xlsx_path):
    """Convert the finished xlsx into a macro-enabled xlsm carrying rfc_vba.bin."""
    if not os.path.exists(MACRO_BIN):
        return None
    import zipfile
    import time
    tmp_x = os.path.join(TMP, os.path.basename(OUT_XLSM))
    with zipfile.ZipFile(xlsx_path) as zin:
        ct = zin.read('[Content_Types].xml').decode('utf-8')
        ct = ct.replace(
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
            'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
        if 'Extension="bin"' not in ct:
            ct = ct.replace(
                '</Types>',
                '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
        rels = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        if 'vbaProject' not in rels:
            rid = 1
            for m in re.finditer(r'Id="rId(\d+)"', rels):
                rid = max(rid, int(m.group(1)) + 1)
            rels = rels.replace(
                '</Relationships>',
                f'<Relationship Id="rId{rid}" Type="http://schemas.microsoft.com/'
                f'office/2006/relationships/vbaProject" Target="vbaProject.bin"/>'
                '</Relationships>')
        with zipfile.ZipFile(tmp_x, 'w', zipfile.ZIP_DEFLATED) as zout:
            zout.writestr('[Content_Types].xml', ct)
            for it in zin.infolist():
                if it.filename == '[Content_Types].xml':
                    continue
                if it.filename == 'xl/_rels/workbook.xml.rels':
                    zout.writestr(it.filename, rels)
                else:
                    zout.writestr(it.filename, zin.read(it.filename))
            with open(MACRO_BIN, 'rb') as fb:
                zout.writestr('xl/vbaProject.bin', fb.read())
    for attempt in range(15):
        try:
            shutil.copy2(tmp_x, OUT_XLSM)
            return OUT_XLSM
        except PermissionError:
            print(f'  [wait] PS5 DPR DASHBOARD.xlsm is open in Excel - close it ... ({attempt + 1}/15)')
            time.sleep(2)
    stamp = datetime.datetime.now().strftime('%H%M%S')
    final = OUT_XLSM.replace('.xlsm', f' {stamp}.xlsm')
    shutil.copy2(tmp_x, final)
    return final


def build_milestones_sheet(wb, miles, st):
    ws = wb.create_sheet('MILESTONE & DIS.')
    heads = ['Subsystem', 'Description', 'Milestone', 'Month', 'Mantrac',
             'Priority', 'PG/Common', 'RFSU', 'M&P Date', 'EIT Date', 'USER NOTES']
    width = len(heads)
    style_title(ws, 1, width, 'MILESTONE & DIS.', size=14)
    for ci, h in enumerate(heads, 1):
        hdr_cell(ws.cell(2, ci)); ws.cell(2, ci).value = h
    keys = ['sub', 'desc', 'milestone', 'month', 'mantrac', 'priority',
            'pg', 'rfsu', 'mp_date', 'eit_date']
    r = 3
    for m in miles:
        band = GRAY if (r % 2 == 0) else WHITE
        for ci, k in enumerate(keys, 1):
            v = m[k]
            if k in ('mp_date', 'eit_date') and v:
                v = fmt_date_token(v)
            c = ws.cell(r, ci, v if v != '' else None)
            c.border = BORDER
            c.font = Font(size=9)
            c.fill = PatternFill('solid', start_color=band, end_color=band)
            if k in ('month', 'mantrac', 'priority', 'pg', 'rfsu'):
                c.alignment = Alignment(horizontal='center')
        pr = n_(m['priority'])
        if pr.replace('.', '').isdigit():
            prio_style(ws.cell(r, 6), pr)
        nc = ws.cell(r, width, st['notes'].get(m['sub'], '') or None)
        nc.border = BORDER
        nc.font = Font(size=9)
        apply_user_colors(ws, r, m['sub'], st)
        r += 1
    ws.auto_filter.ref = f'A2:{get_column_letter(width)}{r - 1}'
    ws.freeze_panes = 'A3'
    for col, w in zip('ABCDEFGHIJK',
                      (12, 46, 16, 9, 11, 9, 10, 10, 12, 12, 34)):
        ws.column_dimensions[col].width = w


def main():
    src = find_latest_dpr()
    if not src:
        print('ERROR: no DPR SUMMERY found'); raise SystemExit(1)
    print(f'Source : {src}')
    rep_date = src_report_date(src)
    wb_src = open_source(src)
    try:
        itr = sum_itr(wb_src['ITR SUMMARY'])
        punch_sum = sum_punch(wb_src['PUNCH SUMMERY'])
        ms_sheet = next((s for s in wb_src.sheetnames if s.upper().startswith('MILESTONE')),
                        wb_src.sheetnames[4])
        miles = extract_milestones(wb_src[ms_sheet])
        dp = det_punch(wb_src['DETAILED PUNCH LIST'])
        di = det_itr(wb_src['DETAILED ITR LIST'])
        ws_rfc_src = wb_src['RFC PROGRESS']
        rfc_rows = extract_rfc(ws_rfc_src)
    finally:
        wb_src.close()
    ov = find_ov_files()
    if ov.get('punch'):
        try:
            dp, _ = merge_with_ov(dp, extract_ov_punch(ov['punch']))
        except Exception as e:
            print(f'  OV punch FAILED: {e}')
    if ov.get('itr'):
        try:
            di, _ = merge_with_ov(di, extract_ov_itr(ov['itr']))
        except Exception as e:
            print(f'  OV itr FAILED: {e}')
    print(f'  ITR rows {len(itr[2])} | Punch disc {len(punch_sum[0])} | Miles {len(miles)} | '
          f'Detail P {len(dp)} | Detail I {len(di)} | RFC {len(rfc_rows)}')

    st = load_prev_dashboard()
    st_prev = OUT_XLSM if os.path.exists(OUT_XLSM) else OUT_XLSX
    wbst = load_previous(st_prev)
    cap = capture_cells(st_prev, {
        'RFC PROGRESS': {'hrow': 3, 'colmap': RFC_COLMAP, 'idfn': lambda r: r['sid'],
                         'rowid': rfc_row_sid, 'rows': rfc_rows},
        'PUNCH LIST': {'hrow': 2, 'colmap': DETAIL_COLMAP['PUNCH LIST'],
                       'idfn': lambda r: r.get('id', ''),
                       'rowid': lambda ws, rr: n_(ws.cell(rr, 1).value), 'rows': dp},
        'ITR LIST': {'hrow': 2, 'colmap': DETAIL_COLMAP['ITR LIST'],
                     'idfn': lambda r: r.get('id', ''),
                     'rowid': lambda ws, rr: n_(ws.cell(rr, 1).value), 'rows': di},
    })
    wbst['cells'] = cap.get('cells', {})
    n_cells = sum(len(v) for v in wbst['cells'].values())
    n_notes = sum(len(v) for v in wbst['notes'].values())
    n_pcol = sum(len(v) for v in wbst['pcolors'].values())
    print(f"  Carry-over: dash {len(st['notes'])} notes | detail {n_notes} notes | "
          f"{len(st['fills'])} colored rows | {n_pcol} platform colors | {n_cells} pinned cells")

    keys = build_keys(dp, di)
    max_p = max((k['p'] for k in keys), default=1)
    max_i = max((k['i'] for k in keys), default=1)

    PAGES_URL_ = PAGES_URL
    base = PAGES_URL_

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_dashboard(wb, rep_date, os.path.basename(src), itr, punch_sum, miles, st)
    build_milestones_sheet(wb, miles, st)
    my_rem = load_prev_rfc_remarks()
    build_rfc(wb, rfc_rows, base, my_rem)
    build_rfc_complete(wb, rfc_rows, base)
    build_detail(wb, 'PUNCH LIST', ['ID', 'TAG', 'CAT', 'DISC', 'DESCRIPTION', 'RESP COMPANY',
                                    'SUB SYSTEM', 'CLOSING DATE', 'STATUS', 'ORIGINATED BY',
                                    'PRIORITY', 'USER NOTES'], dp, None, wbst, is_done_punch, 'cat')
    build_detail(wb, 'ITR LIST', ['ID', 'TAG', 'CATEGORY', 'DISC', 'TASK TYPE', 'ASSET DESCRIPTION',
                                  'RESP COMPANY', 'SUB SYSTEM', 'SUB ID', 'CLOSING DATE', 'STATE',
                                  'M&P', 'EIT', 'RFC', 'PRIORITY', 'TEST FORM', 'USER NOTES'],
                 di, None, wbst, is_done_itr, None)
    out = save_with_retry_path(wb, OUT_XLSX)
    patch_codenames(out)
    xlsm = make_xlsm(out)
    if xlsm:
        if os.path.abspath(xlsm) != os.path.abspath(out):
            try:
                os.remove(out)
            except OSError:
                pass
        out = xlsm
    else:
        if os.path.exists(OUT_XLSM) and os.path.abspath(OUT_XLSM) != os.path.abspath(out):
            try:
                os.remove(OUT_XLSM)
            except OSError:
                pass
    print(f'DONE -> {out}')


def save_with_retry_path(wb, path):
    try:
        wb.save(path)
        return path
    except PermissionError:
        pass
    alt = os.path.join(TMP, os.path.basename(path))
    wb.save(alt)
    for attempt in range(15):
        try:
            shutil.copy2(alt, path)
            return path
        except PermissionError:
            print(f'  [wait] {os.path.basename(path)} is open in Excel - close it ... ({attempt + 1}/15)')
            time.sleep(2)
    stamp = datetime.datetime.now().strftime('%H%M%S')
    final = path.replace('.xlsx', f' {stamp}.xlsx')
    shutil.copy2(alt, final)
    return final


if __name__ == '__main__':
    main()
