"""
PS5 Master Tracker - Professional Update (openpyxl)
- Preserves ALL existing data (only updates specific cells)
- Professional color scheme
"""
import os, datetime, shutil
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

BASE = r"C:\Users\mylap\OneDrive\Desktop\dashboard"
TMP = os.path.join(BASE, "_temp")
MASTER = os.path.join(BASE, "PS5 Master tracker EIT Combined.xlsx")

# Professional colors
NAVY      = 'FF003366'
BLUE      = 'FF4472C4'
LBLUE     = 'FFD9E2F3'
ORANGE    = 'FFED7D31'
WHITE     = 'FFFFFFFF'
LGREEN    = 'FFC6EFCE'
LRED      = 'FFFFC7CE'
LYELLOW   = 'FFFFEB9C'
LIGHT_GRAY = 'FFF2F2F2'
DARK_GRAY = 'FF404040'

def n(v):
    if pd.isna(v): return ''
    return str(v).strip()

def nd(v):
    if pd.isna(v): return ''
    if isinstance(v, datetime.datetime): return v.strftime('%d-%b-%y')
    s = str(v).strip()
    if s and len(s) >= 8 and s[2] == '/':
        try: return datetime.datetime.strptime(s, '%d/%m/%Y').strftime('%d-%b-%y')
        except: pass
    return s

def find_file(patterns):
    for folder in (BASE, TMP):
        if not os.path.isdir(folder): continue
        for f in os.listdir(folder):
            low = f.lower()
            if f.startswith('~$'): continue
            if low.endswith('.xlsx') and all(p.lower() in low for p in patterns):
                return os.path.join(folder, f)
    return None

def prof_header(ws, row, cols):
    """Professional dark navy header"""
    for c in range(1, cols+1):
        cell = ws.cell(row, c)
        cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type='solid')
        cell.font = Font(bold=True, color=WHITE, size=10)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(
            bottom=Side(style='thin', color=WHITE),
            left=Side(style='thin', color=WHITE),
        )

def prof_alt_row(ws, row, cols, even):
    """Alternating row colors"""
    fill = 'FFFFFFFF' if even else LIGHT_GRAY
    for c in range(1, cols+1):
        cell = ws.cell(row, c)
        if cell.value is not None and not cell.fill or cell.fill.start_color.rgb in ('00000000', '0'):
            cell.fill = PatternFill(start_color=fill, end_color=fill, fill_type='solid')

disc_short = {'Electrical': 'E', 'Instrument': 'I', 'Telecom': 'T', 'E': 'E', 'I': 'I', 'T': 'T'}

print("=" * 60)
print("  PS5 MASTER TRACKER - PROFESSIONAL UPDATE")
print("=" * 60)

# Locate files
print("\n[0] Locating files...")
ov_path    = find_file(['ovTasks'])
dpr_path   = find_file(['PS-5 COMPLETIONS DPR SUMMERY'])
punch_path = find_file(['PUNCH LIST REGISTER'])
insp_path  = find_file(['INSPECTION REGISTER'])

for name, p in [('ovTasks', ov_path), ('DPR', dpr_path), ('Punch', punch_path), ('RFI', insp_path)]:
    print(f"  {name:15s}: {p or 'NOT FOUND'}")

if not all([ov_path, dpr_path, punch_path, insp_path]):
    print("  ERROR"); exit(1)

print("\n[1] Reading source data...")
ov = pd.read_excel(ov_path, header=None)
dpr = pd.read_excel(dpr_path, sheet_name='DETAILED ITR LIST', header=None)
pw = pd.read_excel(punch_path, header=None, skiprows=6)
rfi = pd.read_excel(insp_path, header=None)
print(f"  ovTasks={len(ov)}, DPR={len(dpr)}, Punch={len(pw)}, RFI={len(rfi)}")

# Punch data
punch_raw = {}
for idx, (_, r) in enumerate(pw.iterrows()):
    tag = n(r.iloc[0]) if len(r) > 0 else ''
    if not tag or tag.startswith('['): continue
    punch_raw[f'P{idx}'] = {
        'tag': tag,
        'desc': n(r.iloc[5]) if len(r) > 5 else '',
        'status': n(r.iloc[6]) if len(r) > 6 else '',
        'category': n(r.iloc[9]) if len(r) > 9 else '',
    }

# Build merged task data
merged = dpr.copy()
merged.columns = range(merged.shape[1])
for i in range(max(merged.shape[1], ov.shape[1])):
    if i >= merged.shape[1]: merged[i] = ''
    if i >= ov.shape[1]: ov[i] = ''

# Merge ovTasks state
ov_state = {}
for _, r in ov.iterrows():
    tid = n(r.iloc[0])
    if tid: ov_state[tid] = n(r.iloc[28])
for idx in range(len(merged)):
    tid = n(merged.iloc[idx, 0])
    if tid and tid in ov_state:
        merged.iloc[idx, 28] = ov_state[tid]

task_state = {}
task_data = {}
asset_task = {}
for _, r in merged.iterrows():
    tid = n(r.iloc[0])
    if tid:
        task_state[tid] = n(r.iloc[28])
        task_data[tid] = r
        at = n(r.iloc[1])
        if at and at not in asset_task:
            asset_task[at] = tid
print(f"  Tasks: {len(task_state)}")

# RFI lookup
rfi_by_asset = {}
for _, r in rfi.iterrows():
    tag = n(r.iloc[0]) if len(r) > 0 else ''
    if tag:
        rfi_by_asset.setdefault(tag, []).append({
            'rfi': n(r.iloc[0]), 'date': nd(r.iloc[2]) if len(r) > 2 else '',
            'status': n(r.iloc[3]) if len(r) > 3 else '',
        })

punch_by_asset = {}
for p in punch_raw.values():
    punch_by_asset.setdefault(p['tag'], []).append(p)

def get_rfi(tag):
    for r in rfi_by_asset.get(tag, []):
        return r.get('status',''), r.get('date',''), r.get('rfi','')
    return '','',''

# Open master
print("\n[2] Opening master...")
try:
    wb = openpyxl.load_workbook(MASTER)
except PermissionError:
    print("  ERROR: File locked! Close Excel first."); exit(1)
print(f"  Sheets: {wb.sheetnames}")

# ===== SHEET 4: TRACE HEATING =====
print("\n  --- TH Cable Schedule ---")
ws4 = wb['Trace Heating - Cable Schedule']
n_rfi = n_punch = 0
for r in range(3, ws4.max_row + 1):
    cable = n(ws4.cell(r, 2).value)
    equip = n(ws4.cell(r, 10).value) or cable
    if not cable: continue
    st, dt, no = get_rfi(equip)
    if no and not n(ws4.cell(r, 31).value):
        ws4.cell(r, 31).value = no; n_rfi += 1
    if dt and not n(ws4.cell(r, 32).value): ws4.cell(r, 32).value = dt
    if st and not n(ws4.cell(r, 33).value): ws4.cell(r, 33).value = st
    for p in punch_by_asset.get(equip, []):
        if p['status'].lower() != 'closed' and not n(ws4.cell(r, 39).value):
            ws4.cell(r, 39).value = p['desc'][:100]; n_punch += 1
print(f"  RFI={n_rfi}, Punch={n_punch}")

# ===== SHEET 5: ITR TRACKER =====
print("\n  --- ITR Tracker ---")
ws5 = wb['ITR Tracker']
cpp = merged[merged.iloc[:, 13].astype(str).str.strip() == 'CPP AGI'].copy()
cpp['disc'] = cpp.iloc[:, 10].astype(str).str.strip().map(lambda x: disc_short.get(x, x[:1] if x else ''))
cpp_eit = cpp[cpp['disc'].isin(['E', 'I', 'T'])].copy()

total = len(cpp_eit)
e_cnt = len(cpp_eit[cpp_eit['disc'] == 'E'])
i_cnt = len(cpp_eit[cpp_eit['disc'] == 'I'])
t_cnt = len(cpp_eit[cpp_eit['disc'] == 'T'])
closed_cnt = int((cpp_eit.iloc[:, 28].astype(str).str.strip() == 'Closed').sum())
open_cnt = total - closed_cnt

ws5.cell(3, 2).value = total
ws5.cell(4, 2).value = e_cnt
ws5.cell(5, 2).value = i_cnt
ws5.cell(6, 2).value = t_cnt
ws5.cell(7, 2).value = closed_cnt
ws5.cell(9, 2).value = open_cnt

# Color data rows
u = 0
for r in range(12, ws5.max_row + 1):
    tid = n(ws5.cell(r, 1).value) or asset_task.get(n(ws5.cell(r, 2).value), '')
    if tid and tid in task_state:
        if not n(ws5.cell(r, 1).value): ws5.cell(r, 1).value = tid
        st = task_state[tid]
        fill = PatternFill(start_color=LGREEN, end_color=LGREEN, fill_type='solid') if st.lower() == 'closed' else PatternFill(start_color=LRED, end_color=LRED, fill_type='solid')
        for c in range(1, 13): ws5.cell(r, c).fill = fill
        u += 1
print(f"  Colored: {u} rows")

# ===== SHEET 6: BACKLOG =====
print("\n  --- Backlog ---")
ws6 = wb['ITR Backlog (E+I+T)']
u = 0
for r in range(4, ws6.max_row + 1):
    tid = n(ws6.cell(r, 2).value) or asset_task.get(n(ws6.cell(r, 3).value), '')
    if tid and tid in task_state:
        if not n(ws6.cell(r, 2).value): ws6.cell(r, 2).value = tid
        st = task_state[tid]
        if n(ws6.cell(r, 14).value) != st: ws6.cell(r, 14).value = st
        fill = PatternFill(start_color=LGREEN, end_color=LGREEN, fill_type='solid') if st.lower() == 'closed' else PatternFill(start_color=LRED, end_color=LRED, fill_type='solid')
        for c in range(1, 21): ws6.cell(r, c).fill = fill
        u += 1
print(f"  Synced: {u} rows")

# ===== SHEET 7: CPP AGI MASTER =====
print("\n  --- CPP AGI Master ---")
ws7 = wb['CPP AGI ITR Master']
for disc, rn in [('E',2),('I',3),('T',4)]:
    sub = cpp_eit[cpp_eit['disc'] == disc]
    t = len(sub)
    c = int((sub.iloc[:,28].astype(str).str.strip() == 'Closed').sum())
    ws7.cell(rn, 3).value = t
    ws7.cell(rn, 5).value = c
    ws7.cell(rn, 7).value = f"{round(c/t*100) if t else 0}%"

u = 0
for r in range(8, ws7.max_row + 1):
    tid = n(ws7.cell(r, 2).value) or asset_task.get(n(ws7.cell(r, 3).value), '')
    if tid and tid in task_state:
        ws7.cell(r, 9).value = task_state[tid]
        if task_state[tid].lower() == 'closed':
            for c in range(1, 15): ws7.cell(r, c).fill = PatternFill(start_color=LGREEN, end_color=LGREEN, fill_type='solid')
        u += 1
print(f"  Synced: {u} rows")

# ===== PUNCH SHEETS =====
print("\n  --- Punch Sheets ---")
cpp_nc = merged[(merged.iloc[:,13].astype(str).str.strip() == 'CPP AGI') &
                (merged.iloc[:,28].astype(str).str.strip() != 'Closed')]
asset_nc = {}
for _, r in cpp_nc.iterrows():
    at = n(r.iloc[1])
    if at: asset_nc.setdefault(at, []).append({
        'tid': n(r.iloc[0]), 'desc': n(r.iloc[2])[:80], 'state': n(r.iloc[28])
    })

def punch_match(status_set):
    m, seen = [], set()
    for plid, p in punch_raw.items():
        if p['status'].lower() not in status_set: continue
        for t in asset_nc.get(p['tag'], []):
            key = t['tid'] + plid
            if key not in seen: seen.add(key); m.append(t)
    return m

for sname, filt, col in [('Closed Punch Tasks', ('closed','cleared','signed off'), LGREEN),
                          ('Open Punch Tasks', ('open',), None)]:
    wss = wb[sname]
    matches = punch_match(filt)
    for i, m in enumerate(matches, 4):
        wss.cell(i, 1).value = m['tid']
        wss.cell(i, 2).value = m.get('asset','')
        wss.cell(i, 3).value = m['desc']
        wss.cell(i, 4).value = m['state']
        if col:
            for c in range(1, 5): wss.cell(i, c).fill = PatternFill(start_color=col, end_color=col, fill_type='solid')
    print(f"  {sname}: {len(matches)}")

# ===== SUMMARY SHEETS =====
print("\n  --- Summary Sheets ---")

# Build from ovTasks (CPP AGI + E/I/T), unique subsystems, count unique Task IDs
ov_cpp = ov[ov.iloc[:, 13].astype(str).str.strip() == 'CPP AGI']
ov_eit = ov_cpp[ov_cpp.iloc[:, 8].astype(str).str.strip().isin(['E', 'I', 'T'])].copy()
ov_eit['subsystem'] = ov_eit.iloc[:, 21].astype(str).str.strip()
ov_eit['is_closed'] = ov_eit.iloc[:, 28].astype(str).str.strip() == 'Closed'
ov_eit = ov_eit[ov_eit['subsystem'] != 'nan']

# ===== Subsystem Summary (E+I+T) =====
# One row per subsystem with E&I&T classification column
sname = 'Subsystem Summary (E+I+T)'
hdrs = ['Subsystem', 'E&I&T', 'Total', 'Open', 'Closed', 'Close %']
col_w = [60, 10, 10, 10, 12, 10]
if sname in wb.sheetnames:
    wss = wb[sname]
    for r in range(wss.max_row, 0, -1): wss.delete_rows(r)
else:
    wss = wb.create_sheet(sname)

rows_data = []
for s, g in ov_eit.groupby('subsystem'):
    u = g.drop_duplicates(subset=0)
    t = len(u); c = int(u['is_closed'].sum())
    discs = '/'.join(sorted(g.iloc[:, 8].astype(str).str.strip().unique()))
    rows_data.append({'s': s, 'eit': discs, 't': t, 'c': c, 'o': t - c,
                      'p': round(c / t * 100, 1) if t else 0})
rows_data.sort(key=lambda x: x['s'])

ncols = len(hdrs)
for ci, h in enumerate(hdrs, 1):
    cell = wss.cell(1, ci)
    cell.value = h
    cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type='solid')
    cell.font = Font(bold=True, color=WHITE, size=11)
    cell.alignment = Alignment(horizontal='center', vertical='center')
for ci, w in enumerate(col_w, 1):
    wss.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w

rn = 2
for rd in rows_data:
    wss.cell(rn, 1).value = rd['s']
    wss.cell(rn, 2).value = rd['eit']
    wss.cell(rn, 3).value = rd['t']
    wss.cell(rn, 4).value = rd['o']
    wss.cell(rn, 5).value = rd['c']
    pct_cell = wss.cell(rn, 6)
    pct_cell.value = f"{rd['p']}%"
    row_fill = None
    if rd['p'] >= 100 and rd['t'] > 0:
        row_fill = PatternFill(start_color=LGREEN, end_color=LGREEN, fill_type='solid')
    elif rd['t'] > 0 and rd['c'] == 0:
        row_fill = PatternFill(start_color=LRED, end_color=LRED, fill_type='solid')
    for c2 in range(1, ncols + 1):
        if row_fill is not None:
            cell2 = wss.cell(rn, c2)
            if cell2.fill.start_color.rgb in ('00000000', '0'):
                cell2.fill = row_fill
    rn += 1
print(f"  {sname}: {len(rows_data)} rows")

# ===== Subsystem Summary (Combined) =====
# Columns per discipline: E, I, T with Total/Open/Closed each
sname = 'Subsystem Summary (Combined)'
hdrs = ['Subsystem',
        'E Total', 'E Open', 'E Closed', 'E Close %',
        'I Total', 'I Open', 'I Closed', 'I Close %',
        'T Total', 'T Open', 'T Closed', 'T Close %']
col_w = [60] + [10, 10, 12, 10] * 3
if sname in wb.sheetnames:
    wss = wb[sname]
    for r in range(wss.max_row, 0, -1): wss.delete_rows(r)
else:
    wss = wb.create_sheet(sname)

# Build per-discipline counts per subsystem
subs_all = sorted(ov_eit['subsystem'].unique())
rows_data = []
for s in subs_all:
    g = ov_eit[ov_eit['subsystem'] == s]
    row = {'s': s}
    for d in ['E', 'I', 'T']:
        gd = g[g.iloc[:, 8].astype(str).str.strip() == d]
        u = gd.drop_duplicates(subset=0)
        t = len(u); c = int(u['is_closed'].sum())
        row[d] = {'t': t, 'c': c, 'o': t - c, 'p': round(c / t * 100, 1) if t else 0}
    rows_data.append(row)

ncols = len(hdrs)
for ci, h in enumerate(hdrs, 1):
    cell = wss.cell(1, ci)
    cell.value = h
    cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type='solid')
    cell.font = Font(bold=True, color=WHITE, size=11)
    cell.alignment = Alignment(horizontal='center', vertical='center')
for ci, w in enumerate(col_w, 1):
    wss.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w

rn = 2
for rd in rows_data:
    wss.cell(rn, 1).value = rd['s']
    ci = 2
    for d in ['E', 'I', 'T']:
        v = rd[d]
        wss.cell(rn, ci).value = v['t']
        wss.cell(rn, ci + 1).value = v['o']
        wss.cell(rn, ci + 2).value = v['c']
        pct = wss.cell(rn, ci + 3)
        pct.value = f"{v['p']}%"
        pct.fill = PatternFill(start_color=LGREEN, end_color=LGREEN, fill_type='solid') if v['p'] >= 100 and v['t'] > 0 else PatternFill(start_color=LRED, end_color=LRED, fill_type='solid') if v['t'] > 0 and v['c'] == 0 else PatternFill()
        ci += 4
    rn += 1
print(f"  {sname}: {len(rows_data)} rows")

# ===== Punch Summary by Subsystem =====
# From Punch Register (PS5 scope): count Punch A/B/C per Systemization - Subsystem (Summary)
sname = 'Punch Summary by Subsystem'
hdrs = ['Systemization - Subsystem (Summary)', 'Punch A', 'Punch B', 'Punch C', 'Total Punch']
col_w = [70, 12, 12, 12, 12]
if sname in wb.sheetnames:
    wss = wb[sname]
    for r in range(wss.max_row, 0, -1): wss.delete_rows(r)
else:
    wss = wb.create_sheet(sname)

# Build asset tag -> subsystem map from merged task data
tag_to_sub = {}
for _, r in merged.iterrows():
    at = n(r.iloc[1])
    sub = n(r.iloc[21])
    if at and sub and at not in tag_to_sub:
        tag_to_sub[at] = sub

# Parse punch register rows: col 7 = TAG No., col 9 = Cat., col 4 = CPP SCOPE
punch_subs = {}
for _, r in pw.iterrows():
    tag = n(r.iloc[7]) if len(r) > 7 else ''
    scope = n(r.iloc[4]) if len(r) > 4 else ''
    cat = n(r.iloc[9]) if len(r) > 9 else ''
    if not tag or scope != 'PS5': continue
    sub = tag_to_sub.get(tag)
    if not sub: continue
    cat = cat.upper()
    if cat not in ('A', 'B', 'C'): continue
    d = punch_subs.setdefault(sub, {'A': 0, 'B': 0, 'C': 0})
    d[cat] += 1

ncols = len(hdrs)
for ci, h in enumerate(hdrs, 1):
    cell = wss.cell(1, ci)
    cell.value = h
    cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type='solid')
    cell.font = Font(bold=True, color=WHITE, size=11)
    cell.alignment = Alignment(horizontal='center', vertical='center')
for ci, w in enumerate(col_w, 1):
    wss.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w

rn = 2
for sub in sorted(punch_subs.keys()):
    d = punch_subs[sub]
    wss.cell(rn, 1).value = sub
    wss.cell(rn, 2).value = d['A']
    wss.cell(rn, 3).value = d['B']
    wss.cell(rn, 4).value = d['C']
    wss.cell(rn, 5).value = d['A'] + d['B'] + d['C']
    rn += 1
print(f"  {sname}: {len(punch_subs)} rows")

# Save
print("\n[3] Saving...")
try:
    wb.save(MASTER)
    print("  DONE!")
except Exception as e:
    print(f"  ERROR: {e}")
    tmp = os.path.join(TMP, '_master_backup.xlsx')
    wb.save(tmp)
    shutil.copy2(tmp, MASTER)
    print("  Saved via backup")

print("=" * 60)
