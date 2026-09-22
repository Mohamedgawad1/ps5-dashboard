#!/usr/bin/env python3
"""Generate the CMT Pre-Commissioning Dashboard HTML from Excel data"""
import openpyxl, json, os, sys
from collections import Counter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, 'index.html')
OUTPUT_FILE2 = r'C:\Users\mylap\OneDrive\Desktop\dashboard\CMT_DASHBOARD\index.html'

# Auto-detect the Excel file in the script directory
import glob
xlsx_files = glob.glob(os.path.join(BASE_DIR, 'CMT BASED PRE COMMISSIONING PROGRESS STATUS - PS5*.xlsx'))
# Prefer exact name without suffix, then latest by name
exact = [f for f in xlsx_files if ' (' not in os.path.basename(f)]
if exact:
    EXCEL_PATH = exact[0]
else:
    EXCEL_PATH = sorted(xlsx_files)[-1]  # highest suffix number
print(f'Using: {os.path.basename(EXCEL_PATH)}')

wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

milestones = ['MILESTONE A', 'MILESTONE B', 'MILESTONE C', 'MILESTONE D', 'MILESTONE E', 'MILESTONE F', 'MILESTONE G', 'MILESTONE H', 'MILESTONE I']
disc_list = ['B - Building', 'H - HVAC', 'M - Mechanical', 'P - Piping & Vessel', 'S - Safety', 'E - Electrical', 'I - Instrumentation', 'T - Telecom']

def si(v):
    try: return int(float(str(v))) if str(v).strip() else 0
    except: return 0

def sp(v):
    try:
        val = float(str(v))
        return round(val * 100, 1) if val <= 1 else round(val, 1)
    except: return 0

def parse_section(rows, start):
    res = {}
    i = start
    while i < len(rows):
        row = rows[i]
        dn = row[3].strip() if len(row) > 3 else ''
        if dn in disc_list:
            d = {'t': si(row[4]), 'c': si(row[5]), 'p': si(row[6]), 'pc': sp(row[7]), 'ms': {}}
            for mi, ms in enumerate(milestones):
                base = 9 + mi * 4
                if base + 3 < len(row):
                    d['ms'][ms] = {'t': si(row[base]), 'c': si(row[base+1]), 'p': si(row[base+2]), 'pc': sp(row[base+3])}
            res[dn] = d
        elif 'Grand Total' in dn:
            d = {'t': si(row[4]), 'c': si(row[5]), 'p': si(row[6]), 'pc': sp(row[7]), 'ms': {}}
            for mi, ms in enumerate(milestones):
                base = 9 + mi * 4
                if base + 3 < len(row):
                    d['ms'][ms] = {'t': si(row[base]), 'c': si(row[base+1]), 'p': si(row[base+2]), 'pc': sp(row[base+3])}
            res['Grand Total'] = d
            return res
        i += 1
    return res

def detect_companies(rows):
    """Scan sheet for company sections dynamically (new file format)."""
    companies = []
    seen = set()
    for i, row in enumerate(rows):
        col5 = row[4].strip() if len(row) > 4 else ''
        if col5 != 'TOTAL TASKS':
            continue
        # Company name is in col 2 of this header row or next data row
        name = row[1].strip() if len(row) > 1 else ''
        if not name and i + 1 < len(rows):
            name = rows[i + 1][1].strip() if len(rows[i + 1]) > 1 else ''
        # Clean up name
        name = name.replace(':', '').strip()
        if not name:
            continue
        if name.lower().startswith('ps5') or name == 'Overall':
            display = 'PS5 (All)'
        else:
            display = name.strip()
        if display in seen:
            continue
        seen.add(display)
        # Data starts at i+1 (next row after header)
        companies.append((display, i + 1))
    return companies

ws = wb['PS5 PROGRESS BY RESP. CPY ']
rows = []
for r in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
    rows.append([str(v) if v is not None else '' for v in r])

company_sections = detect_companies(rows)
overall_data = {}
for name, start_row in company_sections:
    overall_data[name] = parse_section(rows, start_row)

ws2 = wb['PS5 - SUBSYSTEMS PROGRESS']
subs = []
for row in ws2.iter_rows(min_row=8, max_row=ws2.max_row, values_only=True):
    vals = [str(v) if v is not None else '' for v in row]
    if len(vals) < 4 or not vals[0].strip() or not vals[1].strip() or not vals[1].startswith('PS5-'):
        continue
    # col 0=sys, 1=sub, 3=pri, 4=tt, 5=cl, 6=pd, 7=pc(decimal), 8=ap, 9=st
    status_raw = vals[9].strip() if len(vals) > 9 and vals[9].strip() else ''
    subs.append({
        'sys': vals[0].strip(), 'sub': vals[1].strip(), 'pri': vals[3].strip(),
        'tt': si(vals[4]), 'cl': si(vals[5]), 'pd': si(vals[6]), 'pc': sp(vals[7]),
        'ap': si(vals[8]) if len(vals) > 8 else 0,
        'st': status_raw
    })

ws3 = wb['EXPORTED PL']
cols = []
pl_rows = []
for r_idx, row in enumerate(ws3.iter_rows(min_row=1, max_row=ws3.max_row, values_only=True), 1):
    vals = [str(v) if v is not None else '' for v in row]
    if r_idx == 1: cols = vals
    else: pl_rows.append(vals)

ci = {c: i for i, c in enumerate(cols)}
a_punch_items = []
for p in pl_rows:
    try:
        cat = p[ci.get('Punchlist Category (Name)', 3)]
        st = p[ci.get('Status', 9)]
        if st.strip() != 'Closed':
            a_punch_items.append({
                'id': p[ci.get('Punchlist ID', 0)],
                'as': p[ci.get('Asset (Name/Tag)', 1)],
                'di': p[ci.get('Discipline (Name)', 4)],
                'co': p[ci.get('Scheduling - Responsible Company (Name)', 6)],
                'su': p[ci.get('Systemization - Subsystem (Summary)', 7)],
                'de': p[ci.get('Description', 5)],
                'ca': cat.strip(),
                'pr': p[ci.get('Priority', 14)]
            })
    except: pass

# ===================== EXPORTED TASKS =====================
ws4 = wb['EXPORTED TASKS']
task_cols = []
task_rows = []
for r_idx, row in enumerate(ws4.iter_rows(min_row=1, max_row=ws4.max_row, values_only=True), 1):
    vals = [str(v) if v is not None else '' for v in row]
    if r_idx == 1: task_cols = vals
    else: task_rows.append(vals)

tci = {c: i for i, c in enumerate(task_cols)}
exported_tasks = []
for p in task_rows:
    try:
        exported_tasks.append({
            'id': p[tci.get('Task ID', 0)],
            'as': p[tci.get('Asset - Tag', 1)],
            'de': p[tci.get('Description', 6)],
            'ca': p[tci.get('Category (Summary)', 7)],
            'di': p[tci.get('Discipline', 8)],
            'ds': p[tci.get('Discipline (Summary)', 10)],
            'co': p[tci.get('Responsible Company (Summary)', 13)],
            'su': p[tci.get('Systemization - Subsystem (Summary)', 21)],
            'pr': p[tci.get('Subsystem Priority', 29)],
            'st': p[tci.get('Task State', 28)]
        })
    except: pass

task_by_state = dict(Counter(x['st'] for x in exported_tasks).most_common())
task_by_comp = dict(Counter(x['co'] for x in exported_tasks).most_common())
task_by_disc = dict(Counter(x['di'] for x in exported_tasks).most_common())
task_by_pri = dict(Counter(x['pr'] for x in exported_tasks).most_common())

by_comp = dict(Counter(x['co'] for x in a_punch_items).most_common())
by_disc = dict(Counter(x['di'] for x in a_punch_items).most_common())

# ===================== HYDROTESTS =====================
ws5 = wb['HYDROTESTS SUMMARY']
ht_rows = []
for r in ws5.iter_rows(min_row=1, max_row=ws5.max_row, values_only=True):
    ht_rows.append([str(v) if v is not None else '' for v in r])

hydrotests = {}
# Summary line (row 10)
if len(ht_rows) >= 10:
    r10 = ht_rows[9]
    hydrotests['overall'] = r10[2].strip() if len(r10) > 2 else ''
    hydrotests['overall_pct'] = sp(r10[3]) if len(r10) > 3 else 0

# Test pack breakdown (rows 4-8)
ht_packs = []
for i in [3, 4, 5, 6, 7]:
    if i < len(ht_rows):
        r = ht_rows[i]
        name = r[1].strip() if len(r) > 1 else ''
        val = r[2].strip() if len(r) > 2 else ''
        pct = sp(r[3]) if len(r) > 3 else 0
        if name:
            ht_packs.append({'name': name, 'value': val, 'pct': pct})
hydrotests['packs'] = ht_packs

# Milestone breakdown (rows 15-23)
ht_ms = []
for i in range(14, min(23, len(ht_rows))):
    r = ht_rows[i]
    ms_name = r[1].strip() if len(r) > 1 else ''
    if not ms_name or ms_name == 'TOTAL':
        if ms_name == 'TOTAL':
            ht_ms.append({
                'name': 'TOTAL',
                'total': si(r[2]) if len(r) > 2 else 0,
                'closed': si(r[3]) if len(r) > 3 else 0,
                'pending': si(r[4]) if len(r) > 4 else 0,
                'pct': sp(r[5]) if len(r) > 5 else 0,
                'packs': {}
            })
        continue
    total = si(r[2]) if len(r) > 2 else 0
    closed = si(r[3]) if len(r) > 3 else 0
    pending = si(r[4]) if len(r) > 4 else 0
    pct = sp(r[5]) if len(r) > 5 else 0
    # Pack data: CLX01 (col 7-8), CLX02A (9-10), TMX03 (11-12), TMX05 (13-14), CLX02B (15-16)
    pack_cols = {'CLX01': (6, 7), 'CLX02A': (8, 9), 'TMX03': (10, 11), 'TMX05': (12, 13), 'CLX02B': (14, 15)}
    packs = {}
    for pk, (tc, cc) in pack_cols.items():
        t = si(r[tc]) if len(r) > tc else 0
        c = si(r[cc]) if len(r) > cc else 0
        packs[pk] = {'total': t, 'closed': c}
    ht_ms.append({'name': ms_name, 'total': total, 'closed': closed, 'pending': pending, 'pct': pct, 'packs': packs})
hydrotests['milestones'] = ht_ms

# ===================== GENERATE HTML =====================
DISC_COLORS = {'E - Electrical':'#9b59f6','I - Instrumentation':'#00e0c6','T - Telecom':'#ff4d8d','B - Building':'#ffb627','H - HVAC':'#3da8ff','M - Mechanical':'#ff8a3d','P - Piping & Vessel':'#7c8cff','S - Safety':'#ff4d8d'}
DISC_SHORT = {'B':'B - Building','H':'H - HVAC','M':'M - Mechanical','P':'P - Piping & Vessel','S':'S - Safety','E':'E - Electrical','I':'I - Instrumentation','T':'T - Telecom'}

def js_str(val):
    return json.dumps(val, ensure_ascii=False)

data_json = json.dumps({
    'milestones': milestones,
    'overall': overall_data,
    'subsystems': subs,
    'a_punches': {'total': len(a_punch_items), 'by_company': by_comp, 'by_discipline': by_disc, 'items': a_punch_items},
    'exp_tasks': {'total': len(exported_tasks), 'by_state': task_by_state, 'by_company': task_by_comp, 'by_discipline': task_by_disc, 'by_priority': task_by_pri, 'items': exported_tasks},
    'hydrotests': hydrotests,
    'comp_names': [n for n, _ in company_sections if n != 'PS5 (All)'],
    'disc_names': disc_list,
    'disc_short': DISC_SHORT,
}, ensure_ascii=False)

HTML = r"""<!DOCTYPE html>
<html lang="en" dir="ltr">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="900">
<title>PS5 — CMT Pre-Commissioning Progress Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --bgmain:#ffffff; --panel:#ffffff; --panel2:#f5f7fa;
  --accent:#ff4d8d; --gold:#ffb627; --teal:#00a88f; --blue:#3da8ff; --purple:#9b59f6;
  --text:#1a1a2e; --muted:#6b7280; --border:#e0e4ea;
}
*{box-sizing:border-box;font-family:'Segoe UI',Tahoma,Arial,sans-serif;margin:0;padding:0;}
body{
  background:#ffffff;
  color:var(--text);display:flex;
}
.sidebar{
  width:240px;min-height:100vh;background:#f8f9fb;color:var(--text);padding:24px 18px;
  position:sticky;top:0;height:100vh;overflow-y:auto;border-right:1px solid var(--border);
}
.sidebar h2{font-size:17px;margin-bottom:18px;display:flex;align-items:center;gap:8px;letter-spacing:.5px;color:var(--teal);}
.sidebar label{
  display:flex;align-items:center;gap:8px;padding:9px 8px;border-radius:8px;
  cursor:pointer;font-size:13px;margin-bottom:3px;transition:.15s;color:var(--muted);
}
.sidebar label:hover{background:rgba(0,0,0,.04);color:var(--text);}
.sidebar input{accent-color:var(--teal);width:16px;height:16px;}
.sidebar .grp{margin-top:20px;font-size:11px;opacity:.6;text-transform:uppercase;letter-spacing:1.5px;border-bottom:1px solid var(--border);padding-bottom:4px;color:var(--gold);}
.main{flex:1;padding:24px 32px;}
.header{
  background:linear-gradient(120deg,#e8edf5 0%, #dce1ea 100%);color:var(--text);padding:22px 30px;
  border-radius:16px;margin-bottom:24px;display:flex;justify-content:space-between;align-items:center;
  border:1px solid var(--border);position:relative;overflow:hidden;
}
.header::after{content:"";position:absolute;inset:0;background:radial-gradient(circle at 90% 10%, rgba(0,168,143,.08), transparent 60%);}
.header h1{font-size:24px;letter-spacing:.5px;}
.header .sub{font-size:12px;opacity:.7;margin-top:4px;color:var(--muted);}
.section{display:none;}
.section.active{display:block;}
.kpi-row{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:22px;}
.kpi{
  flex:1;min-width:170px;background:var(--panel);border-radius:14px;padding:20px;
  text-align:center;border:1px solid var(--border);border-top:4px solid var(--purple);
  position:relative;overflow:hidden;transition:transform .2s, border-color .2s;
}
.kpi:hover{transform:translateY(-4px);border-color:var(--teal);}
.kpi.gold{border-top-color:var(--gold);}
.kpi.teal{border-top-color:var(--teal);}
.kpi.pink{border-top-color:var(--accent);}
.kpi.blue{border-top-color:var(--blue);}
.kpi .icon{font-size:24px;margin-bottom:6px;}
.kpi .val{font-size:32px;font-weight:800;color:var(--text);letter-spacing:.5px;}
.kpi .lbl{font-size:12px;color:var(--muted);margin-top:4px;}
.chart-row{display:flex;gap:18px;flex-wrap:wrap;margin-bottom:22px;}
.chart-card{
  flex:1;min-width:340px;background:var(--panel);border-radius:14px;padding:20px;
  border:1px solid var(--border);
}
.chart-card h3{font-size:15px;color:var(--teal);margin-bottom:12px;font-weight:700;}
canvas{max-height:300px;}
.section-title{
  font-size:19px;color:var(--text);font-weight:800;margin:8px 0 16px;
  border-left:6px solid var(--accent);padding-left:12px;display:flex;align-items:center;gap:8px;
}
.progress-bar{height:10px;background:#e5e7eb;border-radius:6px;overflow:hidden;margin-top:8px;border:1px solid var(--border);}
.progress-fill{height:100%;background:linear-gradient(90deg,var(--blue),var(--teal));}
.footer{text-align:center;padding:16px;color:var(--muted);font-size:12px;}
table{color:var(--text);}
thead tr{background:var(--panel2) !important;}
tbody tr{border-bottom:1px solid var(--border) !important;}

.filter-bar{display:flex;gap:12px;flex-wrap:wrap;align-items:center;margin-bottom:16px;}
.filter-bar select,.filter-bar input{
  background:var(--panel2);border:1px solid var(--border);color:var(--text);padding:8px 12px;
  border-radius:8px;font-size:13px;min-width:160px;
}
.filter-bar select option{background:#fff;color:var(--text);}
.filter-bar label{font-size:12px;color:var(--muted);}
.btn-export{
  background:linear-gradient(120deg,#1d6f42,#2e9e5b);color:#fff;border:none;
  padding:10px 20px;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;
  display:inline-flex;align-items:center;gap:8px;box-shadow:0 4px 14px rgba(46,158,91,.35);transition:.15s;
}
.btn-export:hover{transform:translateY(-1px);box-shadow:0 6px 18px rgba(46,158,91,.5);}
.btn-go{
  background:linear-gradient(120deg,#3da8ff,#2b7fd4);color:#fff;border:none;
  padding:8px 20px;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;
  transition:.15s;
}
.btn-go:hover{transform:translateY(-1px);box-shadow:0 4px 14px rgba(61,168,255,.4);}
.table-wrap{background:#fff;border-radius:10px;overflow:auto;max-height:540px;border:1px solid #c9c9c9;margin-bottom:18px;}
table.data-table{width:100%;border-collapse:collapse;font-size:12px;color:#111;}
table.data-table th{
  background:#ED7D31;color:#000;font-weight:800;padding:8px 6px;text-align:center;
  border:1px solid #9c4a14;position:sticky;top:0;z-index:2;
}
table.data-table td{padding:5px 6px;border:1px solid #d4d4d4;text-align:center;}
table.data-table td.left{text-align:left;}
table.data-table tr.total td{background:#404040 !important;color:#fff;font-weight:800;border:1px solid #222;}

.badge{
  display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;
  background:rgba(0,224,198,.15);color:var(--teal);border:1px solid rgba(0,224,198,.25);
}
.badge-open{background:rgba(255,77,141,.15);color:var(--accent);border:1px solid rgba(255,77,141,.25);}

.tab-bar{display:flex;gap:4px;margin-bottom:20px;flex-wrap:wrap;}
.tab-btn{
  padding:10px 24px;border-radius:10px 10px 0 0;font-size:14px;font-weight:700;
  cursor:pointer;border:1px solid var(--border);border-bottom:none;
  background:var(--panel2);color:var(--muted);transition:.15s;
}
.tab-btn.active{background:var(--panel);color:var(--teal);border-color:var(--teal);}
.tab-btn:hover{color:var(--text);}

.pagination{display:flex;justify-content:center;gap:6px;margin-top:12px;flex-wrap:wrap;}
.pagination button{
  padding:6px 12px;border-radius:6px;border:1px solid var(--border);background:var(--panel2);
  color:var(--text);cursor:pointer;font-size:12px;
}
.pagination button.active{background:var(--teal);color:#000;border-color:var(--teal);font-weight:700;}
</style>
</head>
<body>

<div class="sidebar">
  <h2>🛢️ PS5 CMT Dashboard</h2>
  <div class="grp">📊 Pages</div>
  <label><input type="checkbox" data-target="page-overall" checked> 1. Overall Progress</label>
  <label><input type="checkbox" data-target="page-apunch" checked> 2. Open A-Punches</label>
  <label><input type="checkbox" data-target="page-subsystem" checked> 3. Subsystem Progress</label>
  <label><input type="checkbox" data-target="page-tasks" checked> 4. Exported Tasks</label>
  <label><input type="checkbox" data-target="page-hydrotests" checked> 5. Hydrotests Summary</label>
</div>

<div class="main">
  <div class="header">
    <div>
      <h1>📊 PS5 — CMT Based Pre-Commissioning Progress Status</h1>
      <div class="sub">EACOP Project | Data extracted from CMT Progress Status - PS5</div>
    </div>
    <div style="font-size:38px;">📋</div>
  </div>

  <!-- ===================== TABS ===================== -->
  <div class="tab-bar">
    <div class="tab-btn active" onclick="switchTab('page-overall')">📈 1. Overall Progress</div>
    <div class="tab-btn" onclick="switchTab('page-apunch')">📌 2. Open A-Punches</div>
    <div class="tab-btn" onclick="switchTab('page-subsystem')">🏗️ 3. Subsystem Progress</div>
    <div class="tab-btn" onclick="switchTab('page-tasks')">📋 4. Exported Tasks</div>
    <div class="tab-btn" onclick="switchTab('page-hydrotests')">💧 5. Hydrotests Summary</div>
  </div>

  <!-- ===================== PAGE 1: OVERALL PROGRESS ===================== -->
  <div class="section active" id="page-overall">
    <div class="section-title">📈 Overall Completion Progress by Milestone</div>

    <div class="kpi-row" id="overallKpis"></div>

    <div class="filter-bar">
      <label>Company:</label>
      <select id="filterCompany" onchange="renderOverall()">""" + '\n        '.join([f'<option value="{n}">{n}</option>' if n != 'PS5 (All)' else '<option value="PS5 (All)">PS5 (All Companies)</option>' for n, _ in company_sections]) + r"""
      </select>
      <label>Discipline:</label>
      <select id="filterDiscOverall" onchange="renderOverall()">
        <option value="ALL">ALL Disciplines</option>
      </select>
      <button class="btn-go" onclick="renderOverall()">🔍 Go</button>
    </div>

    <div class="chart-row">
      <div class="chart-card"><h3>Task Completion by Milestone</h3><canvas id="chartOverallMs"></canvas></div>
      <div class="chart-card"><h3>Overall Completion % by Discipline</h3><canvas id="chartOverallDisc"></canvas></div>
    </div>

    <div class="chart-row">
      <div class="chart-card" style="flex:1.5;"><h3>Milestone-wise Closed vs Pending</h3><canvas id="chartOverallStacked"></canvas></div>
    </div>

    <div class="section-title" style="font-size:15px;border-left-color:var(--gold);">📋 Detailed Table by Discipline & Milestone</div>
    <div class="table-wrap" id="overallTableWrap"></div>
    <div style="display:flex;justify-content:flex-end;margin-bottom:12px;">
      <button class="btn-export" onclick="exportOverallExcel()">⬇️ Download Overall Progress Excel</button>
    </div>
  </div>

  <!-- ===================== PAGE 2: OPEN A-PUNCHES ===================== -->
  <div class="section" id="page-apunch">
    <div class="section-title">📌 Status of Open Punches (A / B / C)</div>
    <div class="kpi-row" id="apunchKpis"></div>

    <div class="filter-bar">
      <label>Category:</label>
      <select id="filterApCat" onchange="renderAPunch()">
        <option value="ALL">All Categories</option>
        <option value="A">A - Punch</option>
        <option value="B">B - Punch</option>
        <option value="C">C - Punch</option>
      </select>
      <label>Company:</label>
      <select id="filterApComp" onchange="renderAPunch()">
        <option value="ALL">All Companies</option>
      </select>
      <label>Discipline:</label>
      <select id="filterApDisc" onchange="renderAPunch()">
        <option value="ALL">All Disciplines</option>
      </select>
      <label>Search:</label>
      <input id="searchAPunch" type="text" placeholder="Search PL ID / Asset..." oninput="renderAPunch()" style="background:var(--panel2);border:1px solid var(--border);color:var(--text);padding:8px 12px;border-radius:8px;font-size:13px;min-width:200px;">
      <button class="btn-go" onclick="renderAPunch()">🔍 Go</button>
    </div>

    <div class="chart-row">
      <div class="chart-card"><h3>Open Punches by Category (A / B / C)</h3><canvas id="chartApCat"></canvas></div>
      <div class="chart-card"><h3>Open Punches by Company</h3><canvas id="chartApComp"></canvas></div>
      <div class="chart-card"><h3>Open Punches by Discipline</h3><canvas id="chartApDisc"></canvas></div>
    </div>

    <div class="table-wrap" id="apunchTableWrap"></div>
    <div class="pagination" id="apunchPagination"></div>
  </div>

  <!-- ===================== PAGE 3: SUBSYSTEM PROGRESS ===================== -->
  <div class="section" id="page-subsystem">
    <div class="section-title">🏗️ Subsystem Progress by Milestone, Discipline & Company</div>
    <div class="kpi-row" id="subKpis"></div>

    <div class="filter-bar">
      <label>Milestone:</label>
      <select id="filterSubMs" onchange="renderSubsystem()">
        <option value="ALL">All Milestones</option>
      </select>
      <label>Discipline:</label>
      <select id="filterSubDisc" onchange="renderSubsystem()">
        <option value="ALL">All Disciplines</option>
      </select>
      <label>Company:</label>
      <select id="filterSubComp" onchange="renderSubsystem()">
        <option value="ALL">All Companies</option>
      </select>
      <label>Search:</label>
      <input id="searchSub" type="text" placeholder="Search subsystem..." oninput="renderSubsystem()" style="background:var(--panel2);border:1px solid var(--border);color:var(--text);padding:8px 12px;border-radius:8px;font-size:13px;min-width:200px;">
      <button class="btn-go" onclick="renderSubsystem()">🔍 Go</button>
    </div>

    <div class="chart-row">
      <div class="chart-card" style="flex:1.5;"><h3>Top 20 Subsystems — Task Completion %</h3><canvas id="chartSubTop"></canvas></div>
      <div class="chart-card"><h3>Subsystems by Milestone</h3><canvas id="chartSubMs"></canvas></div>
    </div>
    <div class="chart-row">
      <div class="chart-card"><h3>A-Punch Distribution by Subsystem</h3><canvas id="chartSubAp"></canvas></div>
      <div class="chart-card"><h3>Status Distribution</h3><canvas id="chartSubStatus"></canvas>
        <div style="display:flex;gap:16px;justify-content:center;margin-top:8px;font-size:12px;">
          <span><span style="display:inline-block;width:12px;height:12px;border-radius:50%;background:#2ecc71;vertical-align:middle;margin-right:4px;"></span>COMPLETED</span>
          <span><span style="display:inline-block;width:12px;height:12px;border-radius:50%;background:#ffb627;vertical-align:middle;margin-right:4px;"></span>In Progress</span>
          <span><span style="display:inline-block;width:12px;height:12px;border-radius:50%;background:#3da8ff;vertical-align:middle;margin-right:4px;"></span>Other</span>
        </div>
      </div>
    </div>

    <div class="section-title" style="font-size:15px;border-left-color:var(--teal);">📋 Subsystem Progress Table</div>
    <div style="display:flex;justify-content:flex-end;margin-bottom:12px;">
      <button class="btn-export" onclick="exportSubsystemExcel()">⬇️ Download Subsystem Excel</button>
    </div>
    <div class="table-wrap" id="subsystemTableWrap"></div>
    <div class="pagination" id="subPagination"></div>
  </div>

  <!-- ===================== PAGE 5: HYDROTESTS SUMMARY ===================== -->
  <div class="section" id="page-hydrotests">
    <div class="section-title">💧 Hydrotests Summary — PS5</div>
    <div class="kpi-row" id="htKpis"></div>

    <div class="filter-bar">
      <label>Test Pack:</label>
      <select id="filterHtPack" onchange="renderHydrotests()">
        <option value="ALL">All Test Packs</option>
        <option value="CLX01">CLX01 - Pipework Test Pack</option>
        <option value="CLX02A">CLX02A - Pipework Before Test</option>
        <option value="CLX02B">CLX02B - Pipework After Test</option>
        <option value="TMX03">TMX03 - Cleaning</option>
        <option value="TMX05">TMX05 - Pressure Test</option>
      </select>
      <label>Milestone:</label>
      <select id="filterHtMs" >
        <option value="ALL">All Milestones</option>
      </select>
      <button class="btn-go" onclick="renderHydrotests()">🔍 Go</button>
    </div>

    <div class="chart-row">
      <div class="chart-card"><h3>Hydrotest Completion % by Test Pack</h3><canvas id="chartHtPacks"></canvas></div>
      <div class="chart-card"><h3>Overall Hydrotest Progress by Milestone</h3><canvas id="chartHtMilestones"></canvas></div>
    </div>
    <div class="chart-row">
      <div class="chart-card" style="flex:1.5;"><h3>Hydrotest Tasks by Milestone (Closed vs Pending)</h3><canvas id="chartHtStacked"></canvas></div>
    </div>

    <div class="section-title" style="font-size:15px;border-left-color:var(--blue);">📋 Detailed Hydrotest Table by Milestone</div>
    <div style="display:flex;justify-content:flex-end;margin-bottom:12px;">
      <button class="btn-export" onclick="exportHydrotestsExcel()">⬇️ Download Hydrotests Excel</button>
    </div>
    <div class="table-wrap" id="htTableWrap"></div>
  </div>

  <!-- ===================== PAGE 4: EXPORTED TASKS ===================== -->
  <div class="section" id="page-tasks">
    <div class="section-title">📋 Exported Tasks — Full Task Registry</div>
    <div class="kpi-row" id="tasksKpis"></div>

    <div class="filter-bar">
      <label>Company:</label>
      <select id="filterTComp" onchange="renderTasks()">
        <option value="ALL">All Companies</option>
      </select>
      <label>Discipline:</label>
      <select id="filterTDisc" onchange="renderTasks()">
        <option value="ALL">All Disciplines</option>
      </select>
      <label>Milestone:</label>
      <select id="filterTPri" onchange="renderTasks()">
        <option value="ALL">All Milestones</option>
      </select>
      <label>State:</label>
      <select id="filterTState" onchange="renderTasks()">
        <option value="ALL">All States</option>
      </select>
      <label>Search:</label>
      <input id="searchTasks" type="text" placeholder="Search Task ID / Asset / Description..." oninput="renderTasks()" style="background:var(--panel2);border:1px solid var(--border);color:var(--text);padding:8px 12px;border-radius:8px;font-size:13px;min-width:200px;">
      <button class="btn-go" onclick="renderTasks()">🔍 Go</button>
    </div>

    <div class="chart-row">
      <div class="chart-card"><h3>Tasks by State</h3><canvas id="chartTState"></canvas></div>
      <div class="chart-card"><h3>Tasks by Company</h3><canvas id="chartTComp"></canvas></div>
    </div>
    <div class="chart-row">
      <div class="chart-card"><h3>Tasks by Discipline</h3><canvas id="chartTDisc"></canvas></div>
      <div class="chart-card"><h3>Tasks by Milestone</h3><canvas id="chartTPri"></canvas></div>
    </div>

    <div class="section-title" style="font-size:15px;border-left-color:var(--gold);">📋 Task Registry Table</div>
    <div style="display:flex;justify-content:flex-end;margin-bottom:12px;">
      <button class="btn-export" onclick="exportTasksExcel()">⬇️ Download Tasks Excel</button>
    </div>
    <div class="table-wrap" id="tasksTableWrap"></div>
    <div class="pagination" id="tasksPagination"></div>
  </div>

  <div class="footer">Auto-generated CMT Pre-Commissioning Dashboard — PS5 / EACOP Project</div>
</div>

<script>
Chart.register(ChartDataLabels);
Chart.defaults.font.family = "'Segoe UI', Tahoma, Arial, sans-serif";
Chart.defaults.color = '#6b7280';
Chart.defaults.borderColor = '#e0e4ea';

const DL_PIE = {display:true, color:'#1a1a2e', font:{weight:'bold', size:10}, formatter:(v,ctx)=>{const t=ctx.chart.getDatasetMeta(0).total||ctx.dataset.data.reduce((a,b)=>a+b,0);const p=t?Math.round(v/t*100):0;return v>0?v+' ('+p+'%)':''}};
const DL_BAR = {display:true, color:'#0b1120', backgroundColor:'rgba(255,255,255,.9)', borderRadius:4, padding:{top:2,bottom:2,left:4,right:4}, anchor:'end', align:'top', offset:4, font:{weight:'bold', size:9}, formatter:(v)=>v>0?v:''};
const DL_STACK = {display:true, color:'#1a1a2e', font:{weight:'bold', size:9}, formatter:(v)=>v>3?v:''};

const DATA = """ + data_json + r""";

const MILESTONES = DATA.milestones;
const OVERALL = DATA.overall;
const SUBS = DATA.subsystems;
const APUNCH = DATA.a_punches;
const ETASKS = DATA.exp_tasks;
const DISC_SHORT = DATA.disc_short;
const DISC_NAMES = DATA.disc_names;
const COMP_NAMES = DATA.comp_names;

const palette = ['#9b59f6','#3da8ff','#ff4d8d','#ffb627','#00e0c6','#7c8cff','#ff8a3d','#2ecc71'];
const DISC_COLORS = {'E - Electrical':'#9b59f6','I - Instrumentation':'#00e0c6','T - Telecom':'#ff4d8d','B - Building':'#ffb627','H - HVAC':'#3da8ff','M - Mechanical':'#ff8a3d','P - Piping & Vessel':'#7c8cff','S - Safety':'#ff4d8d'};
const MS_COLORS = ['#9b59f6','#3da8ff','#ff4d8d','#ffb627','#00e0c6','#7c8cff','#ff8a3d','#2ecc71','#e74c3c'];
const palette2 = ['#ff4d8d','#3da8ff','#ffb627','#00e0c6','#9b59f6','#7c8cff'];

document.querySelectorAll('.sidebar input[type=checkbox]').forEach(cb=>{
  cb.addEventListener('change', ()=>{
    const el = document.getElementById(cb.dataset.target);
    if(cb.checked) el.classList.add('active');
    else el.classList.remove('active');
  });
});

function switchTab(id){
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.querySelector(`.tab-btn[onclick*="${id}"]`).classList.add('active');
}

function pctBar(pct){
  const w = Math.min(pct,100);
  return `<div style="background:#d9d9d9;border-radius:4px;height:18px;min-width:60px;position:relative;">
    <div style="position:absolute;left:0;top:0;height:100%;width:${w}%;background:linear-gradient(90deg,#70ad47,#a9d18e);border-radius:4px;"></div>
    <div style="position:relative;font-size:11px;font-weight:700;line-height:18px;">${pct}%</div></div>`;
}

// ==================== PAGE 1: OVERALL ====================
function initFilters(){
  // Disc filter
  const sel = document.getElementById('filterDiscOverall');
  DISC_NAMES.forEach(d=>{ const o=document.createElement('option'); o.value=d; o.textContent=d; sel.appendChild(o); });
  // Apunch company
  const sel2 = document.getElementById('filterApComp');
  Object.keys(APUNCH.by_company).forEach(c=>{ const o=document.createElement('option'); o.value=c; o.textContent=c; sel2.appendChild(o); });
  const sel3 = document.getElementById('filterApDisc');
  Object.keys(APUNCH.by_discipline).forEach(d=>{ const o=document.createElement('option'); o.value=d; o.textContent=DISC_SHORT[d]||d; sel3.appendChild(o); });
  // Subsystem
  const sel4 = document.getElementById('filterSubMs');
  Object.keys(ETASKS.by_priority).sort().forEach(m=>{ const o=document.createElement('option'); o.value=m; o.textContent=m.replace('PS5 - ','M-'); sel4.appendChild(o); });
  const sel5 = document.getElementById('filterSubDisc');
  DISC_NAMES.forEach(d=>{ const o=document.createElement('option'); o.value=d; o.textContent=d; sel5.appendChild(o); });
  const sel6 = document.getElementById('filterSubComp');
  COMP_NAMES.forEach(c=>{ const o=document.createElement('option'); o.value=c; o.textContent=c; sel6.appendChild(o); });
  // Hydrotests milestone filter
  const selHt = document.getElementById('filterHtMs');
  MILESTONES.forEach(m=>{ const o=document.createElement('option'); o.value=m; o.textContent=m.replace('MILESTONE ','M-'); selHt.appendChild(o); });
}

function getOverallData(){
  const comp = document.getElementById('filterCompany').value;
  const disc = document.getElementById('filterDiscOverall').value;
  let src = OVERALL[comp] || OVERALL['PS5 (All)'];
  if(!src) return null;
  if(disc === 'ALL'){
    const gt = src['Grand Total'];
    return gt ? gt : {t:0,c:0,p:0,pc:0,ms:{}};
  }
  return src[disc] || {t:0,c:0,p:0,pc:0,ms:{}};
}

function getOverallDataByDisc(comp){
  return OVERALL[comp] || OVERALL['PS5 (All)'];
}

function renderOverall(){
  const comp = document.getElementById('filterCompany').value;
  const disc = document.getElementById('filterDiscOverall').value;
  const src = getOverallDataByDisc(comp);
  if(!src) return;

  // KPIs
  const gt = src['Grand Total'];
  document.getElementById('overallKpis').innerHTML = `
    <div class="kpi teal"><div class="icon">📋</div><div class="val">${gt ? gt.t : 0}</div><div class="lbl">Total Tasks (${comp})</div></div>
    <div class="kpi"><div class="icon">✅</div><div class="val">${gt ? gt.c : 0}</div><div class="lbl">Closed Tasks</div></div>
    <div class="kpi pink"><div class="icon">⏳</div><div class="val">${gt ? gt.p : 0}</div><div class="lbl">Pending Tasks</div></div>
    <div class="kpi gold"><div class="icon">📈</div><div class="val">${gt ? gt.pc : 0}%</div><div class="lbl">Completion Rate</div></div>
  `;

  // Milestone chart
  const msData = MILESTONES.map(m => {
    let t = 0, c = 0, p = 0, pc = 0;
    if(disc === 'ALL' && src['Grand Total'] && src['Grand Total'].ms[m]) {
      t = src['Grand Total'].ms[m].t; c = src['Grand Total'].ms[m].c; p = src['Grand Total'].ms[m].p; pc = src['Grand Total'].ms[m].pc;
    } else if(disc !== 'ALL' && src[disc] && src[disc].ms[m]) {
      t = src[disc].ms[m].t; c = src[disc].ms[m].c; p = src[disc].ms[m].p; pc = src[disc].ms[m].pc;
    } else if(disc === 'ALL') {
      // Sum all disciplines
      DISC_NAMES.forEach(dn => {
        if(src[dn] && src[dn].ms[m]) { t += src[dn].ms[m].t; c += src[dn].ms[m].c; p += src[dn].ms[m].p; }
      });
      pc = t ? Math.round(c/t*100*10)/10 : 0;
    }
    return {m: m.replace('MILESTONE ','M-'), t, c, p, pc};
  });

  const labels = msData.map(d=>d.m);
  //destroyChart('chartOverallMs');
  makeChart('chartOverallMs', {
    type:'bar',
    data:{ labels,
      datasets:[
        {label:'Closed', data:msData.map(d=>d.c), backgroundColor:'#00e0c6', datalabels:DL_BAR},
        {label:'Total', data:msData.map(d=>d.t), backgroundColor:'rgba(42,52,80,0.8)', datalabels:{display:false}}
      ]
    },
    options:{ plugins:{legend:{position:'bottom'}}, scales:{y:{beginAtZero:true}} }
  });

  // Disc completion chart
  const discData = DISC_NAMES.map(dn => {
    if(!src[dn]) return {dn, pc:0};
    return {dn, pc: src[dn].pc};
  });
  //destroyChart('chartOverallDisc');
  makeChart('chartOverallDisc', {
    type:'bar',
    data:{ labels: discData.map(d=>d.dn), datasets:[{label:'% Complete', data:discData.map(d=>d.pc), backgroundColor: DISC_NAMES.map(d=>DISC_COLORS[d]||'#9b59f6'), datalabels: DL_BAR}] },
    options:{ indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{beginAtZero:true,max:100}} }
  });

  // Stacked milestone
  //destroyChart('chartOverallStacked');
  makeChart('chartOverallStacked', {
    type:'bar',
    data:{ labels: labels,
      datasets: [
        {label:'Closed', data:msData.map(d=>d.c), backgroundColor:'#00e0c6', datalabels:DL_STACK},
        {label:'Pending', data:msData.map(d=>d.p), backgroundColor:'#ff4d8d', datalabels:DL_STACK}
      ]
    },
    options:{ plugins:{legend:{position:'bottom'}}, scales:{x:{stacked:true},y:{stacked:true,beginAtZero:true}} }
  });

  // Table
  let html = `<table class="data-table" style="min-width:1400px;"><thead><tr>
    <th style="text-align:left;">Discipline</th><th>Total</th><th>Closed</th><th>Pending</th><th>%</th>`;
  MILESTONES.forEach(m => { html += `<th colspan="4">${m.replace('MILESTONE ','M-')}</th>`; });
  html += `</tr><tr><th></th><th></th><th></th><th></th><th></th>`;
  MILESTONES.forEach(() => { html += `<th>T</th><th>C</th><th>P</th><th>%</th>`; });
  html += `</tr></thead><tbody>`;

  let discFilterList = DISC_NAMES;
  if(disc !== 'ALL') discFilterList = [disc];

  discFilterList.forEach(dn => {
    if(!src[dn]) return;
    const d = src[dn];
    html += `<tr><td class="left" style="font-weight:bold;">${dn}</td>
      <td>${d.t}</td><td style="color:#1d6f42;font-weight:700;">${d.c}</td><td style="color:#C00000;font-weight:700;">${d.p}</td>
      <td>${pctBar(d.pc)}</td>`;
    MILESTONES.forEach(m => {
      const ms = d.ms[m] || {};
      const pc = ms.pc ? ms.pc + '%' : '-';
      html += `<td>${ms.t||'-'}</td>
        <td style="${ms.c>0?'color:#1d6f42;font-weight:700;':''}">${ms.c||'-'}</td>
        <td style="${ms.p>0?'color:#C00000;font-weight:700;':''}">${ms.p||'-'}</td>
        <td>${pc}</td>`;
    });
    html += `</tr>`;
  });
  if(src['Grand Total']) {
    const g = src['Grand Total'];
    html += `<tr class="total"><td style="text-align:left;">GRAND TOTAL</td>
      <td>${g.t}</td><td>${g.c}</td><td>${g.p}</td><td>${pctBar(g.pc)}</td>`;
    MILESTONES.forEach(m => {
      const ms = g.ms[m] || {};
      const pc = ms.pc ? ms.pc + '%' : '-';
      html += `<td>${ms.t||'-'}</td><td>${ms.c||'-'}</td><td>${ms.p||'-'}</td><td>${pc}</td>`;
    });
    html += `</tr>`;
  }
  html += `</tbody></table>`;
  document.getElementById('overallTableWrap').innerHTML = html;
}

// ==================== PAGE 2: A-PUNCHES ====================
let apunchPage = 1;
const APUNCH_PER_PAGE = 25;

function renderAPunch(){
  const compF = document.getElementById('filterApComp').value;
  const discF = document.getElementById('filterApDisc').value;
  const catF = document.getElementById('filterApCat').value;
  const search = document.getElementById('searchAPunch').value.toLowerCase().trim();

  let filtered = APUNCH.items.filter(p => {
    if(catF !== 'ALL' && p.ca !== catF) return false;
    if(compF !== 'ALL' && p.co !== compF) return false;
    if(discF !== 'ALL' && p.di !== discF) return false;
    if(search) return p.id.toLowerCase().includes(search) || p.as.toLowerCase().includes(search) || p.de.toLowerCase().includes(search);
    return true;
  });

  // KPIs (from filtered)
  const apByCompany = {}; const apByDisc = {}; const apByCat = {};
  filtered.forEach(p => { 
    apByCompany[p.co] = (apByCompany[p.co]||0)+1; 
    apByDisc[p.di] = (apByDisc[p.di]||0)+1;
    apByCat[p.ca] = (apByCat[p.ca]||0)+1;
  });
  document.getElementById('apunchKpis').innerHTML = `
    <div class="kpi pink"><div class="icon">📌</div><div class="val">${filtered.length}</div><div class="lbl">Filtered Open Punches</div></div>
    <div class="kpi gold"><div class="icon">🔍</div><div class="val">${APUNCH.total}</div><div class="lbl">Total (Unfiltered)</div></div>
    <div class="kpi blue"><div class="icon">🏢</div><div class="val">${Object.keys(apByCompany).length}</div><div class="lbl">Companies</div></div>
    <div class="kpi teal"><div class="icon">📂</div><div class="val">${Object.keys(apByDisc).length}</div><div class="lbl">Disciplines</div></div>
  `;

  // Category chart (A/B/C)
  destroyChart('chartApCat');
  const catLabels = Object.keys(apByCat).sort();
  const catColors = {'A':'#ff4d8d','B':'#ffaa00','C':'#00b4d8'};
  makeChart('chartApCat', {
    type:'doughnut',
    data:{ labels:catLabels.map(c=>c+' - Punch'), datasets:[{ data:catLabels.map(k=>apByCat[k]), backgroundColor:catLabels.map(c=>catColors[c]||'#666'), datalabels:DL_PIE }] },
    options:{ plugins:{legend:{position:'bottom'}} }
  });

  // Charts (from filtered)
  destroyChart('chartApComp');
  const compLabels = Object.keys(apByCompany);
  makeChart('chartApComp', {
    type:'doughnut',
    data:{ labels:compLabels, datasets:[{ data:compLabels.map(k=>apByCompany[k]), backgroundColor:palette2, datalabels:DL_PIE }] },
    options:{ plugins:{legend:{position:'bottom'}} }
  });

  destroyChart('chartApDisc');
  const discLabels = Object.keys(apByDisc);
  const discFullLabels = discLabels.map(d => DISC_SHORT[d] || d);
  makeChart('chartApDisc', {
    type:'pie',
    data:{ labels:discFullLabels, datasets:[{ data:discLabels.map(k=>apByDisc[k]), backgroundColor:palette, datalabels:DL_PIE }] },
    options:{ plugins:{legend:{position:'bottom'}} }
  });

  // Table
  const totalPages = Math.ceil(filtered.length / APUNCH_PER_PAGE);
  if(apunchPage > totalPages) apunchPage = 1;
  const start = (apunchPage - 1) * APUNCH_PER_PAGE;
  const pageItems = filtered.slice(start, start + APUNCH_PER_PAGE);

  let html = `<table class="data-table" style="min-width:1000px;"><thead><tr>
    <th>PL ID</th><th style="text-align:left;">Asset Tag</th><th>Discipline</th><th>Company</th><th style="text-align:left;">Subsystem</th><th style="text-align:left;">Description</th>
  </tr></thead><tbody>`;
  pageItems.forEach(p => {
    const discFull = DISC_SHORT[p.di] || p.di || '-';
    html += `<tr>
      <td style="font-weight:700;color:#C00000;">${p.id}</td>
      <td class="left">${p.as}</td>
      <td><span class="badge-open">${discFull}</span></td>
      <td>${p.co}</td>
      <td class="left" style="font-size:11px;">${p.su}</td>
      <td class="left" style="font-size:11px;max-width:300px;overflow:hidden;text-overflow:ellipsis;">${p.de}</td>
    </tr>`;
  });
  if(!pageItems.length) html += `<tr><td colspan="6" style="padding:20px;color:#999;">No matching A-punches</td></tr>`;
  html += `</tbody></table>`;
  document.getElementById('apunchTableWrap').innerHTML = html;

  // Pagination
  let pagHtml = `<button onclick="apunchPage=1;renderAPunch()" ${apunchPage<=1?'disabled':''}>«</button>
    <button onclick="apunchPage=Math.max(1,apunchPage-1);renderAPunch()" ${apunchPage<=1?'disabled':''}>‹</button>`;
  for(let i=Math.max(1,apunchPage-2); i<=Math.min(totalPages,apunchPage+2); i++){
    pagHtml += `<button class="${i===apunchPage?'active':''}" onclick="apunchPage=${i};renderAPunch()">${i}</button>`;
  }
  pagHtml += `<button onclick="apunchPage=Math.min(${totalPages},apunchPage+1);renderAPunch()" ${apunchPage>=totalPages?'disabled':''}>›</button>
    <button onclick="apunchPage=${totalPages};renderAPunch()" ${apunchPage>=totalPages?'disabled':''}>»</button>
    <span style="font-size:12px;color:var(--muted);margin-left:8px;">${filtered.length} items</span>`;
  document.getElementById('apunchPagination').innerHTML = pagHtml;
}

// ==================== PAGE 3: SUBSYSTEM ====================
let subPage = 1;
const SUB_PER_PAGE = 30;

function getUniquePriorities(){
  const set = new Set();
  SUBS.forEach(s => { if(s.pri) set.add(s.pri); });
  return Array.from(set).sort();
}

function renderSubsystem(){
  const msF = document.getElementById('filterSubMs').value;
  const discF = document.getElementById('filterSubDisc').value;
  const compF = document.getElementById('filterSubComp').value;
  const search = document.getElementById('searchSub').value.toLowerCase().trim();

  // Filter ETASKS by company, discipline, and category (matches Excel COUNTIFS)
  let tasks = ETASKS.items.filter(t => {
    if(t.ca !== 'PCOM - Precommissioning') return false;
    if(t.st === 'Blank' || t.st === '') return false;
    if(compF !== 'ALL' && t.co !== compF) return false;
    if(discF !== 'ALL' && t.ds !== discF) return false;
    return true;
  });

  // Per-subsystem priority map (matches Excel: each row counts tasks with same priority as D column)
  const subPri = {};
  SUBS.forEach(s => { subPri[s.sub] = s.pri; });

  // Per-subsystem stats from filtered tasks (per-row priority, per Excel D column)
  const stats = {};
  tasks.forEach(t => {
    const pri = subPri[t.su];
    if(pri && t.pr !== pri) return;
    if(!stats[t.su]) stats[t.su] = {tt:0,cl:0};
    stats[t.su].tt++;
    if(t.st === 'Closed') stats[t.su].cl++;
  });

  // Filter APUNCH by category A, company, discipline, priority (matches Excel I column)
  let punches = APUNCH.items.filter(p => p.ca === 'A');
  if(compF !== 'ALL') punches = punches.filter(p => p.co === compF);
  if(discF !== 'ALL') punches = punches.filter(p => p.di && p.di.startsWith(discF[0]));
  if(msF !== 'ALL') punches = punches.filter(p => p.pr === msF);
  const apCount = {};
  punches.forEach(p => { if(p.su) apCount[p.su] = (apCount[p.su]||0)+1; });

  // Build filtered list from SUBS, attach computed stats
  let filtered = SUBS.filter(s => {
    const st = stats[s.sub] || {tt:0,cl:0};
    if(msF !== 'ALL' && s.pri !== msF) return false;
    if(compF !== 'ALL' || discF !== 'ALL'){
      if(st.tt === 0) return false;
    }
    if(search && !s.sub.toLowerCase().includes(search) && !s.sys.toLowerCase().includes(search)) return false;
    return true;
  }).map(s => {
    const st = stats[s.sub] || {tt:0,cl:0};
    return {
      sys:s.sys, sub:s.sub, pri:s.pri, st:s.st,
      tt: st.tt,
      cl: st.cl,
      pd: st.tt - st.cl,
      pc: st.tt ? Math.round(st.cl/st.tt*1000)/10 : 0,
      ap: apCount[s.sub] || 0
    };
  });

  // KPIs
  const totalTasks = filtered.reduce((a,s)=>a+s.tt,0);
  const closedTasks = filtered.reduce((a,s)=>a+s.cl,0);
  const totalAp = filtered.reduce((a,s)=>a+s.ap,0);
  const completed = filtered.filter(s => s.st === 'COMPLETED').length;
  const pct = totalTasks ? Math.round(closedTasks/totalTasks*100*10)/10 : 0;

  document.getElementById('subKpis').innerHTML = `
    <div class="kpi"><div class="icon">📋</div><div class="val">${filtered.length}</div><div class="lbl">Subsystems</div></div>
    <div class="kpi teal"><div class="icon">✅</div><div class="val">${completed}</div><div class="lbl">Completed</div></div>
    <div class="kpi gold"><div class="icon">📈</div><div class="val">${pct}%</div><div class="lbl">Avg Task Completion</div></div>
    <div class="kpi pink"><div class="icon">📌</div><div class="val">${totalAp}</div><div class="lbl">Open A-Punches</div></div>
  `;

  // Top 20 chart
  const sorted = [...filtered].sort((a,b)=>b.pc - a.pc).slice(0,20);
  //destroyChart('chartSubTop');
  makeChart('chartSubTop', {
    type:'bar',
    data:{ labels: sorted.map(s=>s.sub.split(' - ').pop()),
      datasets:[{label:'% Complete', data:sorted.map(s=>s.pc), backgroundColor:'#00e0c6', datalabels:{...DL_BAR, anchor:'end', align:'right'}}]
    },
    options:{ indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{beginAtZero:true,max:100}} }
  });

  // By milestone
  const msCount = {};
  filtered.forEach(s => { const k = s.pri.replace('PS5 - ',''); msCount[k] = (msCount[k]||0)+1; });
  const msLabels = Object.keys(msCount).sort();
  //destroyChart('chartSubMs');
  makeChart('chartSubMs', {
    type:'pie',
    data:{ labels:msLabels, datasets:[{ data:msLabels.map(k=>msCount[k]), backgroundColor:MS_COLORS, datalabels:DL_PIE }] },
    options:{ plugins:{legend:{position:'bottom'}} }
  });

  // A-punch by subsystem (top 15)
  const topAp = [...filtered].sort((a,b)=>b.ap - a.ap).slice(0,15);
  //destroyChart('chartSubAp');
  makeChart('chartSubAp', {
    type:'bar',
    data:{ labels: topAp.map(s=>s.sub.split(' - ').pop()),
      datasets:[{label:'Open A-Punches', data:topAp.map(s=>s.ap), backgroundColor:'#ff4d8d', datalabels:{...DL_BAR, anchor:'end', align:'right'}}]
    },
    options:{ indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{beginAtZero:true}} }
  });

  // Status distribution
  const statusCount = {};
  filtered.forEach(s => { const k = s.st || 'In Progress'; statusCount[k] = (statusCount[k]||0)+1; });
  const stLabels = Object.keys(statusCount);
  const stColors = stLabels.map(k => k === 'COMPLETED' ? '#2ecc71' : (k === 'In Progress' ? '#ffb627' : '#3da8ff'));
  makeChart('chartSubStatus', {
    type:'doughnut',
    data:{ labels:stLabels, datasets:[{ data:stLabels.map(k=>statusCount[k]), backgroundColor:stColors, datalabels:DL_PIE }] },
    options:{ plugins:{legend:{position:'bottom', labels:{font:{size:13},generateLabels:function(chart){return chart.data.labels.map((l,i)=>({text:l+' ('+chart.data.datasets[0].data[i]+')',fillStyle:chart.data.datasets[0].backgroundColor[i],strokeStyle:'transparent',index:i}))}}}} }
  });

  // Table
  const totalPg = Math.ceil(filtered.length / SUB_PER_PAGE);
  if(subPage > totalPg) subPage = 1;
  const st = (subPage-1)*SUB_PER_PAGE;
  const pgItems = filtered.slice(st, st+SUB_PER_PAGE);

  let html = `<table class="data-table" style="min-width:1100px;"><thead><tr>
    <th style="text-align:left;">System</th>
    <th style="text-align:left;">Subsystem</th>
    <th>Milestone</th>
    <th>Total</th><th>Closed</th><th>Pending</th><th>%</th><th>A-Punch</th><th>Status</th>
  </tr></thead><tbody>`;
  pgItems.forEach(s => {
    const statColor = s.st === 'COMPLETED' ? '#1d6f42' : (s.ap > 0 ? '#C00000' : '#333');
    const statBg = s.st === 'COMPLETED' ? '#e8f5e9' : (s.ap > 0 ? '#ffebee' : '#fff');
    html += `<tr style="background:${statBg};">
      <td class="left" style="font-size:11px;">${s.sys.replace('PS5-','')}</td>
      <td class="left" style="font-weight:bold;font-size:11px;">${s.sub}</td>
      <td><span class="badge">${s.pri.replace('PS5 - ','')}</span></td>
      <td>${s.tt}</td>
      <td style="color:#1d6f42;font-weight:700;">${s.cl}</td>
      <td style="color:#C00000;font-weight:700;">${s.pd}</td>
      <td>${pctBar(s.pc)}</td>
      <td style="color:${s.ap>0?'#C00000;font-weight:700':'#999'}">${s.ap}</td>
      <td style="color:${statColor};font-weight:700;">${s.st || 'In Progress'}</td>
    </tr>`;
  });
  if(!pgItems.length) html += `<tr><td colspan="9" style="padding:20px;color:#999;">No matching subsystems</td></tr>`;
  html += `</tbody></table>`;
  document.getElementById('subsystemTableWrap').innerHTML = html;

  // Pagination
  let pagHtml = `<button onclick="subPage=1;renderSubsystem()" ${subPage<=1?'disabled':''}>«</button>
    <button onclick="subPage=Math.max(1,subPage-1);renderSubsystem()" ${subPage<=1?'disabled':''}>‹</button>`;
  for(let i=Math.max(1,subPage-2); i<=Math.min(totalPg,subPage+2); i++){
    pagHtml += `<button class="${i===subPage?'active':''}" onclick="subPage=${i};renderSubsystem()">${i}</button>`;
  }
  pagHtml += `<button onclick="subPage=Math.min(${totalPg},subPage+1);renderSubsystem()" ${subPage>=totalPg?'disabled':''}>›</button>
    <button onclick="subPage=${totalPg};renderSubsystem()" ${subPage>=totalPg?'disabled':''}>»</button>
    <span style="font-size:12px;color:var(--muted);margin-left:8px;">${filtered.length} items</span>`;
  document.getElementById('subPagination').innerHTML = pagHtml;
}

const _charts = {};

function destroyChart(id){
  if(_charts[id]) { _charts[id].destroy(); delete _charts[id]; }
}

function makeChart(id, config){
  destroyChart(id);
  const el = document.getElementById(id);
  if(!el) return null;
  const chart = new Chart(el, config);
  _charts[id] = chart;
  return chart;
}

// ==================== PAGE 4: TASKS ====================
let tasksPage = 1;
const TASKS_PER_PAGE = 50;

function initTasksFilters(){
  const selc = document.getElementById('filterTComp');
  Object.keys(ETASKS.by_company).forEach(c=>{ const o=document.createElement('option'); o.value=c; o.textContent=c; selc.appendChild(o); });
  const seld = document.getElementById('filterTDisc');
  Object.keys(ETASKS.by_discipline).forEach(d=>{ const o=document.createElement('option'); o.value=d; o.textContent=DISC_SHORT[d]||d; seld.appendChild(o); });
  const selp = document.getElementById('filterTPri');
  Object.keys(ETASKS.by_priority).forEach(p=>{ const o=document.createElement('option'); o.value=p; o.textContent=p.replace('PS5 - ',''); selp.appendChild(o); });
  const sels = document.getElementById('filterTState');
  Object.keys(ETASKS.by_state).forEach(s=>{ const o=document.createElement('option'); o.value=s; o.textContent=s; sels.appendChild(o); });
}

function renderTasks(){
  const compF = document.getElementById('filterTComp').value;
  const discF = document.getElementById('filterTDisc').value;
  const priF = document.getElementById('filterTPri').value;
  const stateF = document.getElementById('filterTState').value;
  const search = document.getElementById('searchTasks').value.toLowerCase().trim();

  let filtered = ETASKS.items.filter(t => {
    if(compF !== 'ALL' && t.co !== compF) return false;
    if(discF !== 'ALL' && t.di !== discF) return false;
    if(priF !== 'ALL' && t.pr !== priF) return false;
    if(stateF !== 'ALL' && t.st !== stateF) return false;
    if(search) return t.id.toLowerCase().includes(search) || t.as.toLowerCase().includes(search) || t.de.toLowerCase().includes(search);
    return true;
  });

  // KPIs (from filtered)
  const closedCount = filtered.filter(t => t.st === 'Closed').length;
  const openCount = filtered.filter(t => t.st === 'To be completed').length;
  const tOther = filtered.length - closedCount - openCount;
  document.getElementById('tasksKpis').innerHTML = `
    <div class="kpi teal"><div class="icon">📋</div><div class="val">${filtered.length.toLocaleString()}</div><div class="lbl">Filtered Tasks</div></div>
    <div class="kpi"><div class="icon">✅</div><div class="val">${closedCount.toLocaleString()}</div><div class="lbl">Closed</div></div>
    <div class="kpi pink"><div class="icon">⏳</div><div class="val">${openCount.toLocaleString()}</div><div class="lbl">To Be Completed</div></div>
    <div class="kpi gold"><div class="icon">🔍</div><div class="val">${ETASKS.total.toLocaleString()}</div><div class="lbl">Total (Unfiltered)</div></div>
  `;

  // Charts (from filtered)
  const tState = {}; const tCompany = {}; const tDisc = {}; const tPri = {};
  filtered.forEach(t => {
    tState[t.st] = (tState[t.st]||0)+1;
    tCompany[t.co] = (tCompany[t.co]||0)+1;
    tDisc[t.di] = (tDisc[t.di]||0)+1;
    tPri[t.pr] = (tPri[t.pr]||0)+1;
  });

  destroyChart('chartTState');
  const stLabels = Object.keys(tState);
  makeChart('chartTState', {
    type:'doughnut',
    data:{ labels:stLabels, datasets:[{ data:stLabels.map(k=>tState[k]), backgroundColor:['#00e0c6','#2ecc71','#ffb627','#3da8ff','#ff4d8d'], datalabels:DL_PIE }] },
    options:{ plugins:{legend:{position:'bottom'}} }
  });

  destroyChart('chartTComp');
  const compLabels = Object.keys(tCompany);
  makeChart('chartTComp', {
    type:'bar',
    data:{ labels:compLabels, datasets:[{ label:'Tasks', data:compLabels.map(k=>tCompany[k]), backgroundColor:palette2, datalabels:DL_BAR }] },
    options:{ plugins:{legend:{display:false}}, scales:{y:{beginAtZero:true}} }
  });

  destroyChart('chartTDisc');
  const discLabels = Object.keys(tDisc);
  const discFullLabels = discLabels.map(d => DISC_SHORT[d] || d);
  makeChart('chartTDisc', {
    type:'bar',
    data:{ labels:discFullLabels, datasets:[{ label:'Tasks', data:discLabels.map(k=>tDisc[k]), backgroundColor:discLabels.map(d=>DISC_COLORS[DISC_SHORT[d]]||'#9b59f6'), datalabels:DL_BAR }] },
    options:{ indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{beginAtZero:true}} }
  });

  destroyChart('chartTPri');
  const priLabels = Object.keys(tPri);
  makeChart('chartTPri', {
    type:'pie',
    data:{ labels:priLabels.map(p=>p.replace('PS5 - ','')), datasets:[{ data:priLabels.map(k=>tPri[k]), backgroundColor:MS_COLORS, datalabels:DL_PIE }] },
    options:{ plugins:{legend:{position:'bottom'}} }
  });

  // Table
  const totalPages = Math.ceil(filtered.length / TASKS_PER_PAGE);
  if(tasksPage > totalPages) tasksPage = 1;
  const start = (tasksPage - 1) * TASKS_PER_PAGE;
  const pageItems = filtered.slice(start, start + TASKS_PER_PAGE);

  let html = `<table class="data-table" style="min-width:1200px;"><thead><tr>
    <th>Task ID</th><th style="text-align:left;">Asset Tag</th><th style="text-align:left;">Description</th>
    <th>Discipline</th><th>Company</th><th style="text-align:left;">Subsystem</th><th>Milestone</th><th>State</th>
  </tr></thead><tbody>`;
  pageItems.forEach(t => {
    const stateColor = t.st === 'Closed' ? '#1d6f42' : (t.st === 'To be completed' ? '#C00000' : '#bf8f00');
    html += `<tr>
      <td style="font-size:11px;font-weight:600;">${t.id}</td>
      <td class="left" style="font-size:11px;">${t.as}</td>
      <td class="left" style="font-size:11px;max-width:250px;overflow:hidden;text-overflow:ellipsis;">${t.de}</td>
      <td><span class="badge">${DISC_SHORT[t.di]||t.di||'-'}</span></td>
      <td style="font-size:11px;">${t.co}</td>
      <td class="left" style="font-size:10px;max-width:180px;overflow:hidden;text-overflow:ellipsis;">${t.su}</td>
      <td style="font-size:11px;">${t.pr.replace('PS5 - ','')}</td>
      <td style="color:${stateColor};font-weight:700;font-size:11px;">${t.st}</td>
    </tr>`;
  });
  if(!pageItems.length) html += `<tr><td colspan="8" style="padding:20px;color:#999;">No matching tasks</td></tr>`;
  html += `</tbody></table>`;
  document.getElementById('tasksTableWrap').innerHTML = html;

  // Pagination
  let pagHtml = `<button onclick="tasksPage=1;renderTasks()" ${tasksPage<=1?'disabled':''}>«</button>
    <button onclick="tasksPage=Math.max(1,tasksPage-1);renderTasks()" ${tasksPage<=1?'disabled':''}>‹</button>`;
  for(let i=Math.max(1,tasksPage-2); i<=Math.min(totalPages,tasksPage+2); i++){
    pagHtml += `<button class="${i===tasksPage?'active':''}" onclick="tasksPage=${i};renderTasks()">${i}</button>`;
  }
  pagHtml += `<button onclick="tasksPage=Math.min(${totalPages},tasksPage+1);renderTasks()" ${tasksPage>=totalPages?'disabled':''}>›</button>
    <button onclick="tasksPage=${totalPages};renderTasks()" ${tasksPage>=totalPages?'disabled':''}>»</button>
    <span style="font-size:12px;color:var(--muted);margin-left:8px;">${filtered.length.toLocaleString()} items</span>`;
  document.getElementById('tasksPagination').innerHTML = pagHtml;
}

// ==================== PAGE 5: HYDROTESTS ====================
const HT_DATA = DATA.hydrotests;

function renderHydrotests(){
  const packF = document.getElementById('filterHtPack').value;
  const msF = document.getElementById('filterHtMs').value;

  const ALL_MS = HT_DATA.milestones.filter(m => m.name !== 'TOTAL');
  const TOTAL_ROW = HT_DATA.milestones.filter(m => m.name === 'TOTAL')[0] || {total:0,closed:0,pending:0,pct:0,packs:{}};

  let filteredMs = ALL_MS;
  if(msF !== 'ALL') filteredMs = filteredMs.filter(m => m.name === msF || m.name.replace('PS5 - ','') === msF);

  const allPacks = ['CLX01','CLX02A','TMX03','TMX05','CLX02B'];
  let showPacks = allPacks;
  if(packF !== 'ALL') showPacks = [packF];

  let totalAll = 0, totalClosed = 0, totalPending = 0;
  filteredMs.forEach(m => {
    if(packF === 'ALL') {
      totalAll += m.total; totalClosed += m.closed; totalPending += m.pending;
    } else {
      const pkd = m.packs[packF] || {total:0, closed:0};
      totalAll += pkd.total; totalClosed += pkd.closed; totalPending += pkd.total - pkd.closed;
    }
  });
  const filteredPct = totalAll ? Math.round(totalClosed/totalAll*100*10)/10 : 0;

  const overallTotal = HT_DATA.overall || '0/0';
  const overallPctVal = HT_DATA.overall_pct || 0;

  document.getElementById('htKpis').innerHTML = `
    <div class="kpi teal"><div class="icon">\u{1F4A7}</div><div class="val">${packF === 'ALL' ? overallTotal : totalClosed+'/'+totalAll}</div><div class="lbl">Hydrotest Progress</div></div>
    <div class="kpi gold"><div class="icon">\u{1F4C8}</div><div class="val">${packF === 'ALL' ? overallPctVal : filteredPct}%</div><div class="lbl">Completion</div></div>
    <div class="kpi"><div class="icon">\u{2705}</div><div class="val">${totalClosed}</div><div class="lbl">Closed Tasks</div></div>
    <div class="kpi pink"><div class="icon">\u{23F3}</div><div class="val">${totalPending}</div><div class="lbl">Pending Tasks</div></div>
  `;

  // Chart 1: Pack completion
  if(packF === 'ALL') {
    const pNames = HT_DATA.packs.map(p => p.name.split(' - ')[0]);
    const pPcts = HT_DATA.packs.map(p => p.pct);
    const pColors = ['#3da8ff','#ffb627','#ff4d8d','#00e0c6','#9b59f6'];
    makeChart('chartHtPacks', {
      type:'bar',
      data:{ labels: pNames,
        datasets:[{label:'% Complete', data:pPcts, backgroundColor:pColors, datalabels:{...DL_BAR, anchor:'end', align:'right'}}]
      },
      options:{ indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{beginAtZero:true,max:100}} }
    });
  } else {
    const msData = filteredMs.map(m => {
      const pkd = m.packs[packF] || {total:0, closed:0};
      const pct = pkd.total ? Math.round(pkd.closed/pkd.total*100*10)/10 : 0;
      return {label: m.name.replace('PS5 - ',''), pct, closed: pkd.closed, total: pkd.total};
    });
    makeChart('chartHtPacks', {
      type:'bar',
      data:{ labels: msData.map(d=>d.label),
        datasets:[{label:packF+' % Complete', data:msData.map(d=>d.pct), backgroundColor:'#3da8ff', datalabels:DL_BAR}]
      },
      options:{ plugins:{legend:{display:false}}, scales:{y:{beginAtZero:true,max:100}} }
    });
  }

  // Chart 2: Milestone completion
  const chartLabels = filteredMs.map(m => m.name.replace('PS5 - ',''));
  let chPcts, chClosed, chPending;
  if(packF === 'ALL') {
    chPcts = filteredMs.map(m => m.pct);
    chClosed = filteredMs.map(m => m.closed);
    chPending = filteredMs.map(m => m.pending);
  } else {
    chPcts = filteredMs.map(m => {
      const pkd = m.packs[packF] || {total:0, closed:0};
      return pkd.total ? Math.round(pkd.closed/pkd.total*100*10)/10 : 0;
    });
    chClosed = filteredMs.map(m => (m.packs[packF] || {closed:0}).closed);
    chPending = filteredMs.map(m => { const pkd=m.packs[packF]||{total:0,closed:0}; return pkd.total-pkd.closed; });
  }

  makeChart('chartHtMilestones', {
    type:'bar',
    data:{ labels: chartLabels,
      datasets:[{label:'% Complete', data:chPcts, backgroundColor:'#3da8ff', datalabels:DL_BAR}]
    },
    options:{ plugins:{legend:{display:false}}, scales:{y:{beginAtZero:true,max:100}} }
  });

  makeChart('chartHtStacked', {
    type:'bar',
    data:{ labels: chartLabels,
      datasets:[
        {label:'Closed', data:chClosed, backgroundColor:'#00e0c6', datalabels:DL_STACK},
        {label:'Pending', data:chPending, backgroundColor:'#ff4d8d', datalabels:DL_STACK}
      ]
    },
    options:{ plugins:{legend:{position:'bottom'}}, scales:{x:{stacked:true},y:{stacked:true,beginAtZero:true}} }
  });

  // Table
  let html = `<table class="data-table" style="min-width:${packF === 'ALL' ? '900' : '500'}px;"><thead><tr>
    <th style="text-align:left;">Milestone</th><th>Total</th><th>Closed</th><th>Pending</th><th>%</th>`;
  if(packF === 'ALL') {
    showPacks.forEach(pk => { html += '<th colspan="2">'+pk+'</th>'; });
    html += '</tr><tr><th></th><th></th><th></th><th></th><th></th>';
    showPacks.forEach(() => { html += '<th>T</th><th>C</th>'; });
  } else {
    html += '<th>Closed</th><th>%</th>';
  }
  html += '</tr></thead><tbody>';

  filteredMs.forEach(m => {
    let rTotal = m.total, rClosed = m.closed, rPct = m.pct;
    if(packF !== 'ALL') {
      const pkd = m.packs[packF] || {total:0, closed:0};
      rTotal = pkd.total; rClosed = pkd.closed;
      rPct = pkd.total ? Math.round(pkd.closed/pkd.total*100*10)/10 : 0;
    }
    html += '<tr><td class="left" style="font-weight:bold;">'+m.name.replace('PS5 - ','')+'</td>' +
      '<td>'+rTotal+'</td><td style="color:#1d6f42;font-weight:700;">'+rClosed+'</td>' +
      '<td style="color:#C00000;font-weight:700;">'+(rTotal-rClosed)+'</td><td>'+pctBar(rPct)+'</td>';
    if(packF === 'ALL') {
      showPacks.forEach(pk => {
        const pkd = m.packs[pk] || {total:0, closed:0};
        html += '<td>'+(pkd.total||'-')+'</td><td style="color:#1d6f42;font-weight:700;">'+(pkd.closed||0)+'</td>';
      });
    } else {
      html += '<td style="color:#1d6f42;font-weight:700;">'+rClosed+'</td><td>'+rPct+'%</td>';
    }
    html += '</tr>';
  });

  // Total row
  html += '<tr class="total"><td class="left">TOTAL</td>' +
    '<td>'+totalAll+'</td><td>'+totalClosed+'</td><td>'+totalPending+'</td><td>'+pctBar(filteredPct)+'</td>';
  if(packF === 'ALL') {
    showPacks.forEach(pk => {
      const pkd = TOTAL_ROW.packs[pk] || {total:0, closed:0};
      html += '<td>'+(pkd.total||'-')+'</td><td>'+(pkd.closed||0)+'</td>';
    });
  } else {
    html += '<td>'+totalClosed+'</td><td>'+filteredPct+'%</td>';
  }
  html += '</tr></tbody></table>';
  document.getElementById('htTableWrap').innerHTML = html;
}function exportHydrotestsExcel(){
  const packF = document.getElementById('filterHtPack').value;
  const msF = document.getElementById('filterHtMs').value;
  const ALL_MS = HT_DATA.milestones.filter(m => m.name !== 'TOTAL');
  const TOTAL_ROW = HT_DATA.milestones.filter(m => m.name === 'TOTAL')[0] || {total:0,closed:0,pending:0,pct:0,packs:{}};
  let tableMs = ALL_MS;
  if(msF !== 'ALL') tableMs = tableMs.filter(m => m.name === msF || m.name.replace('PS5 - ','') === msF);
  const allPacks = ['CLX01','CLX02A','TMX03','TMX05','CLX02B'];
  let showPacks = allPacks;
  if(packF !== 'ALL') showPacks = [packF];

  let html = buildExcelHeader('Hydrotests Summary - PS5');
  html += '<tr><th colspan="5" class="hdr" style="text-align:left;">Milestone</th>';
  if(packF === 'ALL') {
    showPacks.forEach(pk => { html += '<th colspan="2" class="hdr">'+pk+'</th>'; });
    html += '</tr><tr><th class="hdr">Milestone</th><th class="hdr">Total</th><th class="hdr">Closed</th><th class="hdr">Pending</th><th class="hdr">%</th>';
    showPacks.forEach(() => { html += '<th class="hdr">T</th><th class="hdr">C</th>'; });
  } else {
    html += '<th class="hdr">Closed</th><th class="hdr">%</th></tr><tr><th class="hdr">Milestone</th><th class="hdr">Total</th><th class="hdr">Closed</th><th class="hdr">Pending</th><th class="hdr">%</th><th class="hdr">Closed</th><th class="hdr">%</th>';
  }
  html += '</tr>';

  tableMs.forEach(m => {
    let rTotal = m.total, rClosed = m.closed, rPct = m.pct;
    if(packF !== 'ALL') {
      const pkd = m.packs[packF] || {total:0, closed:0};
      rTotal = pkd.total; rClosed = pkd.closed;
      rPct = pkd.total ? Math.round(pkd.closed/pkd.total*1000)/10 : 0;
    }
    html += '<tr><td style="font-weight:bold;">'+m.name+'</td><td>'+rTotal+'</td><td>'+rClosed+'</td><td>'+(rTotal-rClosed)+'</td><td>'+rPct+'%</td>';
    if(packF === 'ALL') {
      showPacks.forEach(pk => {
        const pkd = m.packs[pk] || {total:0, closed:0};
        html += '<td>'+(pkd.total||'-')+'</td><td>'+(pkd.closed||0)+'</td>';
      });
    } else {
      html += '<td>'+rClosed+'</td><td>'+rPct+'%</td>';
    }
    html += '</tr>';
  });

  html += '<tr class="tot"><td class="tot">TOTAL</td><td class="tot">'+TOTAL_ROW.total+'</td><td class="tot">'+TOTAL_ROW.closed+'</td><td class="tot">'+TOTAL_ROW.pending+'</td><td class="tot">'+TOTAL_ROW.pct+'%</td>';
  if(packF === 'ALL') {
    showPacks.forEach(pk => {
      const pkd = TOTAL_ROW.packs[pk] || {total:0, closed:0};
      html += '<td class="tot">'+(pkd.total||'-')+'</td><td class="tot">'+(pkd.closed||0)+'</td>';
    });
  } else {
    html += '<td class="tot">'+TOTAL_ROW.closed+'</td><td class="tot">'+TOTAL_ROW.pct+'%</td>';
  }
  html += '</tr></table></body></html>';
  downloadExcel(html, 'Hydrotests_Summary');
}
// ==================== EXPORTS ====================
function buildExcelHeader(title){
  return '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
  '<head><meta charset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Sheet1</x:Name>' +
  '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
  '<style>td,th{font-family:Arial,sans-serif;font-size:12px;vertical-align:middle;}' +
  '.hdr{background:#ED7D31;color:#000;font-weight:bold;text-align:center;border:1px solid #9c4a14;padding:8px;}' +
  '.tot{background:#404040;color:#fff;font-weight:bold;}</style></head><body>' +
  '<table border="1" cellspacing="0" cellpadding="5" style="border-collapse:collapse;">' +
  '<tr><th colspan="10" class="hdr" style="font-size:14px;">EACOP PS5 - ' + title + '</th></tr>';
}

function downloadExcel(html, name){
  const blob = new Blob(['\ufeff'+html], {type:'application/vnd.ms-excel;charset=utf-8'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name+'.xls'; a.click();
  URL.revokeObjectURL(url);
}

function exportOverallExcel(){
  const comp = document.getElementById('filterCompany').value;
  const src = getOverallDataByDisc(comp);
  if(!src) return;
  let html = buildExcelHeader('Overall Pre-Commissioning Progress - ' + comp);
  html += `<tr><th rowspan="2" class="hdr" style="text-align:left;">Discipline</th>
    <th rowspan="2" class="hdr">Total</th><th rowspan="2" class="hdr">Closed</th><th rowspan="2" class="hdr">Pending</th><th rowspan="2" class="hdr">%</th>`;
  MILESTONES.forEach(m => { html += `<th colspan="4" class="hdr">${m.replace('MILESTONE ','M-')}</th>`; });
  html += `</tr><tr>`;
  MILESTONES.forEach(() => { html += `<th class="hdr">T</th><th class="hdr">C</th><th class="hdr">P</th><th class="hdr">%</th>`; });
  html += `</tr>`;

  DISC_NAMES.forEach(dn => {
    if(!src[dn]) return;
    const d = src[dn];
    html += `<tr><td style="font-weight:bold;">${dn}</td><td>${d.t}</td><td>${d.c}</td><td>${d.p}</td><td>${d.pc}%</td>`;
    MILESTONES.forEach(m => {
      const ms = d.ms[m] || {};
      html += `<td>${ms.t||'-'}</td><td>${ms.c||'-'}</td><td>${ms.p||'-'}</td><td>${ms.pc?ms.pc+'%':'-'}</td>`;
    });
    html += `</tr>`;
  });
  if(src['Grand Total']){
    const g = src['Grand Total'];
    html += `<tr class="tot"><td class="tot">GRAND TOTAL</td><td class="tot">${g.t}</td><td class="tot">${g.c}</td><td class="tot">${g.p}</td><td class="tot">${g.pc}%</td>`;
    MILESTONES.forEach(m => {
      const ms = g.ms[m] || {};
      html += `<td class="tot">${ms.t||'-'}</td><td class="tot">${ms.c||'-'}</td><td class="tot">${ms.p||'-'}</td><td class="tot">${ms.pc?ms.pc+'%':'-'}</td>`;
    });
    html += `</tr>`;
  }
  html += `</table></body></html>`;
  downloadExcel(html, 'Overall_Progress_' + comp.replace(/\s+/g,'_'));
}

function exportAPunchExcel(){
  const compF = document.getElementById('filterApComp').value;
  const discF = document.getElementById('filterApDisc').value;
  const search = document.getElementById('searchAPunch').value.toLowerCase().trim();
  let filtered = APUNCH.items.filter(p => {
    if(compF !== 'ALL' && p.co !== compF) return false;
    if(discF !== 'ALL' && p.di !== discF) return false;
    if(search) return p.id.toLowerCase().includes(search) || p.as.toLowerCase().includes(search);
    return true;
  });
  let html = buildExcelHeader('Open A-Punches List');
  html += `<tr><th class="hdr">PL ID</th><th class="hdr" style="text-align:left;">Asset Tag</th><th class="hdr">Discipline</th><th class="hdr">Company</th><th class="hdr" style="text-align:left;">Subsystem</th><th class="hdr" style="text-align:left;">Description</th></tr>`;
  filtered.forEach(p => {
    html += `<tr><td style="font-weight:bold;color:#C00000;">${p.id}</td><td>${p.as}</td><td>${DISC_SHORT[p.di]||p.di||'-'}</td><td>${p.co}</td><td>${p.su}</td><td>${p.de}</td></tr>`;
  });
  html += `</table></body></html>`;
  downloadExcel(html, 'Open_A_Punches');
}

function exportSubsystemExcel(){
  const msF = document.getElementById('filterSubMs').value;
  const discF = document.getElementById('filterSubDisc').value;
  const search = document.getElementById('searchSub').value.toLowerCase().trim();
  let filtered = SUBS.filter(s => {
    if(msF !== 'ALL' && s.pri !== msF) return false;
    if(search && !s.sub.toLowerCase().includes(search)) return false;
    return true;
  });
  let html = buildExcelHeader('Subsystem Progress Status');
  html += `<tr><th class="hdr" style="text-align:left;">System</th><th class="hdr" style="text-align:left;">Subsystem</th>
    <th class="hdr">Milestone</th><th class="hdr">Total</th><th class="hdr">Closed</th><th class="hdr">Pending</th>
    <th class="hdr">%</th><th class="hdr">A-Punch</th><th class="hdr">Status</th></tr>`;
  filtered.forEach(s => {
    html += `<tr><td>${s.sys}</td><td style="font-weight:bold;">${s.sub}</td><td>${s.pri}</td>
      <td>${s.tt}</td><td>${s.cl}</td><td>${s.pd}</td><td>${s.pc}%</td><td>${s.ap}</td><td>${s.st||'In Progress'}</td></tr>`;
  });
  html += `</table></body></html>`;
  downloadExcel(html, 'Subsystem_Progress');
}

function exportTasksExcel(){
  const compF = document.getElementById('filterTComp').value;
  const discF = document.getElementById('filterTDisc').value;
  const priF = document.getElementById('filterTPri').value;
  const stateF = document.getElementById('filterTState').value;
  const search = document.getElementById('searchTasks').value.toLowerCase().trim();
  let filtered = ETASKS.items.filter(t => {
    if(compF !== 'ALL' && t.co !== compF) return false;
    if(discF !== 'ALL' && t.di !== discF) return false;
    if(priF !== 'ALL' && t.pr !== priF) return false;
    if(stateF !== 'ALL' && t.st !== stateF) return false;
    if(search && !t.id.toLowerCase().includes(search) && !t.as.toLowerCase().includes(search)) return false;
    return true;
  });
  let html = buildExcelHeader('Exported Tasks Registry');
  html += `<tr><th class="hdr">Task ID</th><th class="hdr" style="text-align:left;">Asset Tag</th>
    <th class="hdr" style="text-align:left;">Description</th><th class="hdr">Discipline</th>
    <th class="hdr">Company</th><th class="hdr" style="text-align:left;">Subsystem</th>
    <th class="hdr">Milestone</th><th class="hdr">State</th></tr>`;
  filtered.forEach(t => {
    html += `<tr><td>${t.id}</td><td>${t.as}</td><td>${t.de}</td>
      <td>${DISC_SHORT[t.di]||t.di||'-'}</td><td>${t.co}</td><td>${t.su}</td>
      <td>${t.pr.replace('PS5 - ','')}</td><td>${t.st}</td></tr>`;
  });
  html += `</table></body></html>`;
  downloadExcel(html, 'Exported_Tasks');
}

// ==================== INIT ====================
initFilters();
initTasksFilters();
renderOverall();
renderAPunch();
renderSubsystem();
renderTasks();
renderHydrotests();
</script>
</body>
</html>"""

for out_path in [OUTPUT_FILE, OUTPUT_FILE2]:
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(HTML)
    print(f'Dashboard created: {out_path}')
    print(f'File size: {os.path.getsize(out_path):,} bytes')
