#!/usr/bin/env python3
"""KENT PLC Completion Progress Dashboard - Oil & Gas Professional Theme"""

import pandas as pd
import json
import os

EXCEL_PATH = r'C:\Users\mylap\Downloads\COMPLETION PROGRESS - KENTPLC.xlsx'
OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'KENTPLC_Completion_Dashboard.html')


def extract_data():
    xls = pd.ExcelFile(EXCEL_PATH)
    precom = pd.read_excel(xls, sheet_name='PRECOM STATUS', header=None)
    subsys_df = pd.read_excel(xls, sheet_name='SUBSYSTEMS PROGRESS', header=None)
    tasks = pd.read_excel(xls, sheet_name='EXPORTED TASKS', header=0)
    pl = pd.read_excel(xls, sheet_name='EXPORTED PL', header=0)

    milestones = []
    for i in range(7, 34):
        row = precom.iloc[i]
        name = str(row[1]) if pd.notna(row[1]) else ''
        if name.startswith('PS5 - Milestone'):
            open_val = int(row[5]) if pd.notna(row[5]) else (int(row[3]) - int(row[4]) if pd.notna(row[3]) and pd.notna(row[4]) else 0)
            ms = {
                'name': name,
                'letter': name.replace('PS5 - Milestone ', ''),
                'subsystems': int(row[2]) if pd.notna(row[2]) else 0,
                'totalTasks': int(row[3]) if pd.notna(row[3]) else 0,
                'closedTasks': int(row[4]) if pd.notna(row[4]) else 0,
                'openTasks': open_val,
                'completion': round(float(row[14]) * 100, 2) if pd.notna(row[14]) else 0,
                'disciplines': []
            }
            for j in range(i + 1, min(i + 3, 34)):
                drow = precom.iloc[j]
                dname = str(drow[1]) if pd.notna(drow[1]) else ''
                if dname.startswith('E -') or dname.startswith('I -'):
                    ms['disciplines'].append({
                        'name': dname,
                        'totalTasks': int(drow[3]) if pd.notna(drow[3]) else 0,
                        'closedTasks': int(drow[4]) if pd.notna(drow[4]) else 0,
                        'openTasks': int(drow[5]) if pd.notna(drow[5]) else 0,
                    })
            milestones.append(ms)

    subsystems = []
    for i in range(5, 128):
        row = subsys_df.iloc[i]
        if pd.notna(row[0]) and str(row[0]).startswith('PS5'):
            total = int(row[4]) if pd.notna(row[4]) else 0
            closed = int(row[5]) if pd.notna(row[5]) else 0
            subsystems.append({
                'system': str(row[0]),
                'subsystem': str(row[1]),
                'mantrac': str(row[2]) if pd.notna(row[2]) else '',
                'milestone': str(row[3]) if pd.notna(row[3]) else '',
                'total': total,
                'closed': closed,
                'elecOpen': int(row[6]) if pd.notna(row[6]) else 0,
                'instrOpen': int(row[7]) if pd.notna(row[7]) else 0,
                'pct': round(closed / total * 100, 1) if total > 0 else 0,
                'punchA': int(row[9]) if pd.notna(row[9]) else 0,
                'punchB': int(row[10]) if pd.notna(row[10]) else 0,
                'punchC': int(row[11]) if pd.notna(row[11]) else 0,
            })

    task_by_ms = tasks.groupby('Subsystem Priority').agg(
        total=('Task ID', 'count'),
        closed=('Task State', lambda x: (x == 'Closed').sum()),
        started=('Task State', lambda x: (x == 'Started (not Completed)').sum()),
        pending=('Task State', lambda x: (x == 'To be completed').sum())
    ).reset_index().to_dict('records')

    pl_by_ms = pl.groupby('Priority').agg(
        total=('Punchlist ID', 'count'),
        originated=('Status', lambda x: (x == 'Originated').sum()),
        closed=('Status', lambda x: (x == 'Closed').sum()),
        completed=('Status', lambda x: (x == 'Completed').sum())
    ).reset_index().to_dict('records')

    pl_by_cat_raw = pl['Punchlist Category (Name)'].value_counts().to_dict()

    ct = pd.crosstab(pl['Priority'], pl['Punchlist Category (Name)'])
    pl_cross = {}
    for ms_name in ct.index:
        letter = ms_name.replace('PS5 - Milestone ', '')
        pl_cross[letter] = {
            'A': int(ct.loc[ms_name, 'A']) if 'A' in ct.columns else 0,
            'B': int(ct.loc[ms_name, 'B']) if 'B' in ct.columns else 0,
            'C': int(ct.loc[ms_name, 'C']) if 'C' in ct.columns else 0,
        }

    sys_groups = {}
    for s in subsystems:
        sys = s['system']
        if sys not in sys_groups:
            sys_groups[sys] = {'total': 0, 'closed': 0, 'count': 0}
        sys_groups[sys]['total'] += s['total']
        sys_groups[sys]['closed'] += s['closed']
        sys_groups[sys]['count'] += 1
    sys_summary = [{'system': k, **v, 'pct': round(v['closed'] / v['total'] * 100, 1) if v['total'] > 0 else 0}
                   for k, v in sorted(sys_groups.items())]

    return {
        'totalTasks': 851,
        'closedTasks': 3,
        'openTasks': 848,
        'totalSubsystems': 123,
        'completeSubsystems': 0,
        'milestones': milestones,
        'subsystems': subsystems,
        'taskByMilestone': task_by_ms,
        'plByMilestone': pl_by_ms,
        'plByCategory': {
            'A - Critical': pl_by_cat_raw.get('A', 0),
            'B - Major': pl_by_cat_raw.get('B', 0),
            'C - Minor': pl_by_cat_raw.get('C', 0),
        },
        'plCross': pl_cross,
        'totalPL': len(pl),
        'closedPL': int((pl['Status'] == 'Closed').sum()),
        'completedPL': int((pl['Status'] == 'Completed').sum()),
        'originatedPL': int((pl['Status'] == 'Originated').sum()),
        'elecTasks': 538,
        'elecClosed': 3,
        'instrTasks': 313,
        'instrClosed': 0,
        'sysSummary': sys_summary,
    }


HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>KENT PLC - Completion Progress Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
:root{
  --bg:#0b1120;--bg2:#111a2e;--card:#151f32;--card2:#1a2640;--card3:#1e2d48;
  --border:#243352;--border2:#2a3d5e;
  --text:#e8edf5;--text2:#8899b4;--text3:#5a6f8a;
  --navy:#0d1b3e;--steel:#3a4f6f;
  --orange:#f59e0b;--orange2:#f97316;
  --green:#22c55e;--green2:#10b981;
  --red:#ef4444;--red2:#dc2626;
  --blue:#3b82f6;--blue2:#2563eb;
  --cyan:#06b6d4;--purple:#8b5cf6;--pink:#ec4899;
  --accent:#f59e0b;
  --glow-orange:0 0 20px rgba(245,158,11,0.15);
  --glow-blue:0 0 20px rgba(59,130,246,0.15);
  --shadow:0 2px 8px rgba(0,0,0,0.3);
  --shadow-lg:0 8px 30px rgba(0,0,0,0.4);
}
body{font-family:'Inter',system-ui,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;-webkit-font-smoothing:antialiased;overflow-x:hidden}

/* HEADER */
.header{
  background:linear-gradient(135deg,#0a1628 0%,#0d2040 30%,#122a50 60%,#1a3660 100%);
  padding:0;border-bottom:2px solid var(--orange);position:sticky;top:0;z-index:100;
}
.header-top{display:flex;align-items:center;justify-content:space-between;padding:20px 36px 12px}
.header h1{font-size:22px;font-weight:800;color:#fff;letter-spacing:-0.3px}
.header h1 span{color:var(--orange)}
.header-badge{display:flex;gap:8px;align-items:center}
.header-badge .live{background:var(--green);color:#000;font-size:10px;font-weight:700;padding:4px 10px;border-radius:20px;animation:pulse 2s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}}
.header-badge .tag{background:rgba(245,158,11,0.15);color:var(--orange);font-size:10px;font-weight:600;padding:4px 10px;border-radius:20px;border:1px solid rgba(245,158,11,0.3)}
.header .subtitle{color:var(--text2);font-size:13px;padding:0 36px 12px;font-weight:400}
.header .meta{display:flex;gap:8px;padding:0 36px 16px;flex-wrap:wrap}
.header .meta span{font-size:10px;color:var(--text2);background:var(--card);padding:4px 12px;border-radius:6px;border:1px solid var(--border);font-weight:500}

.container{max-width:1600px;margin:0 auto;padding:24px 20px}

/* TABS */
.tabs{display:flex;gap:4px;margin-bottom:24px;background:var(--card);padding:5px;border-radius:14px;border:1px solid var(--border);flex-wrap:wrap}
.tab{padding:10px 20px;border-radius:10px;cursor:pointer;font-size:12px;font-weight:600;color:var(--text3);transition:all .3s;border:none;background:none;text-transform:uppercase;letter-spacing:0.8px}
.tab:hover{background:var(--card2);color:var(--text2)}
.tab.active{background:var(--orange);color:#000;box-shadow:0 2px 12px rgba(245,158,11,0.4)}

.panel{display:none;animation:panelIn .4s ease}
.panel.active{display:block}
@keyframes panelIn{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}

.grid{display:grid;gap:16px}
.g2{grid-template-columns:repeat(2,1fr)}
.g3{grid-template-columns:repeat(3,1fr)}
.g4{grid-template-columns:repeat(4,1fr)}
.g5{grid-template-columns:repeat(5,1fr)}
@media(max-width:1200px){.g4,.g5{grid-template-columns:repeat(2,1fr)}}
@media(max-width:768px){.g2,.g3,.g4,.g5{grid-template-columns:1fr}}

/* CARDS */
.card{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:22px;transition:all .3s;box-shadow:var(--shadow)}
.card:hover{box-shadow:var(--shadow-lg);border-color:var(--border2)}
.card h3{font-size:11px;color:var(--text3);margin-bottom:14px;font-weight:600;text-transform:uppercase;letter-spacing:1px}

/* STAT CARDS */
.stat-card{position:relative;overflow:hidden}
.stat-card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px}
.stat-card.orange::before{background:linear-gradient(90deg,#f59e0b,#f97316)}
.stat-card.green::before{background:linear-gradient(90deg,#22c55e,#10b981)}
.stat-card.red::before{background:linear-gradient(90deg,#ef4444,#dc2626)}
.stat-card.blue::before{background:linear-gradient(90deg,#3b82f6,#2563eb)}
.stat-card.cyan::before{background:linear-gradient(90deg,#06b6d4,#0891b2)}
.stat-card.purple::before{background:linear-gradient(90deg,#8b5cf6,#7c3aed)}
.stat-card .value{font-size:34px;font-weight:900;letter-spacing:-1px;line-height:1}
.stat-card .label{font-size:11px;color:var(--text3);margin-top:6px;font-weight:500}
.stat-card.orange .value{color:var(--orange)}
.stat-card.green .value{color:var(--green)}
.stat-card.red .value{color:var(--red)}
.stat-card.blue .value{color:var(--blue)}
.stat-card.cyan .value{color:var(--cyan)}
.stat-card.purple .value{color:var(--purple)}

/* PROGRESS BAR */
.progress-bar{height:8px;background:var(--card3);border-radius:6px;overflow:hidden;margin-top:10px}
.progress-bar .fill{height:100%;border-radius:6px;transition:width 1s ease}
.progress-bar .fill.green{background:linear-gradient(90deg,#22c55e,#10b981)}
.progress-bar .fill.blue{background:linear-gradient(90deg,#3b82f6,#60a5fa)}
.progress-bar .fill.orange{background:linear-gradient(90deg,#f59e0b,#fbbf24)}
.progress-bar .fill.red{background:linear-gradient(90deg,#ef4444,#f87171)}
.progress-bar .fill.cyan{background:linear-gradient(90deg,#06b6d4,#22d3ee)}
.progress-bar .fill.purple{background:linear-gradient(90deg,#8b5cf6,#a78bfa)}

.chart-container{position:relative;height:300px}
.chart-container.tall{height:420px}

/* TABLES */
table{width:100%;border-collapse:collapse;font-size:12px}
table th{text-align:left;padding:12px 14px;background:var(--navy);color:var(--orange);font-weight:700;border-bottom:2px solid var(--orange);position:sticky;top:0;z-index:1;font-size:10px;text-transform:uppercase;letter-spacing:0.8px}
table td{padding:10px 14px;border-bottom:1px solid var(--border);transition:background .2s}
table tr{transition:all .2s}
table tr:hover td{background:rgba(245,158,11,0.06)}
.table-scroll{max-height:520px;overflow-y:auto;border-radius:12px;border:1px solid var(--border);box-shadow:var(--shadow)}
.table-scroll::-webkit-scrollbar{width:6px}
.table-scroll::-webkit-scrollbar-track{background:var(--card)}
.table-scroll::-webkit-scrollbar-thumb{background:var(--border2);border-radius:3px}

/* BADGES */
.badge{display:inline-block;padding:3px 10px;border-radius:8px;font-size:10px;font-weight:700;letter-spacing:0.3px;text-transform:uppercase}
.badge.green{background:rgba(34,197,94,0.15);color:#4ade80;border:1px solid rgba(34,197,94,0.3)}
.badge.red{background:rgba(239,68,68,0.15);color:#f87171;border:1px solid rgba(239,68,68,0.3)}
.badge.orange{background:rgba(245,158,11,0.15);color:#fbbf24;border:1px solid rgba(245,158,11,0.3)}
.badge.blue{background:rgba(59,130,246,0.15);color:#60a5fa;border:1px solid rgba(59,130,246,0.3)}
.badge.purple{background:rgba(139,92,246,0.15);color:#a78bfa;border:1px solid rgba(139,92,246,0.3)}
.badge.cyan{background:rgba(6,182,212,0.15);color:#22d3ee;border:1px solid rgba(6,182,212,0.3)}
.badge.steel{background:rgba(58,79,111,0.3);color:var(--text2);border:1px solid var(--border2)}

/* MILESTONE CARDS */
.ms-card{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:20px;margin-bottom:14px;transition:all .3s;box-shadow:var(--shadow)}
.ms-card:hover{box-shadow:var(--shadow-lg);border-color:var(--border2)}
.ms-card .ms-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px}
.ms-card .ms-header h4{font-size:16px;font-weight:800}
.ms-card .ms-stats{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}
.ms-card .ms-stat{text-align:center;padding:12px 8px;background:var(--bg2);border-radius:10px;border:1px solid var(--border)}
.ms-card .ms-stat .num{font-size:20px;font-weight:800}
.ms-card .ms-stat .lbl{font-size:10px;color:var(--text3);margin-top:2px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px}

/* SEARCH */
.search-wrap{position:relative;margin-bottom:16px}
.search-box{width:100%;padding:14px 20px 14px 48px;background:var(--card);border:2px solid var(--border);border-radius:12px;color:var(--text);font-size:14px;outline:none;transition:all .3s;font-family:inherit}
.search-box:focus{border-color:var(--orange);box-shadow:0 0 0 4px rgba(245,158,11,0.1);background:var(--card2)}
.search-box::placeholder{color:var(--text3)}
.search-icon{position:absolute;left:16px;top:50%;transform:translateY(-50%);color:var(--text3);font-size:18px;transition:color .3s}
.search-box:focus ~ .search-icon{color:var(--orange)}
.search-count{position:absolute;right:16px;top:50%;transform:translateY(-50%);font-size:11px;color:var(--text3);background:var(--card3);padding:4px 10px;border-radius:6px;font-weight:600;transition:all .3s}
.search-active .search-count{color:var(--orange);background:rgba(245,158,11,0.1);border:1px solid rgba(245,158,11,0.3)}

.filter-row{display:flex;gap:6px;margin-bottom:16px;flex-wrap:wrap}
.filter-btn{padding:7px 14px;border-radius:8px;border:1px solid var(--border);background:var(--card);color:var(--text3);font-size:11px;font-weight:700;cursor:pointer;transition:all .25s;text-transform:uppercase;letter-spacing:0.5px}
.filter-btn:hover{background:var(--card2);color:var(--text2);border-color:var(--border2)}
.filter-btn.active{background:var(--orange);color:#000;border-color:var(--orange);box-shadow:0 2px 10px rgba(245,158,11,0.3)}

/* SEARCH HIGHLIGHT */
.search-highlight{background:rgba(245,158,11,0.25);color:var(--orange);padding:1px 4px;border-radius:3px;font-weight:600}

/* ROW COLORS WHEN SEARCHING */
.searching tr.search-match td{background:rgba(245,158,11,0.04);border-bottom-color:rgba(245,158,11,0.15)}
.searching tr.search-match:hover td{background:rgba(245,158,11,0.1)}
.searching tr.search-nomatch td{opacity:0.3}
.searching tr.search-nomatch:hover td{opacity:0.6}

/* HEADER SEARCH STATE */
.header.searching-mode{border-bottom-color:var(--blue)}
.header.searching-mode .meta span{border-color:rgba(59,130,246,0.3);background:rgba(59,130,246,0.1);color:#60a5fa}
</style>
</head>
<body>
<div class="header" id="mainHeader">
  <div class="header-top">
    <h1><span>KENT</span> PLC - Completion Progress Dashboard</h1>
    <div class="header-badge">
      <span class="tag">PS5 - Tanzania</span>
      <span class="live">LIVE DATA</span>
    </div>
  </div>
  <div class="subtitle">Pre-Commissioning Progress Tracker | Pumping Station 5 | Cut-off: 31 July 2026</div>
  <div class="meta">
    <span>Company: KENT</span>
    <span>Tasks: __TOTAL_TASKS__</span>
    <span>Subsystems: __TOTAL_SUBSYS__</span>
    <span>Punch List: __TOTAL_PL__</span>
    <span>Completion: __COMP_PCT__%</span>
  </div>
</div>

<div class="container">
  <div class="tabs">
    <div class="tab active" onclick="showPanel('overview',this)">Overview</div>
    <div class="tab" onclick="showPanel('milestones',this)">Milestones</div>
    <div class="tab" onclick="showPanel('subsystems',this)">Subsystems</div>
    <div class="tab" onclick="showPanel('punchlist',this)">Punch List</div>
    <div class="tab" onclick="showPanel('charts',this)">Analytics</div>
    <div class="tab" onclick="showPanel('systemview',this)">System View</div>
  </div>

  <!-- OVERVIEW -->
  <div id="overview" class="panel active">
    <div class="grid g5" style="margin-bottom:20px">
      <div class="card stat-card orange">
        <h3>Total Tasks</h3>
        <div class="value">__TOTAL_TASKS__</div>
        <div class="label">Across 9 Milestones</div>
      </div>
      <div class="card stat-card green">
        <h3>Closed Tasks</h3>
        <div class="value">__CLOSED_TASKS__</div>
        <div class="label">__COMP_PCT__% Completion</div>
        <div class="progress-bar" style="margin-top:10px"><div class="fill green" style="width:__COMP_PCT_VIS__%"></div></div>
      </div>
      <div class="card stat-card red">
        <h3>Open Tasks</h3>
        <div class="value">__OPEN_TASKS__</div>
        <div class="label">845 Pending + 3 Started</div>
      </div>
      <div class="card stat-card cyan">
        <h3>Subsystems</h3>
        <div class="value">__TOTAL_SUBSYS__</div>
        <div class="label">0 Complete / 123 Active</div>
      </div>
      <div class="card stat-card purple">
        <h3>Punch List</h3>
        <div class="value">__TOTAL_PL__</div>
        <div class="label">__CLOSED_PL__ Closed / __ORIG_PL__ Open</div>
      </div>
    </div>

    <div class="grid g3" style="margin-bottom:20px">
      <div class="card" style="border-left:4px solid #3b82f6">
        <h3>E - Electrical</h3>
        <div style="text-align:center;padding:8px 0">
          <div style="font-size:42px;font-weight:900;color:#3b82f6">538</div>
          <div style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:1px;margin-top:2px">Tasks</div>
          <div style="display:flex;justify-content:center;gap:24px;margin-top:12px">
            <div><span style="font-size:20px;font-weight:800;color:#22c55e">3</span><div style="font-size:10px;color:var(--text3)">CLOSED</div></div>
            <div><span style="font-size:20px;font-weight:800;color:#ef4444">535</span><div style="font-size:10px;color:var(--text3)">OPEN</div></div>
          </div>
          <div class="progress-bar" style="margin-top:12px"><div class="fill blue" style="width:0.6%"></div></div>
        </div>
      </div>
      <div class="card" style="border-left:4px solid #8b5cf6">
        <h3>I - Instrumentation</h3>
        <div style="text-align:center;padding:8px 0">
          <div style="font-size:42px;font-weight:900;color:#8b5cf6">313</div>
          <div style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:1px;margin-top:2px">Tasks</div>
          <div style="display:flex;justify-content:center;gap:24px;margin-top:12px">
            <div><span style="font-size:20px;font-weight:800;color:#22c55e">0</span><div style="font-size:10px;color:var(--text3)">CLOSED</div></div>
            <div><span style="font-size:20px;font-weight:800;color:#ef4444">313</span><div style="font-size:10px;color:var(--text3)">OPEN</div></div>
          </div>
          <div class="progress-bar" style="margin-top:12px"><div class="fill purple" style="width:0%"></div></div>
        </div>
      </div>
      <div class="card" style="border-left:4px solid #f59e0b">
        <h3>Subsystem Punch Summary</h3>
        <div style="padding:4px 0">
          <div style="display:flex;justify-content:space-between;padding:10px 12px;border-bottom:1px solid var(--border)">
            <span style="font-size:12px;font-weight:600">Total Open Punch</span>
            <span style="font-size:18px;font-weight:800;color:var(--orange)">6</span>
          </div>
          <div style="display:flex;justify-content:space-between;padding:10px 12px;border-bottom:1px solid var(--border)">
            <span style="font-size:12px;color:var(--text2)">Category A - Critical</span>
            <span class="badge red">0</span>
          </div>
          <div style="display:flex;justify-content:space-between;padding:10px 12px;border-bottom:1px solid var(--border)">
            <span style="font-size:12px;color:var(--text2)">Category B - Major</span>
            <span class="badge orange">3</span>
          </div>
          <div style="display:flex;justify-content:space-between;padding:10px 12px">
            <span style="font-size:12px;color:var(--text2)">Category C - Minor</span>
            <span class="badge blue">3</span>
          </div>
        </div>
      </div>
    </div>

    <div class="grid g2">
      <div class="card">
        <h3>Task Status Distribution</h3>
        <div class="chart-container"><canvas id="overviewPie"></canvas></div>
      </div>
      <div class="card">
        <h3>Punch List Status</h3>
        <div class="chart-container"><canvas id="plPie"></canvas></div>
      </div>
    </div>
  </div>

  <!-- MILESTONES -->
  <div id="milestones" class="panel">
    <div class="card" style="margin-bottom:20px">
      <h3>Tasks by Milestone</h3>
      <div class="chart-container tall"><canvas id="msBar"></canvas></div>
    </div>
    <div id="milestoneCards"></div>
  </div>

  <!-- SUBSYSTEMS -->
  <div id="subsystems" class="panel">
    <div class="search-wrap" id="subsysSearchWrap">
      <input type="text" class="search-box" id="subsysSearch" placeholder="Search systems, subsystems, milestones..." oninput="filterSubsystems()" onfocus="onSearchFocus()" onblur="onSearchBlur()">
      <div class="search-icon">&#128269;</div>
      <div class="search-count" id="subsysCount">123 results</div>
    </div>
    <div class="filter-row">
      <div class="filter-btn active" onclick="filterMs(this,'all')">All (123)</div>
      <div class="filter-btn" onclick="filterMs(this,'PS5 - Milestone A')">A (2)</div>
      <div class="filter-btn" onclick="filterMs(this,'PS5 - Milestone B')">B (13)</div>
      <div class="filter-btn" onclick="filterMs(this,'PS5 - Milestone C')">C (8)</div>
      <div class="filter-btn" onclick="filterMs(this,'PS5 - Milestone D')">D (11)</div>
      <div class="filter-btn" onclick="filterMs(this,'PS5 - Milestone E')">E (17)</div>
      <div class="filter-btn" onclick="filterMs(this,'PS5 - Milestone F')">F (14)</div>
      <div class="filter-btn" onclick="filterMs(this,'PS5 - Milestone G')">G (15)</div>
      <div class="filter-btn" onclick="filterMs(this,'PS5 - Milestone H')">H (32)</div>
      <div class="filter-btn" onclick="filterMs(this,'PS5 - Milestone I')">I (11)</div>
    </div>
    <div class="table-scroll" id="subsysTableWrap">
      <table id="subsysTable">
        <thead>
          <tr>
            <th>#</th>
            <th>System</th>
            <th>Subsystem</th>
            <th>MILESTONE</th>
            <th style="text-align:right">Total</th>
            <th style="text-align:right">Closed</th>
            <th style="text-align:right">E-Open</th>
            <th style="text-align:right">I-Open</th>
            <th style="text-align:right">%</th>
            <th style="text-align:center">A</th>
            <th style="text-align:center">B</th>
            <th style="text-align:center">C</th>
            <th style="width:130px">Progress</th>
          </tr>
        </thead>
        <tbody id="subsysBody"></tbody>
      </table>
    </div>
  </div>

  <!-- PUNCH LIST -->
  <div id="punchlist" class="panel">
    <div class="grid g4" style="margin-bottom:20px">
      <div class="card stat-card purple">
        <h3>Total Punch Items</h3>
        <div class="value">__TOTAL_PL__</div>
        <div class="label">All Categories Combined</div>
      </div>
      <div class="card stat-card red">
        <h3>Originated (Open)</h3>
        <div class="value">__ORIG_PL__</div>
        <div class="label">__ORIG_PCT__% of Total</div>
      </div>
      <div class="card stat-card green">
        <h3>Closed</h3>
        <div class="value">__CLOSED_PL__</div>
        <div class="label">__CLOSED_PL_PCT__% of Total</div>
      </div>
      <div class="card stat-card blue">
        <h3>Completed</h3>
        <div class="value">__COMPLETED_PL__</div>
        <div class="label">Fully Resolved</div>
      </div>
    </div>
    <div class="grid g2">
      <div class="card">
        <h3>Punch List by Category</h3>
        <div class="chart-container"><canvas id="plCatChart"></canvas></div>
      </div>
      <div class="card">
        <h3>Punch List by Milestone</h3>
        <div class="chart-container"><canvas id="plMsChart"></canvas></div>
      </div>
    </div>
  </div>

  <!-- ANALYTICS -->
  <div id="charts" class="panel">
    <div class="grid g2">
      <div class="card">
        <h3>Tasks by Milestone</h3>
        <div class="chart-container tall"><canvas id="chartTasksMs"></canvas></div>
      </div>
      <div class="card">
        <h3>Completion Radar</h3>
        <div class="chart-container tall"><canvas id="chartCompMs"></canvas></div>
      </div>
      <div class="card">
        <h3>Discipline Split</h3>
        <div class="chart-container"><canvas id="chartDisc"></canvas></div>
      </div>
      <div class="card">
        <h3>Punch by Category per Milestone</h3>
        <div class="chart-container tall"><canvas id="chartPlCross"></canvas></div>
      </div>
    </div>
  </div>

  <!-- SYSTEM VIEW -->
  <div id="systemview" class="panel">
    <div class="card" style="margin-bottom:20px">
      <h3>Progress by System Group</h3>
      <div class="chart-container tall"><canvas id="chartSys"></canvas></div>
    </div>
    <div class="table-scroll">
      <table>
        <thead>
          <tr>
            <th>System</th>
            <th style="text-align:right">Subsystems</th>
            <th style="text-align:right">Total Tasks</th>
            <th style="text-align:right">Closed</th>
            <th style="text-align:right">Open</th>
            <th style="text-align:right">%</th>
            <th style="width:180px">Progress</th>
          </tr>
        </thead>
        <tbody id="sysBody"></tbody>
      </table>
    </div>
  </div>
</div>

<script>
var DATA = __DATA_JSON__;
var searchActive = false;

Chart.defaults.color = '#5a6f8a';
Chart.defaults.borderColor = '#243352';
Chart.defaults.font.family = "'Inter',system-ui,sans-serif";

var msColors = ['#f59e0b','#ef4444','#22c55e','#3b82f6','#8b5cf6','#ec4899','#06b6d4','#f97316','#6366f1'];
var msColorsAlpha = ['rgba(245,158,11,0.7)','rgba(239,68,68,0.7)','rgba(34,197,94,0.7)','rgba(59,130,246,0.7)','rgba(139,92,246,0.7)','rgba(236,72,153,0.7)','rgba(6,182,212,0.7)','rgba(249,115,22,0.7)','rgba(99,102,241,0.7)'];

function showPanel(id, el) {
  document.querySelectorAll('.panel').forEach(function(p){ p.classList.remove('active') });
  document.querySelectorAll('.tab').forEach(function(t){ t.classList.remove('active') });
  document.getElementById(id).classList.add('active');
  el.classList.add('active');
  if (id === 'charts') initCharts();
}

function initOverview() {
  new Chart(document.getElementById('overviewPie'), {
    type: 'doughnut',
    data: {
      labels: ['Closed (3)', 'Started (3)', 'Pending (845)'],
      datasets: [{ data: [3, 3, 845], backgroundColor: ['#22c55e','#f59e0b','#ef4444'], borderWidth: 3, borderColor: '#151f32', hoverOffset: 10 }]
    },
    options: { responsive: true, maintainAspectRatio: false, cutout: '68%', plugins: { legend: { position: 'bottom', labels: { padding: 16, usePointStyle: true, pointStyle: 'circle', color: '#8899b4' } } } }
  });
  new Chart(document.getElementById('plPie'), {
    type: 'doughnut',
    data: {
      labels: ['Originated (1766)', 'Closed (437)', 'Completed (30)'],
      datasets: [{ data: [1766, 437, 30], backgroundColor: ['#ef4444','#22c55e','#3b82f6'], borderWidth: 3, borderColor: '#151f32', hoverOffset: 10 }]
    },
    options: { responsive: true, maintainAspectRatio: false, cutout: '68%', plugins: { legend: { position: 'bottom', labels: { padding: 16, usePointStyle: true, pointStyle: 'circle', color: '#8899b4' } } } }
  });
}

function initMilestones() {
  var sorted = DATA.milestones.slice().sort(function(a,b){ return a.name.localeCompare(b.name) });
  var labels = sorted.map(function(m){ return 'Milestone ' + m.letter });
  new Chart(document.getElementById('msBar'), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        { label: 'Closed', data: sorted.map(function(m){ return m.closedTasks }), backgroundColor: '#22c55e', borderRadius: 6, barPercentage: 0.65 },
        { label: 'Open', data: sorted.map(function(m){ return m.openTasks }), backgroundColor: '#ef4444', borderRadius: 6, barPercentage: 0.65 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: { x: { stacked: true, grid: { display: false }, ticks: { color: '#5a6f8a' } }, y: { stacked: true, beginAtZero: true, grid: { color: '#1e2d48' }, ticks: { color: '#5a6f8a' } } },
      plugins: { legend: { position: 'top', labels: { usePointStyle: true, pointStyle: 'circle', padding: 16, color: '#8899b4' } } }
    }
  });

  var container = document.getElementById('milestoneCards');
  container.innerHTML = '';
  sorted.forEach(function(m, i) {
    var discHtml = m.disciplines.map(function(d) {
      return '<div style="display:flex;justify-content:space-between;padding:8px 12px;border-bottom:1px solid var(--border);font-size:12px">' +
        '<span style="font-weight:500">' + d.name + '</span>' +
        '<span><span style="color:#22c55e;font-weight:700">' + d.closedTasks + '</span> / ' + d.totalTasks + '</span>' +
      '</div>';
    }).join('');

    container.innerHTML += '<div class="ms-card" style="border-left:5px solid ' + msColors[i % msColors.length] + '">' +
      '<div class="ms-header">' +
        '<h4 style="color:' + msColors[i % msColors.length] + '">Milestone ' + m.letter + '</h4>' +
        '<span class="badge ' + (m.completion > 0 ? 'green' : 'red') + '">' + m.completion + '% Complete</span>' +
      '</div>' +
      '<div class="ms-stats">' +
        '<div class="ms-stat"><div class="num" style="color:#f59e0b">' + m.subsystems + '</div><div class="lbl">Subsystems</div></div>' +
        '<div class="ms-stat"><div class="num" style="color:#06b6d4">' + m.totalTasks + '</div><div class="lbl">Total Tasks</div></div>' +
        '<div class="ms-stat"><div class="num" style="color:#22c55e">' + m.closedTasks + '</div><div class="lbl">Closed</div></div>' +
        '<div class="ms-stat"><div class="num" style="color:#ef4444">' + m.openTasks + '</div><div class="lbl">Open</div></div>' +
      '</div>' +
      '<div class="progress-bar" style="margin-top:14px;height:10px"><div class="fill green" style="width:' + Math.max(m.completion, 0.5) + '%"></div></div>' +
      '<div style="margin-top:12px;background:var(--bg2);border-radius:10px;padding:6px 0;border:1px solid var(--border)">' + discHtml + '</div>' +
    '</div>';
  });
}

var currentMsFilter = 'all';
function initSubsystems() { renderSubsystems(); }

function filterMs(el, ms) {
  currentMsFilter = ms;
  document.querySelectorAll('.filter-btn').forEach(function(b){ b.classList.remove('active') });
  el.classList.add('active');
  renderSubsystems();
}

function onSearchFocus() {
  searchActive = true;
  document.getElementById('subsysSearchWrap').classList.add('search-active');
  document.getElementById('mainHeader').classList.add('searching-mode');
}

function onSearchBlur() {
  var val = document.getElementById('subsysSearch').value;
  if (!val) {
    searchActive = false;
    document.getElementById('subsysSearchWrap').classList.remove('search-active');
    document.getElementById('mainHeader').classList.remove('searching-mode');
  }
}

function filterSubsystems() { renderSubsystems(); }

function highlightText(text, query) {
  if (!query) return text;
  var regex = new RegExp('(' + query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'gi');
  return text.replace(regex, '<span class="search-highlight">$1</span>');
}

function renderSubsystems() {
  var query = document.getElementById('subsysSearch').value.toLowerCase();
  var filtered = DATA.subsystems.filter(function(s) {
    var matchSearch = !query || s.system.toLowerCase().indexOf(query) >= 0 || s.subsystem.toLowerCase().indexOf(query) >= 0 || s.milestone.toLowerCase().indexOf(query) >= 0;
    var matchMs = currentMsFilter === 'all' || s.milestone === currentMsFilter;
    return matchSearch && matchMs;
  });

  var isSearching = query.length > 0;
  var tableWrap = document.getElementById('subsysTableWrap');
  var tbody = document.getElementById('subsysBody');

  if (isSearching) {
    tableWrap.classList.add('searching');
  } else {
    tableWrap.classList.remove('searching');
  }

  document.getElementById('subsysCount').textContent = filtered.length + ' results';

  tbody.innerHTML = filtered.map(function(s, idx) {
    var pct = s.total > 0 ? (s.closed / s.total * 100).toFixed(1) : '0.0';
    var fillColor = pct >= 50 ? 'green' : pct > 0 ? 'orange' : 'red';
    var rowClass = isSearching ? 'search-match' : '';
    var systemText = highlightText(s.system, document.getElementById('subsysSearch').value);
    var subsystemText = highlightText(s.subsystem, document.getElementById('subsysSearch').value);

    return '<tr class="' + rowClass + '">' +
      '<td style="color:var(--text3);font-size:10px">' + (idx + 1) + '</td>' +
      '<td style="font-size:11px;max-width:200px;word-break:break-all;color:var(--text2)">' + systemText + '</td>' +
      '<td style="font-size:11px;max-width:260px;word-break:break-all;font-weight:500">' + subsystemText + '</td>' +
      '<td><span class="badge steel">' + s.milestone.replace('PS5 - ','') + '</span></td>' +
      '<td style="text-align:right;font-weight:800;color:var(--text)">' + s.total + '</td>' +
      '<td style="text-align:right;color:#22c55e;font-weight:700">' + s.closed + '</td>' +
      '<td style="text-align:right;color:#3b82f6">' + s.elecOpen + '</td>' +
      '<td style="text-align:right;color:#8b5cf6">' + s.instrOpen + '</td>' +
      '<td style="text-align:right;font-weight:800">' + pct + '%</td>' +
      '<td style="text-align:center">' + (s.punchA > 0 ? '<span class="badge red">' + s.punchA + '</span>' : '<span style="color:var(--text3)">-</span>') + '</td>' +
      '<td style="text-align:center">' + (s.punchB > 0 ? '<span class="badge orange">' + s.punchB + '</span>' : '<span style="color:var(--text3)">-</span>') + '</td>' +
      '<td style="text-align:center">' + (s.punchC > 0 ? '<span class="badge blue">' + s.punchC + '</span>' : '<span style="color:var(--text3)">-</span>') + '</td>' +
      '<td><div class="progress-bar" style="height:6px"><div class="fill ' + fillColor + '" style="width:' + pct + '%"></div></div></td>' +
    '</tr>';
  }).join('');
}

var chartsInitialized = false;
function initCharts() {
  if (chartsInitialized) return;
  chartsInitialized = true;
  var sorted = DATA.milestones.slice().sort(function(a,b){ return a.name.localeCompare(b.name) });
  var labels = sorted.map(function(m){ return 'Ms ' + m.letter });

  new Chart(document.getElementById('chartTasksMs'), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        { label: 'Total Tasks', data: sorted.map(function(m){ return m.totalTasks }), backgroundColor: msColorsAlpha, borderRadius: 6, barPercentage: 0.7 },
        { label: 'Closed', data: sorted.map(function(m){ return m.closedTasks }), backgroundColor: '#22c55e', borderRadius: 6, barPercentage: 0.7 }
      ]
    },
    options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, grid: { color: '#1e2d48' }, ticks: { color: '#5a6f8a' } }, x: { grid: { display: false }, ticks: { color: '#5a6f8a' } } }, plugins: { legend: { labels: { usePointStyle: true, pointStyle: 'circle', color: '#8899b4' } } } }
  });

  new Chart(document.getElementById('chartCompMs'), {
    type: 'radar',
    data: {
      labels: sorted.map(function(m){ return 'Ms ' + m.letter }),
      datasets: [{
        label: 'Completion %',
        data: sorted.map(function(m){ return m.completion }),
        backgroundColor: 'rgba(245,158,11,0.1)',
        borderColor: '#f59e0b',
        pointBackgroundColor: '#f59e0b',
        pointBorderColor: '#151f32',
        pointBorderWidth: 2,
        borderWidth: 2
      }]
    },
    options: { responsive: true, maintainAspectRatio: false, scales: { r: { beginAtZero: true, max: 5, ticks: { stepSize: 1, color: '#5a6f8a', backdropColor: 'transparent' }, grid: { color: '#1e2d48' }, angleLines: { color: '#1e2d48' }, pointLabels: { color: '#8899b4' } } }, plugins: { legend: { display: false } } }
  });

  new Chart(document.getElementById('chartDisc'), {
    type: 'doughnut',
    data: {
      labels: ['E - Electrical (538)', 'I - Instrumentation (313)'],
      datasets: [{ data: [538, 313], backgroundColor: ['#3b82f6', '#8b5cf6'], borderWidth: 3, borderColor: '#151f32', hoverOffset: 10 }]
    },
    options: { responsive: true, maintainAspectRatio: false, cutout: '62%', plugins: { legend: { position: 'bottom', labels: { padding: 16, usePointStyle: true, pointStyle: 'circle', color: '#8899b4' } } } }
  });

  var plLetters = Object.keys(DATA.plCross).sort();
  new Chart(document.getElementById('chartPlCross'), {
    type: 'bar',
    data: {
      labels: plLetters.map(function(l){ return 'Ms ' + l }),
      datasets: [
        { label: 'A - Critical', data: plLetters.map(function(l){ return DATA.plCross[l].A }), backgroundColor: '#ef4444', borderRadius: 4 },
        { label: 'B - Major', data: plLetters.map(function(l){ return DATA.plCross[l].B }), backgroundColor: '#f59e0b', borderRadius: 4 },
        { label: 'C - Minor', data: plLetters.map(function(l){ return DATA.plCross[l].C }), backgroundColor: '#06b6d4', borderRadius: 4 }
      ]
    },
    options: { responsive: true, maintainAspectRatio: false, scales: { x: { stacked: true, grid: { display: false }, ticks: { color: '#5a6f8a' } }, y: { stacked: true, beginAtZero: true, grid: { color: '#1e2d48' }, ticks: { color: '#5a6f8a' } } }, plugins: { legend: { position: 'top', labels: { usePointStyle: true, pointStyle: 'circle', padding: 16, color: '#8899b4' } } } }
  });
}

function initSysTable() {
  var tbody = document.getElementById('sysBody');
  tbody.innerHTML = DATA.sysSummary.map(function(s) {
    var pctColor = s.pct >= 50 ? 'green' : s.pct > 0 ? 'orange' : 'red';
    var open = s.total - s.closed;
    return '<tr>' +
      '<td style="font-weight:600;font-size:12px">' + s.system + '</td>' +
      '<td style="text-align:right;color:var(--text2)">' + s.count + '</td>' +
      '<td style="text-align:right;font-weight:800;color:var(--text)">' + s.total + '</td>' +
      '<td style="text-align:right;color:#22c55e;font-weight:700">' + s.closed + '</td>' +
      '<td style="text-align:right;color:#ef4444;font-weight:700">' + open + '</td>' +
      '<td style="text-align:right;font-weight:800;color:var(--text)">' + s.pct + '%</td>' +
      '<td><div class="progress-bar" style="height:8px"><div class="fill ' + pctColor + '" style="width:' + s.pct + '%"></div></div></td>' +
    '</tr>';
  }).join('');

  new Chart(document.getElementById('chartSys'), {
    type: 'bar',
    data: {
      labels: DATA.sysSummary.map(function(s){ return s.system.replace('PS5-','') }),
      datasets: [{
        label: 'Total Tasks',
        data: DATA.sysSummary.map(function(s){ return s.total }),
        backgroundColor: DATA.sysSummary.map(function(s,i){ return msColors[i % msColors.length] + 'cc' }),
        borderRadius: 6
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      scales: { x: { beginAtZero: true, grid: { color: '#1e2d48' }, ticks: { color: '#5a6f8a' } }, y: { grid: { display: false }, ticks: { color: '#8899b4', font: { size: 11 } } } },
      plugins: { legend: { display: false } }
    }
  });
}

initOverview();
initMilestones();
initSubsystems();
initSysTable();
</script>
</body>
</html>'''


def generate_html(data):
    data_json = json.dumps(data, default=str)
    comp_pct = round(data["closedTasks"] / data["totalTasks"] * 100, 2)

    html = HTML_TEMPLATE
    html = html.replace('__DATA_JSON__', data_json)
    html = html.replace('__TOTAL_TASKS__', str(data["totalTasks"]))
    html = html.replace('__CLOSED_TASKS__', str(data["closedTasks"]))
    html = html.replace('__OPEN_TASKS__', str(data["openTasks"]))
    html = html.replace('__TOTAL_SUBSYS__', str(data["totalSubsystems"]))
    html = html.replace('__COMP_SUBSYS__', str(data["completeSubsystems"]))
    html = html.replace('__COMP_PCT__', str(comp_pct))
    html = html.replace('__COMP_PCT_VIS__', str(max(comp_pct, 2)))
    html = html.replace('__TOTAL_PL__', str(data["totalPL"]))
    html = html.replace('__CLOSED_PL__', str(data["closedPL"]))
    html = html.replace('__ORIG_PL__', str(data["originatedPL"]))
    html = html.replace('__COMPLETED_PL__', str(data["completedPL"]))
    html = html.replace('__ORIG_PCT__', str(round(data["originatedPL"] / data["totalPL"] * 100, 1)))
    html = html.replace('__CLOSED_PL_PCT__', str(round(data["closedPL"] / data["totalPL"] * 100, 1)))
    return html


if __name__ == '__main__':
    print("Extracting data from Excel...")
    data = extract_data()
    print(f"Milestones: {len(data['milestones'])} | Subsystems: {len(data['subsystems'])}")
    print(f"Tasks: {data['totalTasks']} (Closed: {data['closedTasks']}, Open: {data['openTasks']})")
    print(f"PL: {data['totalPL']} (Originated: {data['originatedPL']}, Closed: {data['closedPL']}, Completed: {data['completedPL']})")
    print("Generating HTML...")
    html = generate_html(data)
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Saved: {OUTPUT_PATH}")
