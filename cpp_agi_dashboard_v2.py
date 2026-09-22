"""
====================================================================
 CPP AGI / PS5 - EACOP Project Dashboard v2.0
 Master Index: Task ID + Asset Tag (Single Source of Truth)
 Sources:
   1. ovTasks_TestsPlanned*.xlsx     -> Base (Task ID, Asset Tag, State...)
   2. *COMPLETIONS*DPR*.xlsx         -> EIT Date, RFC Date, Month, Milestone
   3. *INSPECTION*REGISTER*.xlsx     -> RFI No, Status, Inspection Date, Type
   4. *PUNCH*LIST*.xlsx              -> PL ID, Category, Status, Raised Date
====================================================================
"""
import os, re, json, webbrowser, shutil
try:
    import pandas as pd
except ImportError:
    os.system("pip install pandas openpyxl --break-system-packages")
    import pandas as pd

DOWNLOADS  = r"C:\Users\mylap\OneDrive\Desktop\dashboard"
OUTPUT_HTML  = os.path.join(DOWNLOADS, "index.html")
OUTPUT_HTML2 = os.path.join(DOWNLOADS, "PS5_Project_Dashboard.html")

DISC_NORM = {
    'I - Instrumentation': 'I - Instrumentation',
    'I - Instrument':      'I - Instrumentation',
    'Instrumentation':     'I - Instrumentation',
    'Instrument':          'I - Instrumentation',
    'E - Electrical':      'E - Electrical',
    'Electrical':          'E - Electrical',
    'T - Telecom':         'T - Telecom',
    'Telecom':             'T - Telecom',
}
def norm_disc(v):
    if not isinstance(v,str): return ''
    return DISC_NORM.get(v.strip(), v.strip())

def norm_rfi(v):
    if not v or str(v).strip().lower() in ('nan','none',''): return ''
    return re.sub(r'[\s/]+','-', str(v).strip().upper())

def n(v):
    return str(v).strip() if pd.notna(v) else ''

def nd(v):
    try: return pd.to_datetime(v).strftime('%Y-%m-%d')
    except: return ''

def find_file(keywords):
    for f in sorted(os.listdir(DOWNLOADS)):
        if f.startswith('~$'): continue
        low = f.lower()
        if low.endswith('.xlsx') and all(k.lower() in low for k in keywords):
            return os.path.join(DOWNLOADS, f)
    return None

# ====================================================================
#  MASTER INDEX
# ====================================================================
def build_master_index(ov_path, dpr_path, rfi_path, punch_path):
    print("="*60)
    print("  Building Master Index (No data dropped)")
    print("="*60)

    # 1. ovTasks base
    print("\n[1] ovTasks...")
    ov = pd.read_excel(ov_path, sheet_name='Exported from SC')
    idx = {}
    for _, r in ov.iterrows():
        tid = n(r['Task ID']); atag = n(r['Asset - Tag'])
        if not tid: continue
        idx[tid] = {
            'tid': tid, 'at': atag,
            'disc': norm_disc(n(r['Discipline (Summary)'])),
            'st':   n(r['Task State']),
            'ty':   n(r['Task Type (Name)']),
            'desc': n(r['Description']),
            'sub':  n(r['Systemization - Subsystem (Summary)']),
            'ms':   n(r['Subsystem Priority']),
            'cd':   nd(r['Closing Date']),
            'cat':  n(r['Category (Summary)']),
            'res':  n(r['Responsible Company (Summary)']),
            'eit':  '', 'rfc': '', 'mon': '',
            'rfi': [], 'punch': [],
        }
    # asset->tids map
    at2t = {}
    for tid,rec in idx.items():
        if rec['at']:
            at2t.setdefault(rec['at'],[])
            if tid not in at2t[rec['at']]: at2t[rec['at']].append(tid)
    print(f"  {len(idx)} tasks, {len(at2t)} unique asset tags")

    # 2. DPR
    if dpr_path:
        print("\n[2] DPR Detailed ITR List...")
        dpr = pd.read_excel(dpr_path, sheet_name='DETAILED ITR LIST')
        m=0
        for _,r in dpr.iterrows():
            tid = n(r['Task ID'])
            if tid in idx:
                idx[tid]['eit'] = nd(r.get('EIT'))
                idx[tid]['rfc'] = nd(r.get('RFC'))
                idx[tid]['mon'] = n(r.get('MONTH'))
                if r.get('Priority.2') and not idx[tid]['ms']:
                    idx[tid]['ms'] = n(r.get('Priority.2'))
                m+=1
        print(f"  {m} enriched with EIT/RFC dates")

    # 3. Inspection Register
    rfi2t = {}
    if rfi_path:
        print("\n[3] Inspection Register...")
        rdf = pd.read_excel(rfi_path, sheet_name='PS-5 EIT INSPECTION REGISTER', header=4)

        def ns(v):
            if not isinstance(v,str): return 'No RFI Yet'
            v=v.strip().lower()
            if 'punch' in v: return 'Accepted with Punch'
            if 'accept' in v: return 'Accepted'
            if 'hold' in v: return 'Hold by EACOP'
            if 'reject' in v: return 'Rejected'
            if 'open' in v: return 'Open'
            return v.title()

        TCOL='INSPECTION TYPE (PULLING/GLANDING & TERMINATION/CABLE TESTING/INSTALLATION)'
        m_tid=m_at=m_no=0
        for _,r in rdf.iterrows():
            tid = n(r.get('Task ID','')); atag = n(r.get('Asset - Tag',''))
            rno = n(r.get('QC RFI No','')); rnorm = norm_rfi(rno)
            rec = {
                'no':  rnorm, 'raw': rno,
                'st':  ns(r.get('Status Of RFI ')),
                'dt':  nd(r.get('Inspection Date')),
                'idt': nd(r.get('Internal Inspection date')),
                'ist': n(r.get('Internal Inspection status (Ready for RFI or not) ')),
                'pcs': n(r.get('Status of Precomm ITR ')),
                'ty':  n(r.get(TCOL,''))[:25],
                'ins': n(r.get('QC INSPECTOR ','')),
                'rfcr':nd(r.get('RFC DATE ','')),
                'rem': n(r.get('Remarks',''))[:60],
            }
            if tid and tid in idx:
                idx[tid]['rfi'].append(rec)
                if rnorm: rfi2t.setdefault(rnorm,[])
                if rnorm and tid not in rfi2t[rnorm]: rfi2t[rnorm].append(tid)
                m_tid+=1
            elif atag and atag in at2t:
                for t in at2t[atag]:
                    idx[t]['rfi'].append(rec)
                    if rnorm: rfi2t.setdefault(rnorm,[])
                    if rnorm and t not in rfi2t[rnorm]: rfi2t[rnorm].append(t)
                m_at+=1
            else: m_no+=1
        print(f"  Matched by Task ID:{m_tid} Asset Tag:{m_at} Unmatched:{m_no}")

    # 4. Punch List
    if punch_path:
        print("\n[4] Punch List...")
        raw = pd.read_excel(punch_path, sheet_name='MP Register', header=None, skiprows=6)
        raw.columns = range(len(raw.columns))
        punch = raw[[1,3,5,8,9,10,13,14,20]].rename(columns={
            1:'plid',3:'rfi_no',5:'sub',8:'desc',9:'cat',
            10:'disc',13:'rdt',14:'resp',20:'st'})
        punch = punch.dropna(subset=['plid'])
        m=0; no=0
        for _,r in punch.iterrows():
            rn = norm_rfi(n(r.get('rfi_no','')))
            prec = {
                'id':  n(r['plid']),
                'cat': n(r.get('cat','')),
                'd':   norm_disc(n(r.get('disc',''))),
                'st':  n(r.get('st','')) or 'Open',
                'dt':  nd(r.get('rdt')),
                'rfi': rn,
                'sub': n(r.get('sub','')),
            }
            tids = rfi2t.get(rn,[])
            if tids:
                for tid in tids:
                    ep = [p['id'] for p in idx[tid]['punch']]
                    if prec['id'] not in ep:
                        idx[tid]['punch'].append(prec)
                m+=1
            else: no+=1
        print(f"  Matched:{m} Unmatched:{no}")

    # Build maps
    print("\n[5] Building search maps...")
    asset_map = at2t
    sub_map = {}
    for tid,rec in idx.items():
        if rec['sub']:
            sub_map.setdefault(rec['sub'],[])
            if tid not in sub_map[rec['sub']]: sub_map[rec['sub']].append(tid)
    pri_map = {}
    for tid,rec in idx.items():
        if rec['ms']:
            pri_map.setdefault(rec['ms'],[])
            if tid not in pri_map[rec['ms']]: pri_map[rec['ms']].append(tid)

    stats = {
        'total':      len(idx),
        'closed':     sum(1 for r in idx.values() if r['st']=='Closed'),
        'with_rfi':   sum(1 for r in idx.values() if r['rfi']),
        'with_punch': sum(1 for r in idx.values() if r['punch']),
        'assets':     len(asset_map),
        'rfis':       len(rfi2t),
        'subsystems': len(sub_map),
        'milestones': len(pri_map),
    }
    print(f"\n{'='*60}")
    for k,v in stats.items(): print(f"  {k:20}: {v}")
    print(f"{'='*60}")

    return {
        'index': idx, 'asset_map': asset_map,
        'rfi_map': rfi2t, 'sub_map': sub_map,
        'pri_map': pri_map, 'stats': stats,
    }

# ====================================================================
#  DASHBOARD DATA (charts/KPIs from ovTasks)
# ====================================================================
def build_dashboard_data(idx, punch_path, rfi_path):
    """بيانات الرسومات والـ KPIs"""
    import pandas as pd

    # use full key names from build_master_index
    # map: task_id,asset_tag,discipline,state,task_type,description,
    #      subsystem,milestone,closing_date,eit_date,rfc_date,month,
    #      category,responsible,rfi[],punch[]
    total    = len(idx)
    closed_n = sum(1 for r in idx.values() if r.get('state')=='Closed')

    disc_total={}; disc_closed={}
    for r in idx.values():
        d = norm_disc(r.get('discipline','')) or 'Other'
        disc_total[d] = disc_total.get(d,0)+1
        if r.get('state')=='Closed': disc_closed[d]=disc_closed.get(d,0)+1

    ms_total={}; ms_closed={}
    for r in idx.values():
        m = r.get('milestone','') or 'Unknown'
        ms_total[m] = ms_total.get(m,0)+1
        if r.get('state')=='Closed': ms_closed[m]=ms_closed.get(m,0)+1

    import collections
    weekly  = collections.defaultdict(lambda:{'E':0,'I':0,'T':0,'Total':0})
    monthly = collections.defaultdict(lambda:{'E':0,'I':0,'T':0,'Total':0})
    for r in idx.values():
        if r.get('state')!='Closed': continue
        cd = r.get('closing_date','')
        if not cd: continue
        dt = pd.to_datetime(cd)
        wk = (dt-pd.Timedelta(days=dt.weekday())).strftime('%Y-%m-%d')
        mo = dt.strftime('%Y-%m')
        d  = norm_disc(r.get('discipline',''))[:1] or 'O'
        for grp,key in [(weekly,wk),(monthly,mo)]:
            grp[key][d]      = grp[key].get(d,0)+1
            grp[key]['Total'] += 1

    weekly_list  = sorted([{'label':k,**v} for k,v in weekly.items()],  key=lambda x:x['label'])
    monthly_list = sorted([{'label':k,**v} for k,v in monthly.items()], key=lambda x:x['label'])

    sub_remain={}
    for r in idx.values():
        sub = r.get('subsystem','') or ''
        if sub:
            sub_remain.setdefault(sub,{'total':0,'closed':0,'remain':0})
            sub_remain[sub]['total']+=1
            if r.get('state')=='Closed': sub_remain[sub]['closed']+=1
    for s,v in sub_remain.items():
        v['remain']=v['total']-v['closed']
        v['pct']=round(v['closed']/v['total']*100,1) if v['total'] else 0
    top_sub_remain = sorted(sub_remain.items(), key=lambda x:x[1]['remain'],reverse=True)[:20]
    top_sub_remain = [{'label':k,**v} for k,v in top_sub_remain]

    punch_status={}; punch_disc={}; punch_sub={}
    for r in idx.values():
        for p in r.get('punch',[]):
            st=p.get('st','') or p.get('status','Open')
            punch_status[st]=punch_status.get(st,0)+1
            d =p.get('d','') or p.get('disc','')[:1] if p.get('disc') else 'O'
            punch_disc[d]=punch_disc.get(d,0)+1
            s =p.get('sub','') or r.get('subsystem','') or 'Unknown'
            punch_sub[s]=punch_sub.get(s,0)+1

    top_punch_sub = sorted(punch_sub.items(),key=lambda x:x[1],reverse=True)[:15]
    top_punch_sub = [{'label':k,'count':v} for k,v in top_punch_sub]

    rfi_status={}; rfi_disc={}
    rfi_daily=collections.defaultdict(int)
    rfi_monthly=collections.defaultdict(int)
    for r in idx.values():
        d=norm_disc(r.get('discipline',''))[:1] or 'O'
        for rf in r.get('rfi',[]):
            st=rf.get('st','') or rf.get('status','')
            rfi_status[st]=rfi_status.get(st,0)+1
            rfi_disc[d]=rfi_disc.get(d,0)+1
            dt=rf.get('dt','') or rf.get('insp_date','')
            if dt:
                rfi_daily[dt]+=1
                rfi_monthly[dt[:7]]+=1

    rfi_daily_list   = sorted([{'label':k,'count':v} for k,v in rfi_daily.items()],  key=lambda x:x['label'])[-30:]
    rfi_monthly_list = sorted([{'label':k,'count':v} for k,v in rfi_monthly.items()],key=lambda x:x['label'])

    return {
        'total':total,'closed':closed_n,
        'pct':round(closed_n/total*100,1) if total else 0,
        'disc_total': {k:int(v) for k,v in disc_total.items()},
        'disc_closed':{k:int(v) for k,v in disc_closed.items()},
        'ms_total':   {k:int(v) for k,v in ms_total.items()},
        'ms_closed':  {k:int(v) for k,v in ms_closed.items()},
        'weekly':weekly_list,'monthly':monthly_list,
        'top_sub_remain':top_sub_remain,
        'punch_status':punch_status,
        'punch_disc':{k:int(v) for k,v in punch_disc.items()},
        'top_punch_sub':top_punch_sub,
        'rfi_status':rfi_status,
        'rfi_disc':{k:int(v) for k,v in rfi_disc.items()},
        'rfi_daily':rfi_daily_list,'rfi_monthly':rfi_monthly_list,
    }

# ====================================================================
#  HTML
# ====================================================================
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="900">
<title>PS5 EACOP Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
:root{
  --bg:#0b1120;--panel:#131a2b;--panel2:#1a2236;
  --accent:#ff4d8d;--gold:#ffb627;--teal:#00e0c6;
  --blue:#3da8ff;--purple:#9b59f6;--green:#00e676;
  --text:#e8ecf7;--muted:#8a93ad;--border:#232c44;
}
*{box-sizing:border-box;font-family:'Segoe UI',Arial,sans-serif;margin:0;padding:0;}
body{background:var(--bg);color:var(--text);display:flex;min-height:100vh;}

.sidebar{
  width:230px;min-height:100vh;background:#0e1626;
  padding:18px 14px;position:sticky;top:0;height:100vh;overflow-y:auto;
  border-right:1px solid var(--border);flex-shrink:0;
}
.sidebar h2{font-size:15px;color:var(--teal);margin-bottom:14px;}
.sidebar .grp{margin-top:16px;font-size:10px;color:var(--gold);text-transform:uppercase;
  letter-spacing:1.5px;border-bottom:1px solid var(--border);padding-bottom:3px;}
.sidebar label{display:flex;align-items:center;gap:7px;padding:7px 6px;border-radius:7px;
  cursor:pointer;font-size:12px;margin-bottom:2px;color:var(--muted);transition:.15s;}
.sidebar label:hover{background:rgba(255,255,255,.07);color:#fff;}
.sidebar input{accent-color:var(--teal);width:14px;height:14px;}

.main{flex:1;padding:20px 28px;overflow:auto;}
.header{background:linear-gradient(120deg,#161f38,#0e1626);
  padding:20px 26px;border-radius:14px;margin-bottom:20px;
  display:flex;justify-content:space-between;align-items:center;
  border:1px solid var(--border);}
.header h1{font-size:22px;font-weight:800;}
.header .sub{font-size:11px;color:var(--muted);margin-top:3px;}

.section{display:none;}
.section.active{display:block;}
.sec-title{font-size:17px;font-weight:800;color:#fff;margin:6px 0 14px;
  border-left:5px solid var(--accent);padding-left:10px;}

.krow{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:18px;}
.kpi{flex:1;min-width:150px;background:var(--panel);border-radius:12px;padding:18px;
  text-align:center;border:1px solid var(--border);border-top:4px solid var(--purple);
  transition:transform .2s;}
.kpi:hover{transform:translateY(-3px);}
.kpi.gold{border-top-color:var(--gold);}
.kpi.teal{border-top-color:var(--teal);}
.kpi.pink{border-top-color:var(--accent);}
.kpi.blue{border-top-color:var(--blue);}
.kpi.green{border-top-color:var(--green);}
.kpi .val{font-size:28px;font-weight:800;color:#fff;}
.kpi .lbl{font-size:11px;color:var(--muted);margin-top:3px;}
.pbar{height:8px;background:#0e1626;border-radius:4px;overflow:hidden;margin-top:8px;}
.pfill{height:100%;background:linear-gradient(90deg,var(--blue),var(--teal));}

.crow{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:18px;}
.cc{flex:1;min-width:320px;background:var(--panel);border-radius:12px;padding:18px;
  border:1px solid var(--border);}
.cc h3{font-size:14px;color:var(--teal);margin-bottom:10px;font-weight:700;}
canvas{max-height:280px;}

/* Search */
#searchInput{width:100%;padding:12px 16px;border:1px solid var(--border);
  border-radius:8px;font-size:14px;background:var(--panel2);color:#fff;margin-bottom:8px;}
#searchInput::placeholder{color:var(--muted);}
.stabs{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;}
.stab{padding:6px 14px;border:1px solid var(--border);border-radius:20px;
  background:transparent;color:var(--muted);font-size:12px;cursor:pointer;transition:.15s;}
.stab.active,.stab:hover{background:var(--teal);color:#000;border-color:var(--teal);}
.result-card{border:1px solid var(--border);border-radius:10px;padding:14px;
  margin-bottom:12px;background:var(--panel2);}
.result-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;}
.result-title{font-size:15px;font-weight:800;color:var(--teal);}
.result-meta{font-size:11px;color:var(--muted);margin-top:2px;}
.rtable{width:100%;border-collapse:collapse;font-size:12px;margin-top:6px;}
.rtable th{color:var(--teal);padding:5px;text-align:left;border-bottom:1px solid var(--border);}
.rtable td{padding:5px;border-bottom:1px solid var(--border);}
.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;}
.no-data{color:var(--muted);font-size:12px;font-style:italic;}
.footer{text-align:center;padding:14px;color:var(--muted);font-size:11px;}
table{color:var(--text);}
</style>
</head>
<body>

<div class="sidebar">
  <h2>🛢️ PS5 EACOP</h2>
  <div class="grp">🔍 Search</div>
  <label><input type="checkbox" data-target="sec-search" checked> Universal Search</label>
  <div class="grp">Overview</div>
  <label><input type="checkbox" data-target="sec-kpi" checked> Project KPIs</label>
  <label><input type="checkbox" data-target="sec-disc" checked> By Discipline (E/I/T)</label>
  <label><input type="checkbox" data-target="sec-ms" checked> By Milestone</label>
  <div class="grp">ITR Closures</div>
  <label><input type="checkbox" data-target="sec-weekly" checked> Weekly</label>
  <label><input type="checkbox" data-target="sec-monthly" checked> Monthly</label>
  <label><input type="checkbox" data-target="sec-backlog" checked> Subsystem Backlog</label>
  <div class="grp">RFI</div>
  <label><input type="checkbox" data-target="sec-rfi" checked> RFI Summary</label>
  <div class="grp">Punch List</div>
  <label><input type="checkbox" data-target="sec-punch" checked> Punch Summary</label>
</div>

<div class="main">
  <div class="header">
    <div>
      <h1>📊 PS5 — EACOP Project Dashboard</h1>
      <div class="sub">CPP AGI | E&I&T | Updated: __NOW__ | Tasks: __TOTAL__ | Closed: __CLOSED__ (__PCT__%)</div>
    </div>
    <div style="font-size:36px;">⚙️</div>
  </div>

  <!-- SEARCH -->
  <div class="section active" id="sec-search">
    <div class="sec-title">🔍 Universal Search</div>
    <div class="cc">
      <input id="searchInput" type="text"
        placeholder="Search by: Asset Tag / Task ID / RFI No / Subsystem / Milestone Priority...">
      <div style="font-size:11px;color:var(--muted);margin-bottom:10px;">
        Index: <b style="color:var(--teal)">__TOTAL__ tasks</b> |
        <b style="color:var(--green)">__ASSETS__ asset tags</b> |
        <b style="color:var(--blue)">__RFIS__ RFI numbers</b> |
        <b style="color:var(--gold)">__SUBS__ subsystems</b>
      </div>
      <div class="stabs">
        <button class="stab active" onclick="setMode('all',this)">All</button>
        <button class="stab" onclick="setMode('task',this)">Task ID</button>
        <button class="stab" onclick="setMode('asset',this)">Asset Tag</button>
        <button class="stab" onclick="setMode('rfi',this)">RFI No</button>
        <button class="stab" onclick="setMode('sub',this)">Subsystem</button>
        <button class="stab" onclick="setMode('ms',this)">Milestone</button>
      </div>
      <div id="searchResults">
        <div style="color:var(--muted);padding:20px;text-align:center;">
          Type at least 3 characters to search...
        </div>
      </div>
    </div>
  </div>

  <!-- KPIs -->
  <div class="section active" id="sec-kpi">
    <div class="sec-title">📊 Project Overview — KPIs</div>
    <div class="krow">
      <div class="kpi"><div class="val">__TOTAL__</div><div class="lbl">Total Tasks</div></div>
      <div class="kpi gold"><div class="val">__CLOSED__</div><div class="lbl">ITRs Closed</div>
        <div class="pbar"><div class="pfill" style="width:__PCT__%"></div></div></div>
      <div class="kpi teal"><div class="val">__OPEN__</div><div class="lbl">Remaining</div></div>
      <div class="kpi pink"><div class="val">__PCT__%</div><div class="lbl">Completion %</div></div>
      <div class="kpi blue"><div class="val">__RFIS__</div><div class="lbl">RFIs Submitted</div></div>
      <div class="kpi green"><div class="val">__PUNCH__</div><div class="lbl">Punch Items</div></div>
    </div>
  </div>

  <!-- DISCIPLINE -->
  <div class="section active" id="sec-disc">
    <div class="sec-title">⚡ By Discipline — E / I / T</div>
    <div class="krow" id="discKpis"></div>
    <div class="crow">
      <div class="cc" style="flex:2;"><h3>Closed vs Total by Discipline</h3><canvas id="cDisc"></canvas></div>
      <div class="cc"><h3>Completion % by Discipline</h3><canvas id="cDiscPct"></canvas></div>
    </div>
  </div>

  <!-- MILESTONE -->
  <div class="section active" id="sec-ms">
    <div class="sec-title">🎯 By Milestone / Subsystem Priority</div>
    <div class="crow">
      <div class="cc" style="flex:2;"><h3>Milestone Progress (Closed vs Total)</h3><canvas id="cMs"></canvas></div>
      <div class="cc"><h3>Completion % by Milestone</h3><canvas id="cMsPct"></canvas></div>
    </div>
  </div>

  <!-- WEEKLY -->
  <div class="section active" id="sec-weekly">
    <div class="sec-title">🗓️ ITR Closures — Weekly</div>
    <div class="crow"><div class="cc" style="flex:2;"><canvas id="cWeekly"></canvas></div></div>
  </div>

  <!-- MONTHLY -->
  <div class="section active" id="sec-monthly">
    <div class="sec-title">📆 ITR Closures — Monthly</div>
    <div class="crow"><div class="cc" style="flex:2;"><canvas id="cMonthly"></canvas></div></div>
  </div>

  <!-- BACKLOG -->
  <div class="section active" id="sec-backlog">
    <div class="sec-title">🏗️ Subsystem Backlog — Top 20 Remaining</div>
    <div class="crow"><div class="cc" style="flex:2;"><canvas id="cBacklog" style="max-height:500px;"></canvas></div></div>
  </div>

  <!-- RFI -->
  <div class="section active" id="sec-rfi">
    <div class="sec-title">📝 RFI Summary</div>
    <div class="crow">
      <div class="cc"><h3>RFI Status</h3><canvas id="cRfiStatus"></canvas></div>
      <div class="cc"><h3>RFI by Discipline</h3><canvas id="cRfiDisc"></canvas></div>
    </div>
    <div class="crow">
      <div class="cc" style="flex:2;"><h3>RFI Submitted — Daily (Last 30 Days)</h3><canvas id="cRfiDaily"></canvas></div>
      <div class="cc" style="flex:2;"><h3>RFI Submitted — Monthly</h3><canvas id="cRfiMonthly"></canvas></div>
    </div>
  </div>

  <!-- PUNCH -->
  <div class="section active" id="sec-punch">
    <div class="sec-title">📌 Punch List Summary</div>
    <div class="crow">
      <div class="cc"><h3>Punch Status</h3><canvas id="cPunchStatus"></canvas></div>
      <div class="cc"><h3>Punch by Discipline</h3><canvas id="cPunchDisc"></canvas></div>
    </div>
    <div class="crow">
      <div class="cc" style="flex:2;"><h3>Top 15 Subsystems (Punch Items)</h3><canvas id="cPunchSub"></canvas></div>
    </div>
  </div>

  <div class="footer">Auto-generated — CPP AGI / EACOP PS5 | __NOW__</div>
</div>

<script>
const MASTER  = __MASTER_JSON__;
const DASH    = __DASH_JSON__;
const IDX     = MASTER.index;
const AMAP    = MASTER.asset_map;
const RMAP    = MASTER.rfi_map;
const SMAP    = MASTER.sub_map;
const PMAP    = MASTER.pri_map;

// ---- Sidebar toggle ----
document.querySelectorAll('.sidebar input').forEach(cb=>{
  cb.addEventListener('change',()=>{
    const el=document.getElementById(cb.dataset.target);
    cb.checked ? el.classList.add('active') : el.classList.remove('active');
  });
});

// ---- Chart defaults ----
Chart.register(ChartDataLabels);
Chart.defaults.color='#8a93ad';
Chart.defaults.borderColor='#232c44';
Chart.defaults.plugins.datalabels.display=false;
const P=['#9b59f6','#00e0c6','#ff4d8d','#ffb627','#3da8ff','#00e676','#ff8a3d'];
const DL={display:true,color:'#000',backgroundColor:'rgba(255,255,255,.88)',
  borderRadius:3,padding:{top:2,bottom:2,left:4,right:4},
  anchor:'end',align:'top',offset:3,font:{weight:'bold',size:10},
  formatter:v=>v>0?v:''};

function bar(id,labels,datasets,opts={}){
  new Chart(document.getElementById(id),{type:'bar',
    data:{labels,datasets},
    options:{plugins:{legend:{position:'bottom'},datalabels:{display:false}},
      scales:{y:{beginAtZero:true}}, ...opts}});
}
function pie(id,labels,data,type='doughnut'){
  new Chart(document.getElementById(id),{type,
    data:{labels,datasets:[{data,backgroundColor:P,
      datalabels:{display:true,color:'#fff',
        font:{weight:'bold',size:11},
        formatter:(v,ctx)=>{
          const t=ctx.dataset.data.reduce((a,b)=>a+b,0);
          return v>0?`${v}\n(${Math.round(v/t*100)}%)`:'';
        }}}]},
    options:{plugins:{legend:{position:'bottom'}}}});
}
function hbar(id,labels,data,color,lab){
  new Chart(document.getElementById(id),{type:'bar',
    data:{labels,datasets:[{label:lab,data,backgroundColor:color,borderRadius:5,
      datalabels:{...DL,anchor:'end',align:'right'}}]},
    options:{indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{beginAtZero:true}}}});
}
function multi(id,rows,type='bar'){
  const labels=rows.map(r=>r.label);
  const S=[{k:'E',n:'Electrical',c:'#9b59f6'},{k:'I',n:'Instrumentation',c:'#00e0c6'},{k:'T',n:'Telecom',c:'#ff4d8d'}];
  new Chart(document.getElementById(id),{type,
    data:{labels,datasets:S.map(s=>({label:s.n,data:rows.map(r=>r[s.k]||0),
      backgroundColor:type==='line'?s.c+'33':s.c,borderColor:s.c,
      fill:type==='line',tension:.3,borderRadius:type==='bar'?5:0,
      datalabels:DL}))},
    options:{plugins:{legend:{position:'bottom'}},scales:{y:{beginAtZero:true}}}});
}

// ---- KPIs ----
(function(){
  const disc_labels=Object.keys(DASH.disc_total);
  const c=document.getElementById('discKpis');
  disc_labels.forEach((d,i)=>{
    const tot=DASH.disc_total[d]||0;
    const cls=DASH.disc_closed[d]||0;
    const pct=tot?Math.round(cls/tot*100):0;
    const colors=['gold','teal','pink'];
    const div=document.createElement('div');
    div.className='kpi '+colors[i%colors.length];
    div.innerHTML=`<div class="val">${cls}/${tot}</div><div class="lbl">${d} — ${pct}%</div>
      <div class="pbar"><div class="pfill" style="width:${pct}%"></div></div>`;
    c.appendChild(div);
  });
})();

// ---- Discipline chart ----
(function(){
  const labs=Object.keys(DASH.disc_total);
  bar('cDisc',labs,[
    {label:'Closed',data:labs.map(k=>DASH.disc_closed[k]||0),backgroundColor:'#00e0c6',datalabels:DL},
    {label:'Total', data:labs.map(k=>DASH.disc_total[k]||0), backgroundColor:'#9b59f620',datalabels:{display:false}},
  ]);
  pie('cDiscPct',labs,labs.map(k=>{
    const t=DASH.disc_total[k]||0,c=DASH.disc_closed[k]||0;
    return t?Math.round(c/t*100):0;
  }),'bar');
})();

// ---- Milestone chart ----
(function(){
  const labs=Object.keys(DASH.ms_total).sort();
  bar('cMs',labs,[
    {label:'Closed',data:labs.map(k=>DASH.ms_closed[k]||0),backgroundColor:'#ffb627',datalabels:DL},
    {label:'Total', data:labs.map(k=>DASH.ms_total[k]||0), backgroundColor:'#ffb62720',datalabels:{display:false}},
  ]);
  bar('cMsPct',labs,[{
    label:'%',data:labs.map(k=>{
      const t=DASH.ms_total[k]||0,c=DASH.ms_closed[k]||0;
      return t?Math.round(c/t*100):0;
    }),backgroundColor:'#3da8ff',datalabels:DL,
  }],{scales:{y:{beginAtZero:true,max:100,ticks:{callback:v=>v+'%'}}}});
})();

// ---- Weekly/Monthly ----
multi('cWeekly',DASH.weekly,'bar');
multi('cMonthly',DASH.monthly,'bar');

// ---- Backlog ----
hbar('cBacklog',
  DASH.top_sub_remain.map(r=>r.label.split(' - ')[0]),
  DASH.top_sub_remain.map(r=>r.remain),
  '#ff4d8d','Remaining');

// ---- RFI ----
pie('cRfiStatus',Object.keys(DASH.rfi_status),Object.values(DASH.rfi_status));
pie('cRfiDisc',Object.keys(DASH.rfi_disc),Object.values(DASH.rfi_disc));
(function(){
  const dl=DASH.rfi_daily, ml=DASH.rfi_monthly;
  bar('cRfiDaily',dl.map(r=>r.label),[{label:'RFIs',data:dl.map(r=>r.count),backgroundColor:'#3da8ff',borderRadius:5,datalabels:DL}]);
  bar('cRfiMonthly',ml.map(r=>r.label),[{label:'RFIs',data:ml.map(r=>r.count),backgroundColor:'#9b59f6',borderRadius:5,datalabels:DL}]);
})();

// ---- Punch ----
pie('cPunchStatus',Object.keys(DASH.punch_status),Object.values(DASH.punch_status));
pie('cPunchDisc',Object.keys(DASH.punch_disc),Object.values(DASH.punch_disc));
hbar('cPunchSub',DASH.top_punch_sub.map(r=>r.label.split(' - ')[0]),
  DASH.top_punch_sub.map(r=>r.count),'#ff4d8d','Punch Items');

// ====================================================================
//  SEARCH
// ====================================================================
const sColors={
  'Closed':'#00e676','To be completed':'#ffb627','Submitted':'#3da8ff',
  'Accepted':'#00e676','Accepted with Punch':'#ffb627','Open':'#ff4d8d',
  'Hold by EACOP':'#3da8ff','Rejected':'#ff4d8d','No RFI Yet':'#8a93ad',
};
function badge(t,color){
  return `<span class="badge" style="background:${color||'#8a93ad'}22;color:${color||'#8a93ad'};border:1px solid ${color||'#8a93ad'}44">${t}</span>`;
}

let searchMode='all';
function setMode(m,btn){
  searchMode=m;
  document.querySelectorAll('.stab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  doSearch();
}

function getTasksByIds(tids){
  return tids.map(tid=>IDX[tid]).filter(Boolean);
}

function doSearch(){
  const q=document.getElementById('searchInput').value.trim();
  const box=document.getElementById('searchResults');
  if(q.length<3){
    box.innerHTML='<div style="color:var(--muted);padding:20px;text-align:center;">Type at least 3 characters...</div>';
    return;
  }
  const ql=q.toLowerCase();
  let tasks=[];
  const seen=new Set();

  function addTasks(tids){
    for(const tid of (tids||[])){
      if(!seen.has(tid) && IDX[tid]){ seen.add(tid); tasks.push(IDX[tid]); }
    }
  }

  const modes=['task','asset','rfi','sub','ms'];
  const checkModes = searchMode==='all' ? modes : [searchMode];

  if(checkModes.includes('task')){
    // exact
    if(IDX[q.toUpperCase()]) addTasks([q.toUpperCase()]);
    // partial
    for(const tid in IDX) if(tid.toLowerCase().includes(ql)) addTasks([tid]);
  }
  if(checkModes.includes('asset')){
    for(const at in AMAP) if(at.toLowerCase().includes(ql)) addTasks(AMAP[at]);
  }
  if(checkModes.includes('rfi')){
    for(const rn in RMAP) if(rn.toLowerCase().includes(ql)) addTasks(RMAP[rn]);
  }
  if(checkModes.includes('sub')){
    for(const s in SMAP) if(s.toLowerCase().includes(ql)) addTasks(SMAP[s]);
  }
  if(checkModes.includes('ms')){
    for(const p in PMAP) if(p.toLowerCase().includes(ql)) addTasks(PMAP[p]);
  }

  tasks=tasks.slice(0,50);

  if(!tasks.length){
    box.innerHTML='<div style="color:var(--accent);padding:20px;text-align:center;">⚠️ No results found.</div>';
    return;
  }

  box.innerHTML=`<div style="color:var(--muted);font-size:12px;margin-bottom:10px;">
    Showing ${tasks.length} result(s)</div>`
  +tasks.map(r=>{
    const discColor={'E':'#ffb627','I':'#00e0c6','T':'#ff4d8d'}[(r.disc||r.discipline||'')[0]]||'#8a93ad';
    return `<div class="result-card">
      <div class="result-header">
        <div>
          <div class="result-title">📍 ${r.at||r.asset_tag||r.tid||r.task_id}</div>
          <div class="result-meta">
            ${badge(r.disc||r.discipline||'Unknown',discColor)}
            ${badge(r.st||r.state,sColors[r.st||r.state]||'#8a93ad')}
            ${(r.ms||r.milestone)?badge('🎯 '+(r.ms||r.milestone),'#ffb627'):''}
            <span style="color:var(--muted);font-size:11px;margin-left:6px;">${r.sub||r.subsystem||''}</span>
          </div>
        </div>
        <div style="text-align:right;font-size:11px;color:var(--muted);">
          Task ID: <b style="color:var(--gold)">${r.tid||r.task_id}</b><br>
          ${(r.cd||r.closing_date)?'Closed: '+(r.cd||r.closing_date):(r.eit||r.eit_date)?'EIT: '+(r.eit||r.eit_date):''}
        </div>
      </div>

      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:10px;">

        <!-- ITR -->
        <div>
          <div style="font-weight:700;color:#fff;margin-bottom:4px;font-size:12px;">
            🧪 ITR Details</div>
          <table class="rtable">
            <tr><th>Field</th><th>Value</th></tr>
            <tr><td>Task Type</td><td>${r.ty||r.task_type||'-'}</td></tr>
            <tr><td>Category</td><td>${r.cat||r.category||'-'}</td></tr>
            <tr><td>Responsible</td><td>${r.res||r.responsible||'-'}</td></tr>
            <tr><td>EIT Date</td><td style="color:var(--blue)">${r.eit||r.eit_date||'-'}</td></tr>
            <tr><td>RFC Date</td><td style="color:var(--green)">${r.rfc||r.rfc_date||'-'}</td></tr>
            <tr><td>Month</td><td>${r.mon||r.month||'-'}</td></tr>
          </table>
        </div>

        <!-- RFI -->
        <div>
          <div style="font-weight:700;color:#fff;margin-bottom:4px;font-size:12px;">
            📝 RFI (${r.rfi.length})</div>
          ${r.rfi.length?`<table class="rtable">
            <tr><th>RFI No</th><th>Status</th><th>Type</th><th>Date</th><th>Inspector</th></tr>
            ${r.rfi.map(rf=>`<tr>
              <td style="color:var(--blue)">${rf.raw||rf.rfi_no_raw||rf.no||'-'}</td>
              <td>${badge(rf.st||rf.status,sColors[rf.st||rf.status]||'#8a93ad')}</td>
              <td>${rf.ty||rf.type||'-'}</td>
              <td>${rf.dt||rf.insp_date||'-'}</td>
              <td>${rf.ins||rf.inspector||'-'}</td>
            </tr>`).join('')}
          </table>`:`<div class="no-data">No RFI submitted yet</div>`}
        </div>

        <!-- Punch -->
        <div>
          <div style="font-weight:700;color:#fff;margin-bottom:4px;font-size:12px;">
            📌 Punch List (${r.punch.length})</div>
          ${r.punch.length?`<table class="rtable">
            <tr><th>PL ID</th><th>Category</th><th>Disc</th><th>Status</th><th>Date</th></tr>
            ${r.punch.map(p=>`<tr>
              <td style="color:var(--accent)">${p.id}</td>
              <td>${p.cat||'-'}</td>
              <td>${p.d||'-'}</td>
              <td>${badge(p.st,sColors[p.st]||'#ff4d8d')}</td>
              <td>${p.dt||'-'}</td>
            </tr>`).join('')}
          </table>`:`<div class="no-data">No punch items ✅</div>`}
        </div>

      </div>
    </div>`;
  }).join('');
}

document.getElementById('searchInput').addEventListener('input', doSearch);
doSearch();
</script>
</body>
</html>
"""

def build_html(master, dash, output_path):
    import pandas as pd
    now = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')
    stats = master['stats']

    html = HTML
    html = html.replace('__MASTER_JSON__', json.dumps(master, ensure_ascii=False, default=str))
    html = html.replace('__DASH_JSON__',   json.dumps(dash,   ensure_ascii=False, default=str))
    html = html.replace('__NOW__',     now)
    html = html.replace('__TOTAL__',   str(stats['total']))
    html = html.replace('__CLOSED__',  str(stats['closed']))
    html = html.replace('__OPEN__',    str(stats['total']-stats['closed']))
    html = html.replace('__PCT__',     str(round(stats['closed']/stats['total']*100,1) if stats['total'] else 0))
    html = html.replace('__RFIS__',    str(stats.get('rfis', stats.get('with_rfi',0))))
    html = html.replace('__PUNCH__',   str(dash.get('total_punch', sum(dash['punch_status'].values()))))
    html = html.replace('__ASSETS__',  str(stats.get('assets', stats.get('asset_tags',0))))
    html = html.replace('__SUBS__',    str(stats.get('subsystems',0)))

    with open(output_path,'w',encoding='utf-8') as f:
        f.write(html)
    print(f"Saved: {output_path} ({len(html)/1e6:.1f} MB)")

def main():
    print("="*60)
    print("  PS5 Dashboard v2.0 — Master Index Build")
    print("="*60)

    ov    = find_file(['ovtasks'])
    dpr   = find_file(['completions','dpr'])
    rfi   = find_file(['inspection','register'])
    punch = find_file(['punch','list'])

    if not ov:
        print("ERROR: ovTasks file not found!"); input("Enter..."); return

    print(f"ovTasks   : {ov}")
    print(f"DPR       : {dpr or 'NOT FOUND'}")
    print(f"Inspection: {rfi or 'NOT FOUND'}")
    print(f"Punch List: {punch or 'NOT FOUND'}")

    master = build_master_index(ov, dpr, rfi, punch)
    dash   = build_dashboard_data(master['index'], punch, rfi)

    build_html(master, dash, OUTPUT_HTML)
    shutil.copy(OUTPUT_HTML, OUTPUT_HTML2)
    print(f"Also saved: {OUTPUT_HTML2}")

    print("\n"+"="*60)
    print("  DONE! Dashboard ready.")
    print(f"  GitHub: upload index.html")
    print("="*60)

    webbrowser.open('file://'+OUTPUT_HTML)
    input("\nPress Enter to exit...")

if __name__ == '__main__':
    main()
"""
Master Index Builder
يبني فهرس موحد من 4 ملفات:
1. ovTasks (Exported from SC)     -> Task ID, Asset Tag, State, Closing Date, Subsystem, Priority
2. DPR (DETAILED ITR LIST)        -> EIT Date, RFC Date, Month, Milestone
3. Inspection Register            -> RFI No, Status, Inspection Date, Type, Internal Status
4. Punch List (MP Register)       -> PL ID, Category, Status, Raised Date

المفتاح الرئيسي: Task ID (لو موجود) ثم Asset Tag
الربط بين Punch و باقي الملفات: عن طريق RFI No (مع normalization)
"""

import re
import pandas as pd
import json

DISC_NORM = {
    'I - Instrumentation': 'I - Instrumentation',
    'I - Instrument':      'I - Instrumentation',
    'Instrumentation':     'I - Instrumentation',
    'Instrument':          'I - Instrumentation',
    'E - Electrical':      'E - Electrical',
    'Electrical':          'E - Electrical',
    'T - Telecom':         'T - Telecom',
    'Telecom':             'T - Telecom',
}

def norm_disc(v):
    if not isinstance(v, str): return str(v) if v else ''
    v = v.strip()
    return DISC_NORM.get(v, v)

def norm_rfi(v):
    """توحيد RFI No بإزالة الفواصل والمسافات"""
    if not v or str(v).strip().lower() in ('nan','none',''): return ''
    return re.sub(r'[\s/]+', '-', str(v).strip().upper())

def norm_str(v):
    return str(v).strip() if pd.notna(v) else ''

def norm_date(v):
    try:
        return pd.to_datetime(v).strftime('%Y-%m-%d')
    except:
        return ''


def build_master_index(ov_path, dpr_path, rfi_path, punch_path):
    """
    يبني فهرس موحد:
    index[task_id] = {
        'task_id':    ...,
        'asset_tag':  ...,
        'discipline': ...,  # normalized
        'state':      ...,
        'task_type':  ...,
        'description':..,
        'subsystem':  ...,
        'milestone':  ...,  # Subsystem Priority
        'closing_date':...,
        'closed_by':  ...,
        'eit_date':   ...,  # من DPR
        'rfc_date':   ...,  # من DPR
        'month':      ...,  # من DPR
        'rfi': [{           # من Inspection Register
            'rfi_no':...,
            'status':...,
            'insp_date':...,
            'type':...,
            'internal_status':...,
            'precomm_itr_status':...,
            'rfc_date':...,
            'milestone':...,
        }],
        'punch': [{         # من Punch List (عن طريق RFI No)
            'plid':...,
            'category':...,
            'discipline':...,
            'status':...,
            'raised_date':...,
            'responsible':...,
            'rfi_no':...,
        }],
    }
    """
    print("=" * 60)
    print("  Building Master Index...")
    print("=" * 60)

    # ================================================================
    # 1. ovTasks - Base
    # ================================================================
    print("\n[1] Loading ovTasks...")
    ov = pd.read_excel(ov_path, sheet_name='Exported from SC')
    ov = ov.where(pd.notna(ov), other=None)

    index = {}  # task_id -> record

    for _, r in ov.iterrows():
        tid  = norm_str(r['Task ID'])
        atag = norm_str(r['Asset - Tag'])
        if not tid: continue

        index[tid] = {
            'task_id':     tid,
            'asset_tag':   atag,
            'discipline':  norm_disc(norm_str(r['Discipline (Summary)'])),
            'state':       norm_str(r['Task State']),
            'task_type':   norm_str(r['Task Type (Name)']),
            'description': norm_str(r['Description']),
            'subsystem':   norm_str(r['Systemization - Subsystem (Summary)']),
            'milestone':   norm_str(r['Subsystem Priority']),
            'closing_date':norm_date(r['Closing Date']),
            'closed_by':   norm_str(r['Closed By']) if 'Closed By' in r else '',
            'category':    norm_str(r['Category (Summary)']),
            'responsible': norm_str(r['Responsible Company (Summary)']),
            # from DPR (to be filled)
            'eit_date':    '',
            'rfc_date':    '',
            'month':       '',
            # from Inspection
            'rfi':   [],
            # from Punch
            'punch': [],
        }

    # secondary map: asset_tag -> [task_ids] (for linking by asset tag)
    asset_to_tids = {}
    for tid, rec in index.items():
        atag = rec['asset_tag']
        if atag:
            asset_to_tids.setdefault(atag, [])
            asset_to_tids[atag].append(tid)

    print(f"  -> {len(index)} tasks loaded")

    # ================================================================
    # 2. DPR DETAILED ITR LIST - add EIT/RFC dates
    # ================================================================
    if dpr_path:
        print("\n[2] Loading DPR Detailed ITR List...")
        dpr = pd.read_excel(dpr_path, sheet_name='DETAILED ITR LIST')
        matched = 0
        for _, r in dpr.iterrows():
            tid = norm_str(r['Task ID'])
            if tid in index:
                index[tid]['eit_date'] = norm_date(r.get('EIT'))
                index[tid]['rfc_date'] = norm_date(r.get('RFC'))
                index[tid]['month']    = norm_str(r.get('MONTH'))
                # override milestone from DPR if better
                if r.get('Priority.2') and not index[tid]['milestone']:
                    index[tid]['milestone'] = norm_str(r.get('Priority.2'))
                matched += 1
        print(f"  -> {matched} tasks enriched with EIT/RFC dates")

    # ================================================================
    # 3. Inspection Register - add RFI data
    # ================================================================
    if rfi_path:
        print("\n[3] Loading Inspection Register...")
        rdf = pd.read_excel(rfi_path, sheet_name='PS-5 EIT INSPECTION REGISTER', header=4)

        def norm_rfi_status(v):
            if not isinstance(v, str): return 'No RFI Yet'
            v = v.strip().lower()
            if 'punch' in v: return 'Accepted with Punch'
            if 'accept' in v: return 'Accepted'
            if 'hold' in v: return 'Hold by EACOP'
            if 'reject' in v: return 'Rejected'
            if 'open' in v: return 'Open'
            return v.title()

        matched_by_tid = 0
        matched_by_tag = 0
        unmatched = 0

        # Build RFI No -> task_ids map (from inspection register itself)
        rfi_to_tids = {}

        for _, r in rdf.iterrows():
            tid  = norm_str(r.get('Task ID', ''))
            atag = norm_str(r.get('Asset - Tag', ''))
            rfi_no = norm_str(r.get('QC RFI No', ''))

            rfi_rec = {
                'rfi_no':            norm_rfi(rfi_no),
                'rfi_no_raw':        rfi_no,
                'status':            norm_rfi_status(r.get('Status Of RFI ')),
                'insp_date':         norm_date(r.get('Inspection Date')),
                'internal_date':     norm_date(r.get('Internal Inspection date')),
                'internal_status':   norm_str(r.get('Internal Inspection status (Ready for RFI or not) ')),
                'precomm_itr_status':norm_str(r.get('Status of Precomm ITR ')),
                'type':              norm_str(r.get('INSPECTION TYPE (PULLING/GLANDING & TERMINATION/CABLE TESTING/INSTALLATION)')),
                'milestone':         norm_str(r.get('MILESTONE')),
                'rfc_date':          norm_date(r.get('RFC DATE ')),
                'inspector':         norm_str(r.get('QC INSPECTOR ')),
                'remarks':           norm_str(r.get('Remarks')),
            }

            # Link by Task ID first
            if tid and tid in index:
                index[tid]['rfi'].append(rfi_rec)
                if rfi_no:
                    rfi_to_tids.setdefault(norm_rfi(rfi_no), [])
                    if tid not in rfi_to_tids[norm_rfi(rfi_no)]:
                        rfi_to_tids[norm_rfi(rfi_no)].append(tid)
                matched_by_tid += 1
                continue

            # Link by Asset Tag
            if atag and atag in asset_to_tids:
                for t in asset_to_tids[atag]:
                    index[t]['rfi'].append(rfi_rec)
                    if rfi_no:
                        rfi_to_tids.setdefault(norm_rfi(rfi_no), [])
                        if t not in rfi_to_tids[norm_rfi(rfi_no)]:
                            rfi_to_tids[norm_rfi(rfi_no)].append(t)
                matched_by_tag += 1
                continue

            unmatched += 1

        print(f"  -> Matched by Task ID: {matched_by_tid}")
        print(f"  -> Matched by Asset Tag: {matched_by_tag}")
        print(f"  -> Unmatched: {unmatched}")
    else:
        rfi_to_tids = {}

    # ================================================================
    # 4. Punch List - add Punch data (link via normalized RFI No)
    # ================================================================
    if punch_path:
        print("\n[4] Loading Punch List...")
        raw = pd.read_excel(punch_path, sheet_name='MP Register', header=None, skiprows=6)
        raw.columns = range(len(raw.columns))
        punch = raw[[1,3,5,8,9,10,13,14,20]].rename(columns={
            1:'plid', 3:'rfi_no', 5:'subsystem', 8:'desc',
            9:'cat', 10:'discipline', 13:'raised_date', 14:'responsible', 20:'status'
        })
        punch = punch.dropna(subset=['plid'])

        matched_punch = 0
        unmatched_punch = []

        for _, r in punch.iterrows():
            rfi_raw = norm_str(r.get('rfi_no',''))
            rfi_norm = norm_rfi(rfi_raw)

            punch_rec = {
                'plid':       norm_str(r['plid']),
                'rfi_no':     rfi_norm,
                'rfi_no_raw': rfi_raw,
                'subsystem':  norm_str(r.get('subsystem','')),
                'discipline': norm_disc(norm_str(r.get('discipline',''))),
                'category':   norm_str(r.get('cat','')),
                'status':     norm_str(r.get('status','')) or 'Open',
                'raised_date':norm_date(r.get('raised_date')),
                'responsible':norm_str(r.get('responsible','')),
            }

            # Link via RFI No
            tids = rfi_to_tids.get(rfi_norm, [])
            if tids:
                for tid in tids:
                    # avoid duplicates
                    existing_plids = [p['plid'] for p in index[tid]['punch']]
                    if punch_rec['plid'] not in existing_plids:
                        index[tid]['punch'].append(punch_rec)
                matched_punch += 1
            else:
                unmatched_punch.append(punch_rec)

        print(f"  -> Matched punch items: {matched_punch}")
        print(f"  -> Unmatched punch items: {len(unmatched_punch)}")

    # ================================================================
    # 5. Build search maps
    # ================================================================
    print("\n[5] Building search maps...")

    # Asset Tag -> [task_ids]
    asset_map = {}
    for tid, rec in index.items():
        atag = rec['asset_tag']
        if atag:
            asset_map.setdefault(atag, [])
            if tid not in asset_map[atag]:
                asset_map[atag].append(tid)

    # RFI No -> [task_ids]
    rfi_map = rfi_to_tids

    # Subsystem -> [task_ids] (deduped by asset tag)
    sub_map = {}
    for tid, rec in index.items():
        sub = rec['subsystem']
        if sub:
            sub_map.setdefault(sub, [])
            if tid not in sub_map[sub]:
                sub_map[sub].append(tid)

    # Milestone/Priority -> [task_ids]
    pri_map = {}
    for tid, rec in index.items():
        ms = rec['milestone']
        if ms:
            pri_map.setdefault(ms, [])
            if tid not in pri_map[ms]:
                pri_map[ms].append(tid)

    # Stats
    with_rfi   = sum(1 for r in index.values() if r['rfi'])
    with_punch = sum(1 for r in index.values() if r['punch'])
    closed     = sum(1 for r in index.values() if r['state'] == 'Closed')

    print(f"\n{'='*60}")
    print(f"  MASTER INDEX COMPLETE")
    print(f"  Total tasks        : {len(index)}")
    print(f"  Closed             : {closed}")
    print(f"  With RFI data      : {with_rfi}")
    print(f"  With Punch data    : {with_punch}")
    print(f"  Asset Tag map      : {len(asset_map)} unique assets")
    print(f"  RFI No map         : {len(rfi_map)} unique RFIs")
    print(f"  Subsystem map      : {len(sub_map)} unique subsystems")
    print(f"  Milestone map      : {len(pri_map)} unique milestones")
    print(f"{'='*60}")

    return {
        'index':    index,
        'asset_map': asset_map,
        'rfi_map':  rfi_map,
        'sub_map':  sub_map,
        'pri_map':  pri_map,
        'stats': {
            'total': len(index),
            'closed': closed,
            'with_rfi': with_rfi,
            'with_punch': with_punch,
        }
    }


