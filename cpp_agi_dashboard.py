"""
====================================================================
 CPP AGI / PS5 - EACOP Project Dashboard
 -----------------------------------------------------------------
 Source 1: ovTasks_TestsPlanned*.xlsx  (sheet 'Exported from SC')
            -> ITR Closures (Hourly / Daily / Weekly / Monthly / Total)
 Source 2: *PUNCH_LIST_REGISTER*.xlsx  (sheet 'MP Register')
            -> Punch List items (Daily / Weekly / Monthly / Total,
               by Status, Category, Discipline)

 Output: HTML Dashboard (Oil & Gas purple theme) with a SIDEBAR
         checklist to show/hide each section.
====================================================================
 المطلوب:
    pip install pandas openpyxl --break-system-packages
====================================================================
"""

import os
import sys
import json
import datetime
import urllib.request
import webbrowser

try:
    import pandas as pd
except ImportError:
    os.system("pip install pandas openpyxl --break-system-packages")
    import pandas as pd


# ====================================================================
#  إعدادات
# ====================================================================
DOWNLOADS = r"C:\Users\mylap\OneDrive\Desktop\dashboard"
OUTPUT_HTML = os.path.join(DOWNLOADS, "index.html")
OUTPUT_HTML2 = os.path.join(DOWNLOADS, "PS5_Project_Dashboard.html")

RESPONSIBLE_COMPANY = "CPP AGI"
DISCIPLINES = {
    'E - Electrical':       'Electrical (E)',
    'I - Instrumentation':  'Instrumentation (I)',
    'I - Instrument':       'Instrumentation (I)',
    'Instrumentation':      'Instrumentation (I)',
    'Instrument':           'Instrumentation (I)',
    'T - Telecom':          'Telecom (T)',
    'Telecom':              'Telecom (T)',
}

def norm_discipline(v):
    """توحيد أسماء التخصصات"""
    if not isinstance(v, str): return str(v) if v else ''
    return DISCIPLINES.get(v.strip(), v.strip())



# ====================================================================
#  1) إيجاد الملفات تلقائياً
# ====================================================================
def find_file(prefix_keywords):
    for f in os.listdir(DOWNLOADS):
        if f.startswith('~$'):
            continue
        low = f.lower()
        if low.endswith('.xlsx') and all(k.lower() in low for k in prefix_keywords):
            return os.path.join(DOWNLOADS, f)
    return None


# ====================================================================
#  2) ITR Closures من ovTasks
def safe_read_excel(path, **kwargs):
    """Read Excel with fallback copy if file is locked."""
    import shutil, tempfile
    try:
        return pd.read_excel(path, **kwargs)
    except PermissionError:
        tmp = os.path.join(tempfile.gettempdir(), 'ov_tmp.xlsx')
        try:
            shutil.copy2(path, tmp)
        except PermissionError:
            pass
        return pd.read_excel(tmp, **kwargs)


# ====================================================================
def build_itr_data(excel_path, today_override=None):
    df = safe_read_excel(excel_path, sheet_name='Exported from SC')

    total_project_tasks = len(df)

    # ---- Closed Tasks ----
    closed = df[df['Task State'] == 'Closed'].copy()
    closed['Closing Date'] = pd.to_datetime(closed['Closing Date'])
    closed = closed.dropna(subset=['Closing Date'])

    # ---- Submitted Tasks ----
    submitted = df[df['Task State'] == 'Submitted'].copy()
    submitted['Closing Date'] = pd.to_datetime(submitted['Closing Date'])
    # Submitted tasks usually have no Closing Date; use today for them
    submitted['Closing Date'] = submitted['Closing Date'].fillna(pd.Timestamp.now())

    # ---- نوحد الـ Discipline لـ E / I / T / Other ----
    disc_short = {
        'E - Electrical':      'E',
        'E - Elect':           'E',
        'Electrical':          'E',
        'I - Instrumentation': 'I',
        'I - Instrument':      'I',
        'Instrumentation':     'I',
        'Instrument':          'I',
        'T - Telecom':         'T',
        'Telecom':             'T',
    }
    closed['disc'] = closed['Discipline (Summary)'].astype(str).str.strip().map(disc_short).fillna('Other')
    submitted['disc'] = submitted['Discipline (Summary)'].astype(str).str.strip().map(disc_short).fillna('Other')

    now = pd.Timestamp(today_override) if today_override else closed['Closing Date'].max()
    print(f"  Today's date used: {now.strftime('%Y-%m-%d')}" + (" (override)" if today_override else " (from data)"))
    total_closed_project = len(closed)

    def pivot_disc(data, col):
        if data.empty:
            return []
        p = data.groupby([col, 'disc']).size().unstack(fill_value=0)
        for d in ['E', 'I', 'T']:
            if d not in p.columns:
                p[d] = 0
        p['Total'] = p[['E', 'I', 'T']].sum(axis=1)
        p = p[['E', 'I', 'T', 'Total']]
        return p.reset_index().rename(columns={col: 'label'}).to_dict('records')

    # ---- Hourly (Today) for Closed ----
    today_closed = closed[closed['Closing Date'].dt.date == now.date()].copy()
    today_closed['hour'] = today_closed['Closing Date'].dt.hour.apply(lambda h: f"{h:02d}:00")
    hourly = pivot_disc(today_closed, 'hour')
    hourly_total = int(today_closed.shape[0])
    # E&I&T CPP AGI today closed
    hourly_closed_eit = int(today_closed[today_closed['disc'].isin(['E','I','T']) & (today_closed['Responsible Company (Summary)'] == RESPONSIBLE_COMPANY)].shape[0])

    # ---- Hourly (Today) for Submitted ----
    today_submitted = submitted[submitted['Closing Date'].dt.date == now.date()].copy()
    hourly_submitted = int(today_submitted.shape[0])

    # ---- Daily (Last 30 Days) for Closed (CPP AGI E/I/T only) ----
    eit_mask = (closed['disc'].isin(['E','I','T'])) & (closed['Responsible Company (Summary)'] == RESPONSIBLE_COMPANY)
    closed['date'] = closed['Closing Date'].dt.strftime('%Y-%m-%d')
    daily = pivot_disc(closed[eit_mask], 'date')
    daily = sorted(daily, key=lambda r: r['label'])[-30:]
    daily_total = int(eit_mask.sum())

    # ---- Fill in all dates in last 30 days even with 0 ----
    if daily:
        all_dates = pd.date_range(end=now.date(), periods=30).strftime('%Y-%m-%d').tolist()
        found = {r['label']: r for r in daily}
        daily = [found.get(d, {'label': d, 'E': 0, 'I': 0, 'T': 0, 'Total': 0}) for d in all_dates]

    # ---- Daily (Last 30 Days) for Submitted (CPP AGI E/I/T only) ----
    sub_eit_mask = (submitted['disc'].isin(['E','I','T'])) & (submitted['Responsible Company (Summary)'] == RESPONSIBLE_COMPANY)
    submitted['date'] = submitted['Closing Date'].dt.strftime('%Y-%m-%d')
    daily_submitted = pivot_disc(submitted[sub_eit_mask], 'date')
    daily_submitted = sorted(daily_submitted, key=lambda r: r['label'])[-30:]
    if daily_submitted:
        daily_submitted = [found.get(d, {'label': d, 'E': 0, 'I': 0, 'T': 0, 'Total': 0}) for d in all_dates]
    else:
        daily_submitted = [{'label': d, 'E': 0, 'I': 0, 'T': 0, 'Total': 0} for d in all_dates]

    # ---- Weekly for Closed (CPP AGI E/I/T only) ----
    closed['week'] = closed['Closing Date'].dt.to_period('W').apply(
        lambda r: r.start_time.strftime('%Y-%m-%d'))
    weekly = pivot_disc(closed[eit_mask], 'week')
    weekly = sorted(weekly, key=lambda r: r['label'])[-16:]

    week_start = (now - pd.Timedelta(days=now.weekday())).normalize()
    weekly_total = int((closed[eit_mask]['Closing Date'] >= week_start).sum())

    # ---- Monthly for Closed (CPP AGI E/I/T only) ----
    closed['month'] = closed['Closing Date'].dt.strftime('%Y-%m')
    monthly = pivot_disc(closed[eit_mask], 'month')
    monthly = sorted(monthly, key=lambda r: r['label'])

    monthly_total = int(((closed[eit_mask]['Closing Date'].dt.year == now.year) &
                          (closed[eit_mask]['Closing Date'].dt.month == now.month)).sum())


    # ---- E&I&T لـ CPP AGI (Summary cards) — 3 كروت فقط: E / I / T ----
    eit_summary = []
    eit = df[(df['Responsible Company (Summary)'] == RESPONSIBLE_COMPANY)].copy()
    eit['disc_norm'] = eit['Discipline (Summary)'].astype(str).str.strip().map(disc_short).fillna('Other')

    GROUPS = [('E', 'Electrical (E)'), ('I', 'Instrumentation (I)'), ('T', 'Telecom (T)')]
    for code, label in GROUPS:
        sub = eit[eit['disc_norm'] == code]
        tot = len(sub)
        cls = int((sub['Task State'] == 'Closed').sum())
        eit_summary.append({
            'label': label, 'total': int(tot), 'closed': cls,
            'pct': round(cls / tot * 100, 1) if tot else 0
        })

    # ---- Progress by Milestone لـ CPP AGI (عمود Subsystem Priority = Milestone) ----
    milestone_summary = []
    eit['ms_raw'] = eit['Subsystem Priority'].apply(
        lambda v: v.strip() if isinstance(v, str) and v.strip() else None)

    STATUS_GROUPS = {
        'Closed': 'Closed',
        'Submitted': 'Submitted',
        'To be completed': 'To be completed',
    }
    STATUS_ORDER = ['Closed', 'Submitted', 'To be completed', 'Other']

    for ms_raw in sorted(eit['ms_raw'].dropna().unique()):
        sub = eit[eit['ms_raw'] == ms_raw]
        tot = len(sub)
        cls = int((sub['Task State'] == 'Closed').sum())
        status_grp = sub['Task State'].astype(str).str.strip().map(STATUS_GROUPS).fillna('Other')
        status_counts = status_grp.value_counts().to_dict()
        status = {k: int(status_counts.get(k, 0)) for k in STATUS_ORDER}
        milestone_summary.append({
            'label': ms_raw.replace('PS5 - ', '').replace('PS5-', ''),
            'raw': ms_raw,
            'total': int(tot), 'closed': cls, 'remaining': int(tot - cls),
            'pct': round(cls / tot * 100, 1) if tot else 0,
            'status': status,
        })

    total_sub = int((eit['Task State'] == 'Submitted').sum())
    today_str = now.strftime('%Y-%m-%d')

    # ---- Subsystem summary: E/I/T per subsystem ----
    subsystem_summary = []
    eit['sub_raw'] = eit['Systemization - Subsystem (Summary)'].apply(
        lambda v: v.strip() if isinstance(v, str) and v.strip() else 'Unknown')
    for sub_name in sorted(eit['sub_raw'].unique()):
        sub_df = eit[eit['sub_raw'] == sub_name]
        for code, label in GROUPS:
            s = sub_df[sub_df['disc_norm'] == code]
            tot = len(s)
            if tot == 0:
                continue
            cls = int((s['Task State'] == 'Closed').sum())
            opn = tot - cls
            subsystem_summary.append({
                'subsystem': sub_name, 'disc': code, 'discipline': label,
                'total': tot, 'closed': cls, 'open': opn,
                'pct': round(cls / tot * 100, 1) if tot else 0,
            })

    # ---- Today's closures by milestone ----
    today_all = df[(df['Task State'] == 'Closed') &
                      (pd.to_datetime(df['Closing Date'], errors='coerce').dt.date == now.date())].copy()
    today_all['disc_norm'] = today_all['Discipline (Summary)'].astype(str).str.strip().map(disc_short).fillna('Other')
    today_all['ms_raw'] = today_all['Subsystem Priority'].apply(
        lambda v: v.strip() if isinstance(v, str) and v.strip() else 'Unknown')
    today_closed_eit = today_all[today_all['disc_norm'].isin(['E', 'I', 'T']) &
                                     (today_all['Responsible Company (Summary)'] == RESPONSIBLE_COMPANY)]
    today_milestone = []
    for ms in sorted(today_closed_eit['ms_raw'].unique()):
        ms_df = today_closed_eit[today_closed_eit['ms_raw'] == ms]
        ms_label = ms.replace('PS5 - ', '').replace('PS5-', '')
        mc = int(len(ms_df))
        for code, label in GROUPS:
            s = ms_df[ms_df['disc_norm'] == code]
            today_milestone.append({
                'milestone': ms_label, 'disc': code, 'discipline': label,
                'count': int(len(s)),
            })
    today_milestone_total = int(len(today_closed_eit))

    # ---- Asset-based metrics (E&I&T CPP AGI) ----
    eit_assets = eit.groupby('Asset - Tag').agg(
        has_closed=('Task State', lambda s: (s == 'Closed').any()),
        has_submitted=('Task State', lambda s: (s == 'Submitted').any()),
    )
    total_assets_eit = len(eit_assets)
    submitted_assets = int(eit_assets['has_submitted'].sum())
    closed_assets = int(eit_assets['has_closed'].sum())

    # ---- Cache diff for daily tracking ----
    CACHE_FILE = os.path.join(os.path.dirname(__file__) or '.', 'submitted_cache.json')
    bl_total = total_sub
    bl_sub_assets = submitted_assets
    bl_cls_assets = closed_assets
    prev_date = ''
    try:
        with open(CACHE_FILE) as cf:
            c = json.load(cf)
            bl_total = c.get('bl_total', total_sub)
            bl_sub_assets = c.get('bl_sub_assets', submitted_assets)
            bl_cls_assets = c.get('bl_cls_assets', closed_assets)
            prev_date = c.get('date', '')
    except:
        pass
    if prev_date and today_str != prev_date:
        # New day: update baseline to yesterday's total
        bl_total = int(json.load(open(CACHE_FILE)).get('total', total_sub))
        bl_sub_assets = int(json.load(open(CACHE_FILE)).get('sub_assets', submitted_assets))
        bl_cls_assets = int(json.load(open(CACHE_FILE)).get('cls_assets', closed_assets))
    hourly_submitted = 0
    today_submitted_assets = 0
    today_closed_assets = max(0, closed_assets - bl_cls_assets)
    with open(CACHE_FILE, 'w') as cf:
        json.dump({'total': total_sub, 'date': today_str,
                   'bl_total': bl_total, 'bl_sub_assets': bl_sub_assets, 'bl_cls_assets': bl_cls_assets,
                   'sub_assets': submitted_assets, 'cls_assets': closed_assets}, cf)


    # E&I&T specific totals
    total_closed_eit = sum(s['closed'] for s in eit_summary)
    total_tasks_eit = sum(s['total'] for s in eit_summary)
    hourly_closed_eit = int(today_closed[today_closed['disc'].isin(['E','I','T']) & (today_closed['Responsible Company (Summary)'] == RESPONSIBLE_COMPANY)].shape[0])

    return {
        'now': now.strftime('%Y-%m-%d %H:%M'),
        'today_label': today_str,
        'hourly': hourly,
        'hourly_total': hourly_total,
        'hourly_closed_eit': hourly_closed_eit,
        'hourly_submitted': hourly_submitted,
        'daily': daily,
        'daily_total': daily_total,
        'daily_submitted': daily_submitted,
        'weekly': weekly,
        'weekly_total': weekly_total,
        'monthly': monthly,
        'monthly_total': monthly_total,
        'total_closed_project': total_closed_project,
        'total_closed_eit': total_closed_eit,
        'total_tasks_eit': total_tasks_eit,
        'total_submitted_project': total_sub,
        'total_project_tasks': total_project_tasks,
        'eit_summary': eit_summary,
        'milestone_summary': milestone_summary,
        'subsystem_summary': subsystem_summary,
        'today_milestone': today_milestone,
        'today_milestone_total': today_milestone_total,
        # Asset-based
        'total_assets_eit': total_assets_eit,
        'submitted_assets': submitted_assets,
        'closed_assets': closed_assets,
        'today_submitted_assets': today_submitted_assets,
        'today_closed_assets': today_closed_assets,
    }


# ====================================================================
#  2b) جدول ITR Description Breakdown — E / I / T (شكل ملف "EIT ITRs")
# ====================================================================
def build_itr_breakdown_table(ov_path):
    """
    يبني جدول بنفس شكل ملف 'EIT ITRs every thursday':
    Type (E-Electrical / I-Instrumentation / T-Telecom) | ITR Description |
    Total Task | Approved by EACOP (Closed) | Balance (Open)
    """
    df = safe_read_excel(ov_path, sheet_name='Exported from SC')

    disc_label = {
        'E - Electrical': 'E-Electrical', 'Electrical': 'E-Electrical',
        'I - Instrumentation': 'I-Instrumentation', 'I - Instrument': 'I-Instrumentation',
        'Instrumentation': 'I-Instrumentation', 'Instrument': 'I-Instrumentation',
        'T - Telecom': 'T-Telecom', 'Telecom': 'T-Telecom',
    }
    df['disc_grp'] = df['Discipline (Summary)'].astype(str).str.strip().map(disc_label)
    eit = df[(df['disc_grp'].notna()) &
             (df['Responsible Company (Summary)'] == RESPONSIBLE_COMPANY)].copy()
    eit['itr_desc'] = eit['Description'].fillna('Unspecified').astype(str).str.strip()
    eit.loc[eit['itr_desc'] == '', 'itr_desc'] = 'Unspecified'

    # ---- تحديد "Closed This Week" زي عمود "From .. to .." في الملف المرجعي ----
    eit['Closing Date'] = pd.to_datetime(eit['Closing Date'], errors='coerce')
    closed_dates = eit.loc[eit['Task State'] == 'Closed', 'Closing Date'].dropna()
    now = closed_dates.max() if not closed_dates.empty else pd.Timestamp.now()
    week_start = (now - pd.Timedelta(days=now.weekday())).normalize()
    week_end = week_start + pd.Timedelta(days=6)
    eit['closed_this_week'] = ((eit['Task State'] == 'Closed') &
                                (eit['Closing Date'] >= week_start) &
                                (eit['Closing Date'] <= week_end + pd.Timedelta(days=1)))

    g = (eit.groupby(['disc_grp', 'itr_desc'])
            .agg(total=('Task ID', 'count'),
                 approved=('Task State', lambda s: int((s == 'Closed').sum())),
                 closed_wk=('closed_this_week', 'sum'))
            .reset_index())
    g['balance'] = g['total'] - g['approved']
    g['pct'] = (g['approved'] / g['total'] * 100).round(1).where(g['total'] > 0, 0)

    order = ['E-Electrical', 'I-Instrumentation', 'T-Telecom']
    g['disc_grp'] = pd.Categorical(g['disc_grp'], categories=order, ordered=True)
    g = g.sort_values(['disc_grp', 'total'], ascending=[True, False])

    rows = []
    for _, r in g.iterrows():
        rows.append({
            'type': str(r['disc_grp']),
            'desc': r['itr_desc'],
            'total': int(r['total']),
            'approved': int(r['approved']),
            'balance': int(r['balance']),
            'pct': float(r['pct']),
            'closed_wk': int(r['closed_wk']),
        })

    totals = {}
    for t in order:
        sub = g[g['disc_grp'] == t]
        tot = int(sub['total'].sum())
        app = int(sub['approved'].sum())
        totals[t] = {
            'total': tot,
            'approved': app,
            'balance': int(sub['balance'].sum()),
            'pct': round(app / tot * 100, 1) if tot else 0,
            'closed_wk': int(sub['closed_wk'].sum()),
        }
    grand_total = int(g['total'].sum())
    grand_approved = int(g['approved'].sum())
    grand = {
        'total': grand_total,
        'approved': grand_approved,
        'balance': int(g['balance'].sum()),
        'pct': round(grand_approved / grand_total * 100, 1) if grand_total else 0,
        'closed_wk': int(g['closed_wk'].sum()),
    }

    # Build EIT_DESC format for individual E/I/T tables
    eit_desc = {}
    for code, label in [('E', 'E-Electrical'), ('I', 'I-Instrumentation'), ('T', 'T-Telecom')]:
        sub = g[g['disc_grp'] == label]
        eit_desc[code] = sub.apply(lambda r: {
            'desc': r['itr_desc'],
            'total': int(r['total']),
            'closed': int(r['approved']),
            'balance': int(r['balance']),
            'pct': float(r['pct']),
        }, axis=1).tolist()

    return {
        'rows': rows, 'totals': totals, 'grand': grand, 'order': order,
        'cutoff': now.strftime('%Y-%m-%d'),
        'week_label': f"From {week_start.strftime('%d %b')} to {week_end.strftime('%d %b')}",
        'eit_desc': eit_desc,
    }


# ====================================================================
#  2b) Cable / CMT OV tasks (for searchable table + Excel export)
# ====================================================================
def build_cable_ov_data(ov_path):
    df = safe_read_excel(ov_path, sheet_name='Exported from SC')
    disc_label = {
        'E - Electrical': 'E-Electrical', 'Electrical': 'E-Electrical',
        'I - Instrumentation': 'I-Instrumentation', 'I - Instrument': 'I-Instrumentation',
        'Instrumentation': 'I-Instrumentation', 'Instrument': 'I-Instrumentation',
        'T - Telecom': 'T-Telecom', 'Telecom': 'T-Telecom',
    }
    df['disc_grp'] = df['Discipline (Summary)'].astype(str).str.strip().map(disc_label)
    eit = df[df['disc_grp'].notna()].copy()
    cable_keywords = ['cable','Cable','CABLE','Cable Tray','Heat Tracing',
                      'Instrument Cable','Fiber Optic','Telecom Cable']
    mask = eit['Test Form - Description'].astype(str).str.contains('|'.join(cable_keywords), na=False, case=False)
    cable = eit[mask].copy()
    records = []
    for _, r in cable.iterrows():
        tag = r['Asset - Tag']
        if pd.isna(tag):
            continue
        closing = r['Closing Date']
        records.append({
            'task_id': str(r['Task ID']) if pd.notna(r['Task ID']) else '',
            'asset_tag': str(tag).strip(),
            'desc': str(r['Test Form - Description']) if pd.notna(r['Test Form - Description']) else '',
            'discipline': r['disc_grp'],
            'state': str(r['Task State']) if pd.notna(r['Task State']) else '',
            'closing_date': pd.to_datetime(closing).strftime('%Y-%m-%d') if pd.notna(closing) else '',
        })
    records.sort(key=lambda r: (r['discipline'], r['task_id']))
    counts = {}
    for r in records:
        counts[r['discipline']] = counts.get(r['discipline'], 0) + 1
    disc_summary = ', '.join(f'{k}: {v}' for k, v in sorted(counts.items()))
    n = len(records)
    print(f"  Cable/CMT OV tasks: {n} records ({disc_summary})")
    return records



# ====================================================================
#  2c) Cable tracker من Pre_Com_Cable_ITR_Tracker.xlsx (PS5 sheet)
# ====================================================================
def build_cable_tracker_data():
    path = os.path.join(os.path.dirname(__file__) or '.', 'Pre_Com_Cable_ITR_Tracker.xlsx')
    if not os.path.exists(path):
        print("  Cable tracker file NOT found")
        return None
    df = pd.read_excel(path, sheet_name='PS5', header=None, skiprows=2)
    df.columns = ['Subsystem','Asset_Tag','Vlookup','Description','Scope','Disc',
                  'Laid_Date','LAYING_RFI','TESTING_RFI','TERM_RFI','CMT','Static_CMT','Remarks']
    df = df[df['Disc'].notna()].copy()
    discs = {'E': 'E-Electrical', 'I': 'I-Instrumentation', 'T': 'T-Telecom'}
    rows = []
    for code, label in discs.items():
        sub = df[df['Disc'] == code]
        tot = len(sub)
        cmt_close = int(sub['CMT'].notna().sum())
        cmt_open = tot - cmt_close
        static_close = int(sub['Static_CMT'].notna().sum())
        static_open = tot - static_close
        laying = int(sub['LAYING_RFI'].notna().sum())
        testing = int(sub['TESTING_RFI'].notna().sum())
        rows.append({
            'disc': label, 'total': tot,
            'cmt_close': cmt_close, 'cmt_open': cmt_open,
            'static_close': static_close, 'static_open': static_open,
            'laying_rfi': laying, 'testing_rfi': testing,
            'pct_cmt': round(cmt_close / tot * 100, 1) if tot else 0,
            'pct_static': round(static_close / tot * 100, 1) if tot else 0,
        })
    print(f"  Cable tracker (PS5): E={rows[0]['total']}, I={rows[1]['total']}, T={rows[2]['total']}")
    # Scope summary — flat per-discipline rows (E/I/T stacked vertically)
    scope_rows = []
    for scope in sorted(df['Scope'].dropna().unique()):
        sub = df[df['Scope'] == scope]
        scope_tot = len(sub)
        scope_cmt = int(sub['CMT'].notna().sum())
        scope_static = int(sub['Static_CMT'].notna().sum())
        for code, label in [('E', 'E — Electrical'), ('I', 'I — Instrumentation'), ('T', 'T — Telecom')]:
            s = sub[sub['Disc'] == code]
            tot = len(s)
            cmt = int(s['CMT'].notna().sum())
            st = int(s['Static_CMT'].notna().sum())
            scope_rows.append({
                'scope': scope, 'disc': code, 'discipline': label,
                'total': tot, 'cmt_close': cmt, 'static_close': st,
                'pct_cmt': round(cmt / tot * 100, 1) if tot else 0,
                'pct_static': round(st / tot * 100, 1) if tot else 0,
                '_scope_total': scope_tot, '_scope_cmt': scope_cmt, '_scope_static': scope_static,
            })
    print(f"  Cable tracker scopes: {len(scope_rows)} scope groups")
    # Raw detail rows for RFI detail table
    detail_rows = []
    for _, r in df.iterrows():
        detail_rows.append({
            'asset': r['Asset_Tag'],
            'subsystem': r['Subsystem'],
            'desc': r['Description'] if pd.notna(r['Description']) else '',
            'disc': r['Disc'],
            'scope': r['Scope'] if pd.notna(r['Scope']) else '',
            'laying_rfi': 1 if pd.notna(r['LAYING_RFI']) else 0,
            'testing_rfi': 1 if pd.notna(r['TESTING_RFI']) else 0,
            'term_rfi': 1 if pd.notna(r['TERM_RFI']) else 0,
        })
    return {'by_disc': rows, 'by_scope': scope_rows, 'detail': detail_rows}

# ====================================================================
#  3) Punch List من PUNCH_LIST_REGISTER
# ====================================================================
def build_punch_data(excel_path):
    if not excel_path:
        return None

    raw = pd.read_excel(excel_path, sheet_name='MP Register', header=None, skiprows=6)
    # الأعمدة حسب الهيدر:
    # 0 MPL No | 1 PL ID | 4 CPP SCOPE | 5 Subsystem | 6 Area | 7 TAG
    # 8 Punch Desc | 9 Cat | 10 Discipline | 13 Raised Date | 20 Status
    cols = {0: 'mpl', 1: 'plid', 3: 'rfi_no', 5: 'subsystem', 8: 'desc',
            9: 'cat', 10: 'discipline', 13: 'raised_date', 20: 'status'}
    df = raw[list(cols.keys())].rename(columns=cols)
    df = df.dropna(subset=['plid'])

    df['raised_date'] = pd.to_datetime(df['raised_date'], errors='coerce')
    df = df.dropna(subset=['raised_date'])
    # توحيد حالة الـ Status (الملف بيجي فيه Open / OPEN / CLOSED ...الخ بأشكال مختلفة)
    df['status'] = df['status'].fillna('Open').astype(str).str.strip().str.title()

    total = len(df)
    now = df['raised_date'].max()

    status_counts = df['status'].value_counts().to_dict()
    cat_counts = df['cat'].value_counts().to_dict()

    # ---- نوحد التخصص لـ E / I / T / Other (قبل أي عد) ----
    disc_short = {
        'Electrical':'E','E - Electrical':'E',
        'Instrumentation':'I','I - Instrumentation':'I',
        'I - Instrument':'I','Instrument':'I',
        'Telecom':'T','T - Telecom':'T',
    }
    df['disc'] = df['discipline'].astype(str).str.strip().map(disc_short).fillna('Other')

    # disc_counts بالأسماء الكاملة — E/I/T فقط، بدون Other
    DISC_FULL_LABEL = {'E': 'Electrical (E)', 'I': 'Instrumentation (I)', 'T': 'Telecom (T)'}
    disc_counts = (df[df['disc'].isin(['E','I','T'])]['disc']
                   .map(DISC_FULL_LABEL).value_counts().to_dict())

    # ---- Open vs Closed لكل Discipline (لشارت "Punch by Discipline (Open vs Closed)") ----
    eit_only = df[df['disc'].isin(['E', 'I', 'T'])].copy()
    eit_only['disc_label'] = eit_only['disc'].map(DISC_FULL_LABEL)
    disc_status = eit_only.groupby(['disc_label', 'status']).size().unstack(fill_value=0)
    for s in ['Open', 'Closed']:
        if s not in disc_status.columns:
            disc_status[s] = 0
    disc_counts_open = disc_status['Open'].to_dict()
    disc_counts_closed = disc_status['Closed'].to_dict()

    def pivot_disc(data, col):
        if data.empty:
            return []
        p = data.groupby([col, 'disc']).size().unstack(fill_value=0)
        for d in ['E', 'I', 'T']:
            if d not in p.columns:
                p[d] = 0
        p['Total'] = p[['E', 'I', 'T']].sum(axis=1)
        p = p[['E', 'I', 'T', 'Total']]
        return p.reset_index().rename(columns={col: 'label'}).to_dict('records')

    # ---- Daily / Weekly / Monthly (تاريخ الرفع) ----
    df['date'] = df['raised_date'].dt.strftime('%Y-%m-%d')
    daily = pivot_disc(df, 'date')
    daily = sorted(daily, key=lambda r: r['label'])[-30:]
    daily_total = int((df['date'] == now.strftime('%Y-%m-%d')).sum())

    df['week'] = df['raised_date'].dt.to_period('W').apply(
        lambda r: r.start_time.strftime('%Y-%m-%d'))
    weekly = pivot_disc(df, 'week')
    weekly = sorted(weekly, key=lambda r: r['label'])[-16:]
    week_start = (now - pd.Timedelta(days=now.weekday())).normalize()
    weekly_total = int((df['raised_date'] >= week_start).sum())

    df['month'] = df['raised_date'].dt.strftime('%Y-%m')
    monthly = pivot_disc(df, 'month')
    monthly = sorted(monthly, key=lambda r: r['label'])
    monthly_total = int(((df['raised_date'].dt.year == now.year) &
                          (df['raised_date'].dt.month == now.month)).sum())

    # ---- Top Subsystems ----
    top_sub = df['subsystem'].value_counts().head(10).reset_index()
    top_sub.columns = ['label', 'count']

    # ---- recent records for drill-down table ----
    recent = df.sort_values('raised_date', ascending=False).head(150)
    recent_records = []
    for _, r in recent.iterrows():
        recent_records.append({
            'plid': str(r.get('plid', '')),
            'rfi_no': str(r.get('rfi_no', '') or ''),
            'subsystem': str(r.get('subsystem', '')),
            'disc': str(r.get('discipline', '')),
            'cat': str(r.get('cat', '')),
            'status': str(r.get('status', '')),
            'date': r['raised_date'].strftime('%Y-%m-%d'),
        })

    return {
        'total': total,
        'status_counts': status_counts,
        'cat_counts': cat_counts,
        'disc_counts': disc_counts,
        'disc_counts_open': disc_counts_open,
        'disc_counts_closed': disc_counts_closed,
        'daily': daily,
        'daily_total': daily_total,
        'weekly': weekly,
        'weekly_total': weekly_total,
        'monthly': monthly,
        'monthly_total': monthly_total,
        'top_subsystems': top_sub.to_dict('records'),
        'recent': recent_records,
    }


# ====================================================================
#  3b) RFI / Inspection Register
# ====================================================================
def build_inspection_data(excel_path):
    if not excel_path:
        return None

    df = pd.read_excel(excel_path, sheet_name='PS-5 EIT INSPECTION REGISTER', header=5)
    df = df.dropna(subset=['Asset - Tag'])

    disc_map = {
        'E':'E','E - Electrical':'E','Electrical':'E',
        'I':'I','I - Instrumentation':'I','I - Instrument':'I',
        'Instrumentation':'I','Instrument':'I',
        'T':'T','T - Telecom':'T','Telecom':'T',
    }
    df['disc'] = df['Discipline'].astype(str).str.strip().map(disc_map).fillna('Other')

    # ---- Normalize RFI Status ----
    def norm_status(v):
        if not isinstance(v, str):
            return 'No RFI Yet'
        v = v.strip().lower()
        if 'punch' in v:
            return 'Accepted with Punch'
        if 'accept' in v:
            return 'Accepted'
        if 'hold' in v:
            return 'Hold by EACOP'
        if 'open' in v:
            return 'Open'
        return 'Other'

    status_col = 'Status Of RFI '
    if status_col not in df.columns:
        for c in df.columns:
            if 'status' in str(c).lower() and 'rfi' in str(c).lower():
                status_col = c
                break
    df['status_norm'] = df[status_col].apply(norm_status)

    # ---- Normalize Inspection Type ----
    def norm_type(v):
        if not isinstance(v, str):
            return None
        v = v.strip().lower()
        if 'pull' in v:
            return 'Pulling'
        if 'test' in v:
            return 'Testing'
        if 'install' in v:
            return 'Installation'
        if 'gland' in v:
            return 'Glanding & Termination'
        return 'Other'

    type_col = 'INSPECTION TYPE (PULLING/CABLE TESTING/ EQUIPMENT INSTALLATION)'
    df['type_norm'] = df[type_col].apply(norm_type)

    # RFI submitted = has QC RFI# (INSTALLATION/TESTING) or QC RFI#.1 (GLANDING & TERMINATION)
    rfi_mask = df['QC RFI#'].notna() | df['QC RFI#.1'].notna()
    total_assets = len(df)
    total_rfi_submitted = int(rfi_mask.sum())

    status_counts = df['status_norm'].value_counts().to_dict()
    type_counts = df['type_norm'].value_counts(dropna=True).to_dict()

    # ---- status by discipline ----
    status_by_disc = (df.groupby(['disc', 'status_norm']).size()
                      .unstack(fill_value=0).reindex(['E', 'I', 'T'], fill_value=0))
    statuses = list(status_counts.keys())
    for s in statuses:
        if s not in status_by_disc.columns:
            status_by_disc[s] = 0
    status_by_disc = status_by_disc[statuses]
    status_by_disc_records = status_by_disc.reset_index().rename(columns={'disc': 'label'}).to_dict('records')

    # ---- RFI trend (Inspection Date) by discipline ----
    df['Inspection Date'] = pd.to_datetime(df['Inspection Date'], errors='coerce')
    dated = df.dropna(subset=['Inspection Date']).copy()

    def pivot_disc(data, col):
        if data.empty:
            return []
        p = data.groupby([col, 'disc']).size().unstack(fill_value=0)
        for d in ['E', 'I', 'T']:
            if d not in p.columns:
                p[d] = 0
        p['Total'] = p[['E', 'I', 'T']].sum(axis=1)
        p = p[['E', 'I', 'T', 'Total']]
        return p.reset_index().rename(columns={col: 'label'}).to_dict('records')

    now = dated['Inspection Date'].max() if not dated.empty else pd.Timestamp.now()

    dated['date'] = dated['Inspection Date'].dt.strftime('%Y-%m-%d')
    daily = pivot_disc(dated, 'date')
    daily = sorted(daily, key=lambda r: r['label'])[-30:]
    daily_total = int((dated['date'] == now.strftime('%Y-%m-%d')).sum())

    dated['week'] = dated['Inspection Date'].dt.to_period('W').apply(
        lambda r: r.start_time.strftime('%Y-%m-%d'))
    weekly = pivot_disc(dated, 'week')
    weekly = sorted(weekly, key=lambda r: r['label'])[-16:]
    week_start = (now - pd.Timedelta(days=now.weekday())).normalize()
    weekly_total = int((dated['Inspection Date'] >= week_start).sum())

    dated['month'] = dated['Inspection Date'].dt.strftime('%Y-%m')
    monthly = pivot_disc(dated, 'month')
    monthly = sorted(monthly, key=lambda r: r['label'])
    monthly_total = int(((dated['Inspection Date'].dt.year == now.year) &
                          (dated['Inspection Date'].dt.month == now.month)).sum())

    # ---- top subsystems by RFI count ----
    sub_col = 'Systemization - Subsystem (Summary)'
    top_sub = (dated.groupby(sub_col).size().sort_values(ascending=False)
               .head(10).reset_index())
    top_sub.columns = ['label', 'count']

    # ---- recent RFI records for drill-down table ----
    recent = dated.sort_values('Inspection Date', ascending=False).head(100)
    table_col = 'Asset - Tag'
    recent_records = []
    for _, r in recent.iterrows():
        rfi_val = str(r.get('QC RFI#', '') or '') if pd.notna(r.get('QC RFI#')) else str(r.get('QC RFI#.1', '') or '')
        recent_records.append({
            'asset': str(r.get(table_col, '')),
            'rfi_no': rfi_val,
            'disc': r.get('disc', ''),
            'status': r.get('status_norm', ''),
            'date': r['Inspection Date'].strftime('%Y-%m-%d'),
        })

    # ---- RFI Inspection Summary (by Discipline, 3 columns: Laying / Testing / Termination) ----
    def sum_rfi(sub_df):
        laying_submitted = int(sub_df[sub_df['type_norm'].isin(['Pulling', 'Installation'])]['QC RFI#'].notna().sum())
        laying_accepted = int(sub_df[sub_df['type_norm'].isin(['Pulling', 'Installation']) & (sub_df['status_norm'].isin(['Accepted', 'Accepted with Punch']))]['QC RFI#'].notna().sum())
        testing_submitted = int(sub_df[sub_df['type_norm'] == 'Testing']['QC RFI#'].notna().sum())
        testing_accepted = int(sub_df[(sub_df['type_norm'] == 'Testing') & (sub_df['status_norm'].isin(['Accepted', 'Accepted with Punch']))]['QC RFI#'].notna().sum())
        term_submitted = int(sub_df['QC RFI#.1'].notna().sum())
        term_accepted = int(sub_df[sub_df['QC RFI#.1'].notna()].apply(
            lambda r: norm_status(r.get(' RFI STATUS', None)) in ('Accepted', 'Accepted with Punch') if pd.notna(r.get(' RFI STATUS', None)) else False, axis=1
        ).sum())
        return {
            'discipline': sub_df['disc'].iloc[0] if not sub_df.empty else 'Other',
            'assets': sub_df['Asset - Tag'].nunique(),
            'laying_submitted': laying_submitted,
            'laying_accepted': laying_accepted,
            'testing_submitted': testing_submitted,
            'testing_accepted': testing_accepted,
            'term_submitted': term_submitted,
            'term_accepted': term_accepted,
        }

    rfi_summary_rows = []
    for disc_code in ['E', 'I', 'T']:
        sub_df = df[df['disc'] == disc_code]
        if sub_df.empty:
            continue
        rfi_summary_rows.append(sum_rfi(sub_df))

    total_laying_sub = sum(r['laying_submitted'] for r in rfi_summary_rows)
    total_laying_acc = sum(r['laying_accepted'] for r in rfi_summary_rows)
    total_testing_sub = sum(r['testing_submitted'] for r in rfi_summary_rows)
    total_testing_acc = sum(r['testing_accepted'] for r in rfi_summary_rows)
    total_term_sub = sum(r['term_submitted'] for r in rfi_summary_rows)
    total_term_acc = sum(r['term_accepted'] for r in rfi_summary_rows)

    return {
        'total_assets': total_assets,
        'total_rfi': total_rfi_submitted,
        'status_counts': status_counts,
        'type_counts': type_counts,
        'status_by_disc': status_by_disc_records,
        'daily': daily, 'daily_total': daily_total,
        'weekly': weekly, 'weekly_total': weekly_total,
        'monthly': monthly, 'monthly_total': monthly_total,
        'top_subsystems': top_sub.to_dict('records'),
        'recent': recent_records,
        'inspection_summary': {
            'rows': rfi_summary_rows,
            'totals': {
                'total_assets': sum(r['assets'] for r in rfi_summary_rows),
                'laying_submitted': total_laying_sub,
                'laying_accepted': total_laying_acc,
                'testing_submitted': total_testing_sub,
                'testing_accepted': total_testing_acc,
                'term_submitted': total_term_sub,
                'term_accepted': total_term_acc,
            }
        } if rfi_summary_rows else None,
    }


# ====================================================================
#  3d) Universal Search Index (Asset / Task ID -> everything)
# ====================================================================
def build_search_index(ov_path, punch_path, rfi_path):
    """
    يبني فهرس موحّد: لكل Asset Tag، كل السجلات المرتبطة بيه من:
    - ovTasks (Exported from SC): Task ID, Task State, Closing Date, Discipline, Task Type, Description
    - Inspection Register: RFI No, Status, Inspection Date, Discipline, Inspection Type
    - Punch List: PL ID, RFI No, Discipline, Category, Status, Raised Date
    """
    index = {}

    def add(tag, source, record):
        tag = str(tag).strip()
        if not tag or tag.lower() == 'nan':
            return
        index.setdefault(tag, {'itr': [], 'rfi': [], 'punch': []})
        index[tag][source].append(record)

    # ---- ovTasks - المصدر الرئيسي (يتجدد يومياً) ----
    df = safe_read_excel(ov_path, sheet_name='Exported from SC')

    # Build cable tag -> company map for precom info
    cable_forms = [
        'HV Power Cable', 'LV Power Cable', 'Multicore Electrical Control Cable',
        'Multi/Single-Core Instrument Control Cable', 'Instrument Power Cable',
        'Telecom Cable Power/Control/Multicore', 'RJ45 Cable', 'Fibre Optic cable',
        'Coaxial Cable', 'Tape Trace Heating', 'Cable Tray/Ladder']
    cdf = df[df['Test Form - Description'].isin(cable_forms)]
    tag_company = {}
    for _, r in cdf.iterrows():
        tag = str(r['Asset - Tag']).strip() if pd.notna(r['Asset - Tag']) else ''
        if tag and tag.lower() != 'nan' and tag not in tag_company:
            tag_company[tag] = str(r['Responsible Company (Summary)']) if pd.notna(r['Responsible Company (Summary)']) else ''

    # إحصائيات سريعة
    total_tasks   = len(df)
    closed_tasks  = (df['Task State'] == 'Closed').sum()
    print(f"  ovTasks: {total_tasks} tasks | {closed_tasks} closed")

    # توحيد التخصصات
    disc_norm_map = {
        'E - Electrical':'E','Electrical':'E',
        'I - Instrumentation':'I','I - Instrument':'I','Instrumentation':'I','Instrument':'I',
        'T - Telecom':'T','Telecom':'T',
    }

    for _, r in df.iterrows():
        tag = r['Asset - Tag']
        if pd.isna(tag):
            continue
        closing   = r['Closing Date']
        raw_disc  = str(r['Discipline (Summary)']).strip() if pd.notna(r['Discipline (Summary)']) else ''
        task_id   = str(r['Task ID']) if pd.notna(r['Task ID']) else ''
        state     = str(r['Task State']) if pd.notna(r['Task State']) else ''
        is_closed = state == 'Closed'

        add(tag, 'itr', {
            'id':    task_id,
            'st':    state,
            'closed': is_closed,
            'ty':    str(r['Task Type (Name)'])   if pd.notna(r['Task Type (Name)'])   else '',
            'd':     disc_norm_map.get(raw_disc, raw_disc[:1] if raw_disc else ''),
            'disc':  disc_norm_map.get(raw_disc, raw_disc),
            'cd':    pd.to_datetime(closing).strftime('%Y-%m-%d') if pd.notna(closing) else '',
            'sub':   str(r['Systemization - Subsystem (Summary)']) if pd.notna(r.get('Systemization - Subsystem (Summary)')) else '',
            'ms':    str(r['Subsystem Priority']) if pd.notna(r.get('Subsystem Priority')) else '',
            'res':   str(r['Responsible Company (Summary)']) if pd.notna(r.get('Responsible Company (Summary)')) else '',
            'precom': tag_company.get(str(tag).strip(), ''),
        })

    # ---- Inspection Register ----
    if rfi_path:
        rdf = pd.read_excel(rfi_path, sheet_name='PS-5 EIT INSPECTION REGISTER', header=5)
        rdf = rdf.dropna(subset=['Asset - Tag'])

        def norm_status(v):
            if not isinstance(v, str):
                return 'No RFI Yet'
            v = v.strip().lower()
            if 'punch' in v: return 'Accepted with Punch'
            if 'accept' in v: return 'Accepted'
            if 'hold' in v: return 'Hold by EACOP'
            if 'open' in v: return 'Open'
            return 'Other'

        type_col = 'INSPECTION TYPE (PULLING/CABLE TESTING/ EQUIPMENT INSTALLATION)'
        for _, r in rdf.iterrows():
            tag = r['Asset - Tag']
            insp_date = pd.to_datetime(r.get('Inspection Date'), errors='coerce')
            rfi_val = str(r.get('QC RFI#', '') or '') if pd.notna(r.get('QC RFI#')) else str(r.get('QC RFI#.1', '') or '')
            add(tag, 'rfi', {
                'rfi_no': rfi_val,
                'status': norm_status(r.get('Status Of RFI ')),
                'discipline': str(r.get('Discipline', '') or ''),
                'type': str(r.get(type_col, '') or ''),
                'date': insp_date.strftime('%Y-%m-%d') if pd.notna(insp_date) else '',
            })

    # ---- Punch List (MP Register via RFI No) ----
    if punch_path:
        rfi_to_tags = {}
        for tag, srcs in index.items():
            for rr in srcs['rfi']:
                if rr['rfi_no']:
                    rfi_to_tags.setdefault(rr['rfi_no'], set()).add(tag)

        raw = pd.read_excel(punch_path, sheet_name='MP Register', header=None, skiprows=6)
        cols = {1: 'plid', 3: 'rfi_no', 5: 'subsystem', 8: 'desc',
                9: 'cat', 10: 'discipline', 13: 'raised_date', 20: 'status'}
        pdf_ = raw[list(cols.keys())].rename(columns=cols)
        pdf_ = pdf_.dropna(subset=['plid'])
        for _, r in pdf_.iterrows():
            rfi_no = str(r.get('rfi_no', '') or '')
            raised = pd.to_datetime(r.get('raised_date'), errors='coerce')
            record = {
                'plid': str(r.get('plid', '')),
                'subsystem': str(r.get('subsystem', '') or ''),
                'discipline': str(r.get('discipline', '') or ''),
                'category': str(r.get('cat', '') or ''),
                'status': str(r.get('status', '') or 'Open'),
                'date': raised.strftime('%Y-%m-%d') if pd.notna(raised) else '',
                'rfi_no': rfi_no,
                'desc': str(r.get('desc', '') or '').strip(),
            }
            tags = rfi_to_tags.get(rfi_no)
            if tags:
                for tag in tags:
                    index[tag]['punch'].append(record)
            else:
                add(f"[Unlinked] {rfi_no or record['plid']}", 'punch', record)

    # ---- Punch List (SC Export via direct Asset Tag) ----
    sc_punch_path = find_file(['ovPunchlist'])
    if sc_punch_path:
        try:
            sc_raw = pd.read_excel(sc_punch_path, sheet_name='Exported from SC', header=None)
            sc_raw = sc_raw.iloc[1:]  # skip header row
            sc_cols = {0: 'plid', 1: 'asset_tag', 3: 'cat', 4: 'discipline',
                       5: 'desc', 7: 'subsystem', 9: 'status', 12: 'raised_date'}
            sc_df = sc_raw[list(sc_cols.keys())].rename(columns=sc_cols)
            sc_df = sc_df.dropna(subset=['plid', 'asset_tag'])
            # توحيد Status (Originated/Completed -> Open, Closed -> Closed)
            status_norm = {'originated': 'Open', 'completed': 'Closed', 'closed': 'Closed'}
            for _, r in sc_df.iterrows():
                tag = str(r['asset_tag']).strip()
                if not tag or tag.lower() == 'nan':
                    continue
                raised = pd.to_datetime(r.get('raised_date'), errors='coerce')
                raw_st = str(r.get('status', 'Open') or 'Open').strip().lower()
                record = {
                    'plid': str(r.get('plid', '')),
                    'subsystem': str(r.get('subsystem', '') or ''),
                    'discipline': str(r.get('discipline', '') or ''),
                    'category': str(r.get('cat', '') or ''),
                    'status': status_norm.get(raw_st, 'Open'),
                    'date': raised.strftime('%Y-%m-%d') if pd.notna(raised) else '',
                    'rfi_no': '',
                    'desc': str(r.get('desc', '') or '').strip(),
                }
                index.setdefault(tag, {'itr': [], 'rfi': [], 'punch': []})
                index[tag]['punch'].append(record)
            print(f"  SC Punch: Added {len(sc_df)} direct-asset punch items")
        except Exception as e:
            print(f"  Warning: Could not read SC punch file {sc_punch_path}: {e}")

    # تقليل الحجم: شيل الـ tags اللي عندها سجل itr واحد فقط بدون rfi/punch ومش closed (قليلة القيمة للبحث)
    pruned = {}
    for tag, srcs in index.items():
        if srcs['rfi'] or srcs['punch'] or any(it['st']=='Closed' for it in srcs['itr']) or len(srcs['itr'])>1:
            pruned[tag] = srcs
    # إزالة التكرار في punch لكل tag وتحديد سقف عشان الحجم
    for tag, srcs in pruned.items():
        seen = set()
        uniq = []
        for p in srcs['punch']:
            key = p['plid']
            if key not in seen:
                seen.add(key)
                uniq.append(p)
        srcs['punch'] = uniq[:10]
        srcs['itr'] = srcs['itr'][:10]

    print(f"Search index built: {len(pruned)} unique Asset Tags (from {len(index)})")

    # ---- Build reverse maps ----
    rfi_map = {}   # RFI No -> [asset tags]
    sub_map = {}   # Subsystem -> [asset tags]
    pri_map = {}   # Milestone Priority -> [asset tags]
    tid_map = {}   # Task ID -> asset tag

    for tag, srcs in pruned.items():
        # RFI map
        for rr in srcs['rfi']:
            rno = rr.get('rfi_no','').strip()
            if rno:
                rfi_map.setdefault(rno, [])
                if tag not in rfi_map[rno]:
                    rfi_map[rno].append(tag)
        # Task ID map
        for it in srcs['itr']:
            tid = it.get('id','').strip()
            if tid: tid_map[tid] = tag
        # Subsystem map
        for it in srcs['itr']:
            sub = it.get('sub','').strip()
            if sub:
                sub_map.setdefault(sub, [])
                if tag not in sub_map[sub]:
                    sub_map[sub].append(tag)
        # Priority map
        for it in srcs['itr']:
            pri = it.get('ms','').strip()
            if pri:
                pri_map.setdefault(pri, [])
                if tag not in pri_map[pri]:
                    pri_map[pri].append(tag)

    return {
        'index': pruned,
        'rfi_map': rfi_map,
        'sub_map': sub_map,
        'pri_map': pri_map,
        'tid_map': tid_map,
    }

# ====================================================================
def build_completed_rfi_table(ov_path, punch_path, rfi_path):
    """
    تبني جدول RFIs اللي كل punch items تبعها اتقفلت بالكامل.
    Columns: Asset Tag, Task IDs, RFI No, Total PL, Closed PL, Open PL, Status
    """
    if not punch_path or not rfi_path:
        return []

    # 1) Punch Register: group by RFI
    raw = pd.read_excel(punch_path, sheet_name='MP Register', header=None, skiprows=6)
    pdf = raw.iloc[:, [1, 3, 20]].copy()
    pdf.columns = ['plid', 'rfi_no', 'status']
    pdf = pdf.dropna(subset=['plid'])
    pdf['rfi_no'] = pdf['rfi_no'].fillna('')
    pdf['status'] = pdf['status'].astype(str).str.strip().str.upper()

    rfi_groups = pdf.groupby('rfi_no')
    rfi_punch_summary = {}
    for rfi, grp in rfi_groups:
        if not rfi:
            continue
        total = len(grp)
        closed = (grp['status'] == 'CLOSED').sum()
        rfi_punch_summary[rfi] = {
            'total_pl': total,
            'closed_pl': closed,
            'open_pl': total - closed,
            'all_closed': total == closed
    }

    # 2) Inspection Register: RFI No -> Asset Tags
    rdf = pd.read_excel(rfi_path, sheet_name='PS-5 EIT INSPECTION REGISTER', header=5)
    rdf = rdf.dropna(subset=['Asset - Tag'])

    rfi_to_assets = {}
    for _, r in rdf.iterrows():
        tag = r['Asset - Tag']
        rfi_val = str(r.get('QC RFI#', '') or '') if pd.notna(r.get('QC RFI#')) else str(r.get('QC RFI#.1', '') or '')
        if rfi_val:
            rfi_to_assets.setdefault(rfi_val.strip(), set()).add(str(tag).strip())

    # 3) OV Tasks: Asset Tag -> Task IDs
    df = safe_read_excel(ov_path, sheet_name='Exported from SC')
    asset_to_tasks = {}
    for _, r in df.iterrows():
        tag = r['Asset - Tag']
        if pd.isna(tag):
            continue
        task_id = str(r['Task ID']) if pd.notna(r['Task ID']) else ''
        asset_to_tasks.setdefault(str(tag).strip(), []).append(task_id)

    # 4) Build final table rows
    rows = []
    for rfi, summary in rfi_punch_summary.items():
        assets = rfi_to_assets.get(rfi, [])
        if not assets:
            continue
        for asset in assets:
            tasks = asset_to_tasks.get(asset, [])
            rows.append({
                'asset_tag': asset,
                'task_ids': ', '.join(tasks[:8]) if tasks else '-',
                'rfi_no': rfi,
                'total_pl': int(summary['total_pl']),
                'closed_pl': int(summary['closed_pl']),
                'open_pl': int(summary['open_pl']),
                'all_closed': bool(summary['all_closed'])
            })

    # Sort: open ones first (red), then closed (green)
    rows.sort(key=lambda r: (r['all_closed'], r['asset_tag']))
    print(f"  Completed RFI table: {len(rows)} rows ({sum(1 for r in rows if r['all_closed'])} all closed)")
    return rows


# ====================================================================
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en" dir="ltr">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="900">
<title>PS5 - CPP AGI Completion Progress Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --bg:#f5f0e8; --bg2:#ede7db; --card:#fffcf7; --card2:#faf6ee; --card3:#f0ead9;
  --border:#d9d0c1; --border2:#c4b9a7;
  --text:#2c2416; --text2:#6b5e4d; --text3:#9a8d7c;
  --gold:#c8940a; --gold2:#a67808; --gold-bg:rgba(200,148,10,0.1);
  --green:#1a8a4a; --green2:#15803d; --green-bg:rgba(26,138,74,0.08);
  --red:#c53030; --red2:#b91c1c; --red-bg:rgba(197,48,48,0.08);
  --blue:#2563eb; --blue2:#1d4ed8; --blue-bg:rgba(37,99,235,0.08);
  --cyan:#0891b2; --purple:#7c3aed; --pink:#ec4899;
  --shadow:0 1px 4px rgba(44,36,22,0.06);
  --shadow-lg:0 8px 30px rgba(44,36,22,0.1);
  --bgmain:var(--bg); --panel:var(--card); --panel2:var(--card2);
  --accent:var(--red); --teal:var(--green); --muted:var(--text2);
}
*{box-sizing:border-box;margin:0;padding:0;}
body{
  font-family:'Inter',system-ui,sans-serif;
  background:var(--bg);color:var(--text);
  min-height:100vh;-webkit-font-smoothing:antialiased;
}

/* ---------- Sticky Header (KENT PLC style) ---------- */
.header{
  background:linear-gradient(135deg,#fffcf7 0%,#faf6ee 50%,#f5f0e8 100%);
  border-bottom:3px solid var(--gold);position:sticky;top:0;z-index:200;
  box-shadow:0 2px 12px rgba(44,36,22,0.08);
}
.header-top{display:flex;align-items:center;justify-content:space-between;padding:18px 36px 8px;flex-wrap:wrap;gap:8px;}
.header h1{font-size:22px;font-weight:800;color:var(--text);}
.header h1 span{color:var(--gold);}
.header-badge{display:flex;gap:8px;align-items:center;}
.header-badge .tag{background:var(--gold-bg);color:var(--gold);font-size:10px;font-weight:700;padding:4px 12px;border-radius:20px;border:1px solid rgba(200,148,10,0.25);}
.header-badge .live{background:var(--green-bg);color:var(--green);font-size:10px;font-weight:700;padding:4px 12px;border-radius:20px;border:1px solid rgba(26,138,74,0.25);animation:pulse 2s infinite;}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}}
.header .subtitle{color:var(--text2);font-size:13px;padding:0 36px 8px;}
.header .meta{display:flex;gap:8px;padding:0 36px 14px;flex-wrap:wrap;align-items:center;}
.header .meta span{font-size:10px;color:var(--text2);background:var(--card2);padding:4px 12px;border-radius:6px;border:1px solid var(--border);font-weight:600;}

/* ---------- Main layout ---------- */
#tab-main-dashboard.tabpage.active{display:flex;align-items:flex-start;}
.sidebar{
  width:248px;flex-shrink:0;min-height:calc(100vh - 150px);
  background:var(--card);color:var(--text);padding:20px 16px;
  position:sticky;top:128px;margin:24px 0 24px 24px;
  border:1px solid var(--border);border-radius:14px;box-shadow:var(--shadow);
}
.sidebar h2{font-size:15px;margin-bottom:16px;display:flex;align-items:center;gap:8px;
  letter-spacing:.5px;color:var(--gold);}
.sidebar label{
  display:flex;align-items:center;gap:8px;padding:8px 8px;border-radius:8px;
  cursor:pointer;font-size:12px;margin-bottom:2px;transition:.15s;color:var(--text2);
}
.sidebar label:hover{background:var(--card2);color:var(--text);}
.sidebar input{accent-color:var(--gold);width:15px;height:15px;}
.sidebar .grp{margin-top:16px;font-size:10px;opacity:.7;text-transform:uppercase;
  letter-spacing:1.2px;border-bottom:1px solid var(--border);padding-bottom:4px;color:var(--gold);}

.main{flex:1;min-width:0;padding:24px 32px;}

/* ---------- Sections ---------- */
.section{display:none;}
.section.active{display:block;animation:panelIn .3s ease;}
@keyframes panelIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}

.kpi-row{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:20px;}
.kpi{
  flex:1;min-width:170px;background:var(--card);border-radius:14px;padding:18px 16px;
  text-align:center;border:1px solid var(--border);
  border-top:4px solid var(--blue);position:relative;overflow:hidden;box-shadow:var(--shadow);
  transition:transform .2s, box-shadow .2s;
}
.kpi:hover{transform:translateY(-4px);box-shadow:var(--shadow-lg);}
.kpi.gold{border-top-color:var(--gold);} .kpi.gold .val{color:var(--gold);}
.kpi.teal{border-top-color:var(--green);} .kpi.teal .val{color:var(--green);}
.kpi.pink{border-top-color:var(--red);} .kpi.pink .val{color:var(--red);}
.kpi.blue{border-top-color:var(--blue);} .kpi.blue .val{color:var(--blue);}
.kpi.purple{border-top-color:var(--purple);} .kpi.purple .val{color:var(--purple);}
.kpi .icon{font-size:22px;margin-bottom:6px;}
.kpi .val{font-size:30px;font-weight:900;letter-spacing:-1px;color:var(--text);}
.kpi .lbl{font-size:11px;color:var(--text3);margin-top:4px;font-weight:600;}

.chart-row{display:flex;gap:18px;flex-wrap:wrap;margin-bottom:20px;}
.chart-card{
  flex:1;min-width:340px;background:var(--card);border-radius:14px;padding:20px;
  border:1px solid var(--border);box-shadow:var(--shadow);transition:box-shadow .25s;
}
.chart-card:hover{box-shadow:var(--shadow-lg);}
.chart-card h3{font-size:11px;color:var(--text3);margin-bottom:12px;font-weight:700;
  text-transform:uppercase;letter-spacing:1px;}
canvas{max-height:300px;}

.section-title{
  font-size:17px;color:var(--text);font-weight:800;margin:10px 0 14px;
  border-left:5px solid var(--gold);padding-left:12px;
  display:flex;align-items:center;gap:8px;
}
.progress-bar{height:9px;background:var(--card3);border-radius:6px;overflow:hidden;margin-top:8px;}
.progress-fill{height:100%;background:linear-gradient(90deg,var(--green),#22c55e);border-radius:6px;}
.footer{text-align:center;padding:16px;color:var(--text3);font-size:12px;}

/* ---------- Tables (light theme) ---------- */
table{color:var(--text);border-collapse:collapse;font-size:12px;}
thead tr{background:var(--card2) !important;}
table th{color:var(--gold);border-bottom:2px solid var(--gold);font-weight:700;
  font-size:10px;text-transform:uppercase;letter-spacing:0.6px;}
tbody tr{border-bottom:1px solid var(--border) !important;}
tbody tr:hover td{background:rgba(200,148,10,0.04);}
input#rfiSearch{background:var(--card);border:2px solid var(--border) !important;color:var(--text);}
input#rfiSearch::placeholder{color:var(--text3);}

/* ---------- EIT ITR Description Table (Excel-style, matches reference sheet) ---------- */
.eit-toolbar{display:flex;justify-content:flex-end;margin-bottom:12px;flex-wrap:wrap;gap:8px;}
.btn-export{
  background:var(--green);color:#fff;border:none;
  padding:9px 18px;border-radius:8px;font-size:12px;font-weight:700;cursor:pointer;
  display:flex;align-items:center;gap:8px;box-shadow:0 4px 14px rgba(26,138,74,.25);
  transition:.15s;
}
.btn-export:hover{transform:translateY(-1px);box-shadow:0 6px 18px rgba(26,138,74,.35);}
.eit-table-wrap{background:var(--card);border-radius:12px;overflow:auto;max-height:640px;
  border:1px solid var(--border);box-shadow:var(--shadow);}
table.eit-table{width:100%;border-collapse:separate;border-spacing:0;font-size:12px;color:var(--text);}
table.eit-table th{
  background:var(--card2);color:var(--gold);font-weight:800;padding:10px 8px;text-align:center;
  border:none;border-bottom:2px solid var(--gold);position:sticky;top:0;z-index:3;
  font-size:10px;text-transform:uppercase;letter-spacing:0.6px;white-space:nowrap;
}
table.eit-table td{padding:7px 10px;border:none;border-bottom:1px solid var(--border);text-align:center;
  color:var(--text);white-space:nowrap;}
table.eit-table td.desc-cell{text-align:left;white-space:normal;}
table.eit-table td.type-cell{font-weight:800;color:var(--text2);white-space:nowrap;}
table.eit-table tr:nth-child(even) td{background:var(--card2);}
table.eit-table tbody tr:hover td{background:rgba(200,148,10,0.06);}
tr.type-E td.type-cell{color:var(--blue);}
tr.type-I td.type-cell{color:#d97706;}
tr.type-T td.type-cell{color:var(--purple);}
table.eit-table td.balance-pos{color:var(--red);font-weight:700;}
table.eit-table td.balance-zero{color:var(--green);font-weight:700;}
tr.eit-total-row td{background:var(--card3) !important;color:var(--text);font-weight:800;
  border-top:2px solid var(--gold);border-bottom:1px solid var(--border);}
tr.eit-total-row td.balance-pos{color:var(--red);}

/* ===== Tab bar (KENT PLC style) ===== */
.tabbar{display:flex;flex-wrap:wrap;gap:4px;max-width:1640px;margin:16px auto 0;
  padding:6px;background:var(--card);border:1px solid var(--border);border-radius:14px;
  box-shadow:var(--shadow);}
.tabbtn{padding:9px 18px;border-radius:10px;cursor:pointer;font-size:11px;font-weight:700;
  color:var(--text3);transition:all .25s;border:none;background:none;
  text-transform:uppercase;letter-spacing:0.8px;display:inline-flex;align-items:center;gap:7px;}
.tabbtn .t-icon{width:22px;height:22px;display:inline-flex;align-items:center;justify-content:center;
  border-radius:7px;background:var(--card2);border:1px solid var(--border);font-size:12px;
  transition:all .25s;}
.tabbtn:hover{background:var(--card2);color:var(--text2);}
.tabbtn:hover .t-icon{background:var(--gold-bg);border-color:rgba(200,148,10,0.4);transform:scale(1.1);}
.tabbtn.active{background:var(--gold);color:#fff;box-shadow:0 2px 10px rgba(200,148,10,0.3);}
.tabbtn.active .t-icon{background:rgba(255,255,255,0.22);border-color:rgba(255,255,255,0.35);color:#fff;}
.tabpage{display:none;}
#tab-main-dashboard.tabpage.active{display:flex;}
.tabpage.active{display:block;}
.eit-page-toolbar{display:flex;align-items:center;gap:10px;margin:16px 0 10px;flex-wrap:wrap;}
.eit-page-toolbar input{padding:10px 14px;border:2px solid var(--border);border-radius:10px;
  background:var(--card);color:var(--text);font-size:13px;flex:1;min-width:200px;
  outline:none;font-family:inherit;transition:border-color .25s, box-shadow .25s;}
.eit-page-toolbar input:focus{border-color:var(--gold);box-shadow:0 0 0 4px rgba(200,148,10,0.1);}
.eit-page-toolbar .count{color:var(--text3);font-size:12px;white-space:nowrap;font-weight:600;}
.eit-page-table-wrap{overflow:auto;max-height:70vh;border:1px solid var(--border);
  border-radius:12px;background:var(--card);box-shadow:var(--shadow);}
.eit-page-table-wrap table{border-collapse:collapse;width:100%;font-size:12px;color:var(--text);}
.eit-page-table-wrap th{position:sticky;top:0;background:var(--card2);color:var(--gold);
  font-weight:700;padding:8px;text-align:left;border-bottom:2px solid var(--gold);
  white-space:nowrap;font-size:10px;text-transform:uppercase;letter-spacing:0.6px;}
.eit-page-table-wrap td{padding:6px 8px;border-bottom:1px solid var(--border);color:var(--text);
  white-space:nowrap;max-width:380px;overflow:hidden;text-overflow:ellipsis;}
.eit-page-table-wrap tr:nth-child(even) td{background:var(--card2);}
.eit-page-table-wrap tr:hover td{background:rgba(200,148,10,0.06);}
.eit-kpi-cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px;margin-bottom:14px;}
.eit-kpi{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:12px;
  text-align:center;box-shadow:var(--shadow);}
.eit-kpi .kpi-v{font-size:22px;font-weight:900;color:var(--green);}
.eit-kpi .kpi-l{font-size:11px;color:var(--text3);margin-top:4px;font-weight:600;}
</style>
</head>
<body>

<div class="header" id="mainHeader">
  <div class="header-top">
    <h1><span>PS5</span> — CPP AGI Completion Progress Dashboard</h1>
    <div class="header-badge">
      <span class="tag">PS5 - Tanzania</span>
      <span class="live">LIVE</span>
    </div>
  </div>
  <div class="subtitle">ITR Closures & Punch List | EACOP PS5 Project | Data last updated: __NOW__</div>
  <div class="meta">
    <span>Tasks: __TOTAL_TASKS__</span>
    <span>Closed: __TOTAL_CLOSED__</span>
    <span>Punch List: __PUNCH_TOTAL__</span>
    <span>RFIs: __RFI_TOTAL__</span>
  </div>
</div>

<div class="tabbar" id="mainTabBar"></div>

<div id="tab-main-dashboard" class="tabpage active">

<!-- ================= SIDEBAR ================= -->
<div class="sidebar">
  <h2>🛢️ PS5 Dashboard</h2>
  <div class="grp">🔍 Search</div>
  <label><input type="checkbox" data-target="sec-search" checked> Asset / Task ID Search</label>

  <div class="grp">General</div>
  <label><input type="checkbox" data-target="sec-kpi" checked> General (KPIs)</label>

  <div class="grp">ITR Closures</div>
  <label><input type="checkbox" data-target="sec-itr-daily" checked> Daily</label>
  <label><input type="checkbox" data-target="sec-itr-weekly" checked> Weekly</label>
  <label><input type="checkbox" data-target="sec-itr-monthly" checked> Monthly</label>
  <label><input type="checkbox" data-target="sec-itr-eit" checked> E&I&T (CPP AGI)</label>
  <label><input type="checkbox" data-target="sec-itr-milestone" checked> Milestone Progress (CPP AGI)</label>
  <label><input type="checkbox" data-target="sec-eit-table" checked> ITR Description Table (E/I/T)</label>

  <div class="grp">Export Tables</div>
  <label><input type="checkbox" data-target="sec-export-punch" checked> 1. Punch Closing Status</label>
  <label><input type="checkbox" data-target="sec-export-cable" checked> 2. Cable / CMT Status</label>
  <label><input type="checkbox" data-target="sec-export-elec" checked> 3. E - Electrical</label>
  <label><input type="checkbox" data-target="sec-export-inst" checked> 4. I - Instrumentation</label>
  <label><input type="checkbox" data-target="sec-export-tele" checked> 5. T - Telecom</label>
  <label><input type="checkbox" data-target="sec-export-rfi" checked> 6. RFI Inspection Summary</label>
  <label><input type="checkbox" data-target="sec-cmt-qc-punch" checked> CMT &amp; QC Punch Summary</label>

  <div class="grp">Punch List</div>
  <label><input type="checkbox" data-target="sec-punch-kpi" checked> Punch List Summary</label>
  <label><input type="checkbox" data-target="sec-punch-daily" checked> Daily</label>
  <label><input type="checkbox" data-target="sec-punch-weekly" checked> Weekly</label>
  <label><input type="checkbox" data-target="sec-punch-monthly" checked> Monthly</label>
  <label><input type="checkbox" data-target="sec-punch-breakdown" checked> Breakdown (Category/Discipline)</label>
  <label><input type="checkbox" data-target="sec-punch-subsystems" checked> Top Subsystems</label>
  <label><input type="checkbox" data-target="sec-punch-recent" checked> Recent Punches (Details)</label>
  <label><input type="checkbox" data-target="sec-punch-tracking" checked> Punch Tracking &amp; Closure</label>

  <div class="grp">RFI / Inspection Register</div>
  <label><input type="checkbox" data-target="sec-rfi-kpi" checked> RFI Summary</label>
  <label><input type="checkbox" data-target="sec-rfi-status" checked> Status Breakdown</label>
  <label><input type="checkbox" data-target="sec-rfi-daily" checked> Daily</label>
  <label><input type="checkbox" data-target="sec-rfi-weekly" checked> Weekly</label>
  <label><input type="checkbox" data-target="sec-rfi-monthly" checked> Monthly</label>
  <label><input type="checkbox" data-target="sec-rfi-type" checked> Inspection Type</label>
  <label><input type="checkbox" data-target="sec-rfi-subsystems" checked> Top Subsystems</label>
  <label><input type="checkbox" data-target="sec-rfi-recent" checked> Recent RFIs (Details)</label>
</div>

<!-- ================= MAIN ================= -->
<div class="main">
  <div class="section active" id="sec-search">
    <div class="section-title">🔍 Universal Asset / Task ID Search — ITR + RFI + Punch (Combined, No Conflicts)</div>
    <div class="chart-card">
      <input id="universalSearch" type="text" placeholder="Type Asset Tag or Task ID (e.g. PS5-25-DM-3100A-CH01 or T-00232-0952)..."
        style="width:100%;padding:12px 16px;border:1px solid var(--border);border-radius:8px;font-size:14px;
        background:var(--panel2);color:var(--text);margin-bottom:6px;">
      <div style="font-size:12px;color:var(--muted);margin-bottom:14px;">
        Library size: <b id="searchLibSize" style="color:var(--teal);"></b> assets indexed.
        Type at least 3 characters.
      </div>
      <div id="searchResults"></div>
    </div>
  </div>

  <!-- ===== KPI Summary ===== -->
  <div class="section active" id="sec-kpi">
    <div class="section-title">Project Overview</div>
    <div class="kpi-row">
      <div class="kpi"><div class="icon">✅</div><div class="val">__TOTAL_CLOSED__</div><div class="lbl">Total ITRs Closed</div></div>
      <div class="kpi gold"><div class="icon">📋</div><div class="val">__TOTAL_TASKS__</div><div class="lbl">Total Project Tasks</div></div>
      <div class="kpi pink"><div class="icon">📌</div><div class="val">__PUNCH_TOTAL__</div><div class="lbl">Total Punch List Items</div></div>
      <div class="kpi blue"><div class="icon">📝</div><div class="val">__RFI_TOTAL__</div><div class="lbl">Total RFIs Submitted</div></div>
      <div class="kpi purple"><div class="icon">⬆️</div><div class="val">__HOURLY_TOTAL__</div><div class="lbl">Today Closed</div></div>
      <div class="kpi teal"><div class="icon">📤</div><div class="val">__HOURLY_SUBMITTED__</div><div class="lbl">Today Submitted</div></div>
    </div>

    <div class="chart-row">
      <div class="chart-card" style="flex:2;">
        <h3>📈 Combined Trend — ITR Closed vs Punch Raised vs RFI Submitted (Daily, Last 30 Days)</h3>
        <canvas id="chartCombinedDaily"></canvas>
      </div>
    </div>
    <div class="chart-row">
      <div class="chart-card" style="flex:2;">
        <h3>📈 Combined Trend — Monthly</h3>
        <canvas id="chartCombinedMonthly"></canvas>
      </div>
    </div>
  </div>

  <!-- ===== ITR Daily ===== -->
  <div class="section active" id="sec-itr-daily">
    <div class="section-title">📅 ITR Closures - Daily (Last 30 Days)</div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartDaily"></canvas></div></div>
  </div>

  <!-- ===== ITR Weekly ===== -->
  <div class="section active" id="sec-itr-weekly">
    <div class="section-title">🗓️ ITR Closures - Weekly</div>
    <div class="kpi-row">
      <div class="kpi"><div class="val">__WEEKLY_TOTAL__</div><div class="lbl">Closed This Week</div></div>
    </div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartWeekly"></canvas></div></div>
  </div>

  <!-- ===== ITR Monthly ===== -->
  <div class="section active" id="sec-itr-monthly">
    <div class="section-title">📆 ITR Closures - Monthly</div>
    <div class="kpi-row">
      <div class="kpi"><div class="val">__MONTHLY_TOTAL__</div><div class="lbl">Closed This Month</div></div>
    </div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartMonthly"></canvas></div></div>
  </div>

  <!-- ===== E&I&T CPP AGI ===== -->
  <div class="section active" id="sec-itr-eit">
    <div class="section-title">⚡ E&I&T — CPP AGI (Completion Rate)</div>
    <div class="kpi-row" id="eitCards"></div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartEIT"></canvas></div></div>
  </div>

  <!-- ===== Milestone Progress CPP AGI ===== -->
  <div class="section active" id="sec-itr-milestone">
    <div class="section-title">🎯 Progress by Milestone — CPP AGI (Completion Rate)</div>
    <div class="kpi-row" id="msCards"></div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartMilestone"></canvas></div></div>
  </div>

  <!-- ===== Today's Closures by Milestone ===== -->
  <div class="section active" id="sec-today-milestone">
    <div class="section-title">📅 Today Closures by Milestone (CPP AGI E/I/T) — Total: <span id="todayMsTotal"></span></div>
    <div class="chart-card">
      <table class="eit-table" id="todayMsTable" style="min-width:600px;">
        <thead>
          <tr>
            <th style="text-align:left;">Milestone</th>
            <th>E — Electrical</th>
            <th>I — Instrumentation</th>
            <th>T — Telecom</th>
            <th>Total</th>
          </tr>
        </thead>
        <tbody id="todayMsBody"></tbody>
      </table>
    </div>
  </div>

  <!-- ===== Subsystem Summary — CPP AGI ===== -->
  <div class="section active" id="sec-subsystem">
    <div class="section-title">📊 Subsystem Summary — CPP AGI E/I/T</div>
    <div class="chart-card">
      <div class="eit-toolbar">
        <button class="btn-export" onclick="exportSubsystemToExcel()">⬇️ Download Excel</button>
      </div>
      <input id="subsystemSearchInput" type="text" placeholder="Search subsystem..."
        style="width:100%;padding:10px 14px;border:1px solid var(--border);border-radius:8px;font-size:13px;background:var(--panel2);color:var(--text);margin-bottom:12px;">
      <div style="max-height:600px;overflow:auto;">
        <table class="eit-table" id="subsystemTable" style="min-width:700px;">
          <thead>
            <tr>
              <th style="text-align:left;">Subsystem</th>
              <th>Discipline</th>
              <th>Total</th>
              <th>Closed</th>
              <th>Open</th>
              <th>% Closed</th>
            </tr>
          </thead>
          <tbody id="subsystemBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- ===== EIT ITR Description Table (Excel-style) ===== -->
  <div class="section active" id="sec-eit-table">
    <div class="section-title">📋 ITR Description Table — E / I / T (<span id="eitCutoffLbl"></span>)</div>
    <div class="eit-toolbar">
      <button class="btn-export" onclick="exportEitTableToExcel()">⬇️ Download Updated Excel</button>
    </div>
    <div class="eit-table-wrap">
      <table class="eit-table" id="eitDescTable">
        <thead>
          <tr>
            <th>Type</th>
            <th style="text-align:left;">ITR Description</th>
            <th>Total Task</th>
            <th>Approved by EACOP</th>
            <th>Balance</th>
            <th>%Closed</th>
            <th id="eitWeekColLbl">This Week</th>
          </tr>
        </thead>
        <tbody id="eitDescTableBody"></tbody>
        <tfoot id="eitDescTableFoot"></tfoot>
      </table>
    </div>
  </div>

  <!-- ===== CMT & QC Punch Summary ===== -->
  <div class="section active" id="sec-cmt-qc-punch">
    <div class="section-title">CMT &amp; QC Punch Summary</div>
    <div class="chart-card">
      <table class="eit-table" id="cmtQcSummaryTable" style="min-width:700px;">
        <thead>
          <tr>
            <th>Source</th><th>Total Assets</th><th>Assets Closed</th><th>Assets Open</th><th>Total Punches</th><th>Punches Closed</th><th>Punches Open</th>
          </tr>
        </thead>
        <tbody id="cmtQcSummaryBody"></tbody>
      </table>
    </div>
    <div class="section-title" style="margin-top:20px;">Per-Asset Detail — Search &amp; Filter</div>
    <div class="chart-card">
      <input id="cmtQcSearchInput" type="text" placeholder="Search by Asset / Subsystem / Description..."
        style="width:100%;padding:10px 14px;border:1px solid var(--border);border-radius:8px;font-size:13px;background:var(--panel2);color:var(--text);margin-bottom:12px;">
      <div style="max-height:500px;overflow:auto;">
        <table id="cmtQcDetailTable" style="width:100%;border-collapse:collapse;font-size:12px;">
          <thead>
            <tr style="background:var(--panel2);color:var(--teal);position:sticky;top:0;z-index:1;">
              <th style="padding:6px;text-align:left;">Asset</th>
              <th style="padding:6px;text-align:left;">Subsystem</th>
              <th style="padding:6px;text-align:center;">Priority</th>
              <th style="padding:6px;text-align:center;">Discipline</th>
              <th style="padding:6px;text-align:center;">CMT (A,B,C)</th>
              <th style="padding:6px;text-align:left;">CMT Description</th>
              <th style="padding:6px;text-align:center;">QC (A,B,C)</th>
              <th style="padding:6px;text-align:left;">QC Description</th>
            </tr>
          </thead>
          <tbody id="cmtQcDetailBody"></tbody>
        </table>
      </div>
      <div id="cmtQcDetailCount" style="margin-top:8px;font-size:12px;color:var(--muted);text-align:right;"></div>
      <div style="margin-top:10px;">
        <button class="btn-export" onclick="exportCmtQcDetailExcel()">⬇️ Download Per-Asset Detail Excel</button>
      </div>
    </div>
  </div>

  <!-- ===== Export: Punch Closing Status ===== -->
  <div class="section active" id="sec-export-punch">
    <div class="section-title">1. Punch Closing Status &mdash; By Discipline</div>
    <div class="eit-toolbar">
      <button class="btn-export" onclick="exportPunchStatusExcel()">Download Punch Status Excel</button>
    </div>
    <div class="eit-table-wrap">
      <table class="eit-table" id="punchStatusTable" style="min-width:500px;">
        <thead>
          <tr>
            <th>Discipline</th><th>Total</th><th>Open</th><th>Closed</th><th>% Closed</th>
          </tr>
        </thead>
        <tbody id="punchStatusBody"></tbody>
        <tfoot id="punchStatusFoot"></tfoot>
      </table>
    </div>
  </div>

  <!-- ===== Export: Cable CMT ===== -->
  <div class="section active" id="sec-export-cable">
    <div class="section-title">2. Cable / CMT &mdash; Closed Status &amp; Backlog</div>
    <div class="eit-toolbar">
      <button class="btn-export" onclick="exportCableCmtExcelOrig()">Download Cable CMT Excel</button>
    </div>
    <div class="eit-table-wrap">
      <table class="eit-table" id="cableCmtOrigTable" style="min-width:700px;">
        <thead>
          <tr>
            <th>Discipline</th><th>Total</th><th>CMT Close</th><th>CMT Open</th><th>Static Close</th><th>Static Open</th><th>% CMT</th><th>% Static</th>
          </tr>
        </thead>
        <tbody id="cableCmtOrigBody"></tbody>
        <tfoot id="cableCmtOrigFoot"></tfoot>
      </table>
    </div>
    <div class="section-title" style="margin-top:20px;font-size:16px;">Conformity Check &amp; Static Test Summary by Scope (E / I / T)</div>
    <div class="eit-toolbar">
      <button class="btn-export" onclick="exportCableScopeExcel()">⬇️ Download Scope Summary Excel</button>
    </div>
    <div class="eit-table-wrap">
      <table class="eit-table" id="cableScopeTable" style="min-width:500px;font-size:11px;">
        <thead>
          <tr>
            <th style="text-align:left;">Scope</th>
            <th style="text-align:center;">Discipline</th>
            <th style="text-align:center;">Total</th>
            <th style="text-align:center;">CMT Close</th>
            <th style="text-align:center;">Static Close</th>
          </tr>
        </thead>
        <tbody id="cableScopeBody"></tbody>
      </table>
    </div>
  </div>

  <!-- ===== Export: E - Electrical ===== -->
    <div class="section active" id="sec-export-elec">
    <div class="section-title">3. E - Electrical &mdash; Status &amp; Backlog (by Description)</div>
    <div class="eit-toolbar">
      <button class="btn-export" onclick="exportDescExcel('E')">Download E - Electrical Excel</button>
    </div>
    <div class="eit-table-wrap">
      <table class="eit-table" id="elecDescTable" style="min-width:600px;border-collapse:separate;border-spacing:0;border-radius:8px;overflow:hidden;">
        <thead id="elecDescHead">
          <tr><th colspan="6" style="background:linear-gradient(135deg,#1e3a5f,#2563eb);color:#fff;padding:14px 12px;font-size:14px;font-weight:700;letter-spacing:1px;border:1px solid #1a252f;">⚡ E - Electrical &mdash; CPP AGI Progress (by Description)</th></tr>
        </thead>
        <tbody id="elecDescBody"></tbody>
      </table>
    </div>
  </div>

  <!-- ===== Export: I - Instrumentation ===== -->
    <div class="section active" id="sec-export-inst">
    <div class="section-title">4. I - Instrumentation &mdash; Status &amp; Backlog (by Description)</div>
    <div class="eit-toolbar">
      <button class="btn-export" onclick="exportDescExcel('I')">Download I - Instrumentation Excel</button>
    </div>
    <div class="eit-table-wrap">
      <table class="eit-table" id="instDescTable" style="min-width:600px;border-collapse:separate;border-spacing:0;border-radius:8px;overflow:hidden;">
        <thead id="instDescHead">
          <tr><th colspan="6" style="background:linear-gradient(135deg,#2d5016,#7c3aed);color:#fff;padding:14px 12px;font-size:14px;font-weight:700;letter-spacing:1px;border:1px solid #1e3a0e;">🔧 I - Instrumentation &mdash; CPP AGI Progress (by Description)</th></tr>
        </thead>
        <tbody id="instDescBody"></tbody>
      </table>
    </div>
  </div>

  <!-- ===== Export: T - Telecom ===== -->
    <div class="section active" id="sec-export-tele">
    <div class="section-title">5. T - Telecom &mdash; Status &amp; Backlog (by Description)</div>
    <div class="eit-toolbar">
      <button class="btn-export" onclick="exportDescExcel('T')">Download T - Telecom Excel</button>
    </div>
    <div class="eit-table-wrap">
      <table class="eit-table" id="teleDescTable" style="min-width:600px;border-collapse:separate;border-spacing:0;border-radius:8px;overflow:hidden;">
        <thead id="teleDescHead">
          <tr><th colspan="6" style="background:linear-gradient(135deg,#5a5a5a,#0891b2);color:#fff;padding:14px 12px;font-size:14px;font-weight:700;letter-spacing:1px;border:1px solid #444;">📡 T - Telecom &mdash; CPP AGI Progress (by Description)</th></tr>
        </thead>
        <tbody id="teleDescBody"></tbody>
      </table>
    </div>
  </div>

  <!-- ===== Export: RFI Inspection Summary ===== -->
  <div class="section active" id="sec-export-rfi">
    <div class="section-title">6. RFI Inspection Summary — By Discipline</div>
    <div class="eit-toolbar">
      <button class="btn-export" onclick="exportRfiInspectionExcel()">⬇️ Download RFI Inspection Excel</button>
    </div>
    <div class="eit-table-wrap">
      <table class="eit-table" id="rfiInspectionTable" style="min-width:700px;font-size:12px;">
        <thead>
          <tr>
            <th>Discipline</th>
            <th>Assets</th>
            <th>LAYING RFI ✓</th>
            <th>TESTING RFI ✓</th>
            <th>TERMINATION RFI ✓</th>
          </tr>
        </thead>
        <tbody id="rfiInspectionBody"></tbody>
        <tfoot id="rfiInspectionFoot"></tfoot>
      </table>
    </div>
    <div class="section-title" style="margin-top:20px;font-size:15px;">📋 RFI Detail — From Pre_Com_Cable_ITR_Tracker (PS5)</div>
    <div class="eit-toolbar">
      <input id="rfiDetailSearch" type="text" placeholder="🔍 Search Asset Tag / Subsystem..." style="flex:1;max-width:300px;padding:6px 10px;border:1px solid #ccc;border-radius:6px;font-size:12px;">
      <button class="btn-export" onclick="exportRfiDetailExcel()">⬇️ Download Excel</button>
    </div>
    <div class="eit-table-wrap" style="max-height:450px;overflow:auto;">
      <table class="eit-table" id="rfiDetailTable" style="min-width:800px;font-size:10px;">
        <thead>
          <tr style="position:sticky;top:0;z-index:2;background:var(--panel2);color:var(--teal);">
            <th style="text-align:left;padding:4px 5px;">Asset Tag</th>
            <th style="text-align:left;padding:4px 5px;">Subsystem</th>
            <th style="padding:4px 5px;">Disc</th>
            <th style="text-align:left;padding:4px 5px;">Description</th>
            <th style="padding:4px 5px;">Scope</th>
            <th style="padding:4px 5px;">LAYING RFI</th>
            <th style="padding:4px 5px;">TESTING RFI</th>
            <th style="padding:4px 5px;">TERMINATION RFI</th>
          </tr>
        </thead>
        <tbody id="rfiDetailBody"></tbody>
      </table>
    </div>
    <div id="rfiDetailCount" style="margin-top:6px;font-size:11px;color:var(--muted);text-align:right;"></div>
  </div>

  <!-- ===== Punch List KPI ===== -->
  <div class="section active" id="sec-punch-kpi">
    <div class="section-title">📌 Punch List — Status Summary</div>
    <div class="chart-row">
      <div class="chart-card"><canvas id="chartPunchStatus"></canvas></div>
    </div>
  </div>

  <!-- ===== Punch Daily ===== -->
  <div class="section active" id="sec-punch-daily">
    <div class="section-title">📅 Punch List - Daily (Last 30 Days)</div>
    <div class="kpi-row"><div class="kpi pink"><div class="val">__PUNCH_DAILY_TOTAL__</div><div class="lbl">Raised Today</div></div></div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartPunchDaily"></canvas></div></div>
  </div>

  <!-- ===== Punch Weekly ===== -->
  <div class="section active" id="sec-punch-weekly">
    <div class="section-title">🗓️ Punch List - Weekly</div>
    <div class="kpi-row"><div class="kpi pink"><div class="val">__PUNCH_WEEKLY_TOTAL__</div><div class="lbl">Raised This Week</div></div></div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartPunchWeekly"></canvas></div></div>
  </div>

  <!-- ===== Punch Monthly ===== -->
  <div class="section active" id="sec-punch-monthly">
    <div class="section-title">📆 Punch List - Monthly</div>
    <div class="kpi-row"><div class="kpi pink"><div class="val">__PUNCH_MONTHLY_TOTAL__</div><div class="lbl">Raised This Month</div></div></div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartPunchMonthly"></canvas></div></div>
  </div>

  <!-- ===== Punch Breakdown ===== -->
  <div class="section active" id="sec-punch-breakdown">
    <div class="section-title">🧩 Punch List - Breakdown</div>
    <div class="chart-row">
      <div class="chart-card"><h3>By Category</h3><canvas id="chartPunchCat"></canvas></div>
      <div class="chart-card"><h3>By Discipline</h3><canvas id="chartPunchDisc"></canvas></div>
    </div>
  </div>

  <!-- ===== Punch Subsystems ===== -->
  <div class="section active" id="sec-punch-subsystems">
    <div class="section-title">🏗️ Top 10 Subsystems (Punch Items)</div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartPunchSub"></canvas></div></div>
  </div>

  <!-- ===== Punch Recent Details ===== -->
  <div class="section active" id="sec-punch-recent">
    <div class="section-title">🔍 Recent Punches — Filter by Category &amp; Status (Last 300)</div>
    <div class="chart-card">
      <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;align-items:center;">
        <input id="punchSearchInput" type="text" placeholder="Filter by PL ID / Subsystem / Description..."
          style="flex:1;min-width:220px;padding:10px 14px;border:1px solid var(--border);border-radius:8px;
          font-size:13px;background:var(--panel2);color:var(--text);">
        <div id="punchCatBtns" style="display:flex;gap:6px;flex-wrap:wrap;"></div>
      </div>
      <div style="max-height:500px;overflow:auto;">
        <table id="punchRecentTable" style="width:100%;border-collapse:collapse;font-size:13px;">
          <thead>
            <tr style="background:var(--panel2);color:var(--teal);position:sticky;top:0;z-index:1;">
              <th style="padding:8px;text-align:left;">PL ID</th>
              <th style="padding:8px;text-align:center;">Category</th>
              <th style="padding:8px;text-align:left;">Discipline</th>
              <th style="padding:8px;text-align:center;">Status</th>
              <th style="padding:8px;text-align:left;">RFI No</th>
              <th style="padding:8px;text-align:left;">Subsystem</th>
              <th style="padding:8px;text-align:left;">Raised Date</th>
              <th style="padding:8px;text-align:left;">Description</th>
            </tr>
          </thead>
          <tbody id="punchRecentBody"></tbody>
        </table>
      </div>
      <div id="punchRecentCount" style="margin-top:10px;font-size:12px;color:var(--muted);text-align:right;"></div>
    </div>
  </div>

  <!-- ===== RFI Summary ===== -->
  <div class="section active" id="sec-rfi-kpi">
    <div class="section-title">📝 RFI / Inspection Register — Summary</div>
    <div class="kpi-row" id="rfiKpiCards"></div>
  </div>

  <!-- ===== RFI Status ===== -->
  <div class="section active" id="sec-rfi-status">
    <div class="section-title">✅ RFI Status Breakdown</div>
    <div class="chart-row">
      <div class="chart-card"><h3>Overall Status</h3><canvas id="chartRfiStatus"></canvas></div>
      <div class="chart-card" style="flex:2;"><h3>Status by Discipline</h3><canvas id="chartRfiStatusDisc"></canvas></div>
    </div>
  </div>

  <!-- ===== RFI Daily ===== -->
  <div class="section active" id="sec-rfi-daily">
    <div class="section-title">📅 RFI Submitted - Daily (Last 30 Days)</div>
    <div class="kpi-row"><div class="kpi blue"><div class="val" id="rfiDailyTotal">0</div><div class="lbl">RFIs Today</div></div></div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartRfiDaily"></canvas></div></div>
  </div>

  <!-- ===== RFI Weekly ===== -->
  <div class="section active" id="sec-rfi-weekly">
    <div class="section-title">🗓️ RFI Submitted - Weekly</div>
    <div class="kpi-row"><div class="kpi blue"><div class="val" id="rfiWeeklyTotal">0</div><div class="lbl">RFIs This Week</div></div></div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartRfiWeekly"></canvas></div></div>
  </div>

  <!-- ===== RFI Monthly ===== -->
  <div class="section active" id="sec-rfi-monthly">
    <div class="section-title">📆 RFI Submitted - Monthly</div>
    <div class="kpi-row"><div class="kpi blue"><div class="val" id="rfiMonthlyTotal">0</div><div class="lbl">RFIs This Month</div></div></div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartRfiMonthly"></canvas></div></div>
  </div>

  <!-- ===== RFI Type ===== -->
  <div class="section active" id="sec-rfi-type">
    <div class="section-title">🔧 Inspection Type Breakdown</div>
    <div class="chart-row"><div class="chart-card"><canvas id="chartRfiType"></canvas></div></div>
  </div>

  <!-- ===== RFI Subsystems ===== -->
  <div class="section active" id="sec-rfi-subsystems">
    <div class="section-title">🏗️ Top 10 Subsystems (RFI Count)</div>
    <div class="chart-row"><div class="chart-card" style="flex:2;"><canvas id="chartRfiSub"></canvas></div></div>
  </div>

  <!-- ===== RFI Recent Details ===== -->
  <div class="section active" id="sec-rfi-recent">
    <div class="section-title">🔍 Recent RFIs — Search by Asset Tag / RFI No (Last 100)</div>
    <div class="chart-card">
      <input id="rfiSearch" type="text" placeholder="Type Asset Tag or RFI No to filter..."
        style="width:100%;padding:10px 14px;border:1px solid #ddd;border-radius:8px;font-size:13px;margin-bottom:14px;">
      <div style="max-height:480px;overflow:auto;">
        <table id="rfiTable" style="width:100%;border-collapse:collapse;font-size:13px;">
          <thead>
            <tr style="background:var(--panel2);color:var(--teal);">
              <th style="padding:8px;text-align:left;">Asset Tag</th>
              <th style="padding:8px;text-align:left;">RFI No</th>
              <th style="padding:8px;text-align:left;">Discipline</th>
              <th style="padding:8px;text-align:left;">Status</th>
              <th style="padding:8px;text-align:left;">Inspection Date</th>
            </tr>
          </thead>
          <tbody id="rfiTableBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- ===== Punch Tracking & Closure ===== -->
  <div class="section active" id="sec-punch-tracking">
    <div class="section-title">📌 Punch Tracking &amp; Closure</div>
    <div class="kpi-row" id="punchTrackingKpis"></div>
    <div class="chart-row">
      <div class="chart-card" style="flex:2;">
        <h3>Open Punch Items — Top 15 Subsystems</h3>
        <canvas id="chartPunchBacklog"></canvas>
      </div>
    </div>
    <div class="chart-row">
      <div class="chart-card">
        <h3>Punch by Discipline (Open vs Closed)</h3>
        <canvas id="chartPunchClosure"></canvas>
      </div>
    </div>
  </div>

  <div class="footer">Auto-generated dashboard — CPP AGI / EACOP PS5 Project</div>
</div>

</div><!-- /tab-main-dashboard -->

<div id="tab-container"></div>



<script>
const ITR = __ITR_JSON__;
const PUNCH = __PUNCH_JSON__;
const RFI = __RFI_JSON__;
const EIT_TABLE = __EIT_TABLE_JSON__;
const EIT_DESC = __EIT_DESC_JSON__;
const EIT_PAGES = __EIT_PAGES_JSON__;

// Reset daily counts if today_label doesn't match real time
(function(){
  if(!ITR) return;
  const today = new Date();
  const y = today.getFullYear();
  const m = String(today.getMonth()+1).padStart(2,'0');
  const d = String(today.getDate()).padStart(2,'0');
  const realDate = y+'-'+m+'-'+d;
  if(ITR.today_label !== realDate){
    ITR.hourly = [];
    ITR.hourly_total = 0;
    ITR.hourly_submitted = 0;
    ITR.hourly_closed_eit = 0;
    // Update DOM
    document.querySelectorAll('.kpi.purple .val, .kpi .val').forEach(el => {
      const lbl = el.parentElement.querySelector('.lbl');
      if(lbl && lbl.textContent === 'Today Closed') el.textContent = '0';
      if(lbl && lbl.textContent === 'Today Submitted') el.textContent = '0';
    });
  }
})();
const CMT_QC_PUNCH = __CMT_QC_PUNCH_JSON__;
const CABLE_OV = __CABLE_OV_JSON__;
const CABLE_TRACKER = __CABLE_TRACKER_JSON__;
const SEARCH_INDEX = __SEARCH_INDEX_JSON__;
const SI       = SEARCH_INDEX.index   || SEARCH_INDEX;
const RFI_MAP  = SEARCH_INDEX.rfi_map || {};
const SUB_MAP  = SEARCH_INDEX.sub_map || {};
const PRI_MAP  = SEARCH_INDEX.pri_map || {};
const TID_MAP  = SEARCH_INDEX.tid_map || {};
const palette = ['#2563eb','#7c3aed','#c8940a','#1a8a4a','#c53030','#0891b2','#ec4899'];
const MS_STATUS_COLORS = {'Closed':'#1a8a4a','Submitted':'#2563eb','To be completed':'#c8940a','Other':'#9a8d7c'};
const MS_STATUS_ORDER = ['Closed','Submitted','To be completed','Other'];

// ---------- Global Chart.js + datalabels setup ----------
Chart.register(ChartDataLabels);
Chart.defaults.font.family = "'Inter', system-ui, sans-serif";
Chart.defaults.color = '#6b5e4d';
Chart.defaults.borderColor = '#d9d0c1';
Chart.defaults.plugins.datalabels.display = false; // off by default, enabled per-chart below

const DL_PIE = {  // for pie/doughnut: show value + %
  display:true, color:'#fff', font:{weight:'bold', size:11},
  formatter:(v,ctx)=>{
    const total = ctx.chart.getDatasetMeta(0).total || ctx.dataset.data.reduce((a,b)=>a+b,0);
    const pct = total ? Math.round(v/total*100) : 0;
    return v>0 ? `${v}\n(${pct}%)` : '';
  }
};
const DL_BAR = {  // for bar/line: show value above point/bar
  display:true, color:'#0b1120',
  backgroundColor:'rgba(255,255,255,.9)', borderRadius:4,
  padding:{top:2,bottom:2,left:4,right:4},
  anchor:'end', align:'top', offset:4,
  font:{weight:'bold', size:10},
  formatter:(v)=> v>0 ? v : ''
};
const DL_STACK = {  // for stacked bars: show value inside segment if big enough
  display:true, color:'#fff', font:{weight:'bold', size:10},
  formatter:(v)=> v>5 ? v : ''
};

// ---------- Sidebar toggle ----------
document.querySelectorAll('.sidebar input[type=checkbox]').forEach(cb=>{
  cb.addEventListener('change', ()=>{
    const el = document.getElementById(cb.dataset.target);
    if(cb.checked){ el.classList.add('active'); }
    else{ el.classList.remove('active'); }
  });
});

function multiChart(id, rows, type='bar'){
  const labels = rows.map(r=>r.label);
  const series = [
    {key:'E', name:'Electrical (E)', color:'#2563eb'},
    {key:'I', name:'Instrumentation (I)', color:'#7c3aed'},
    {key:'T', name:'Telecom (T)', color:'#0891b2'},
  ];
  new Chart(document.getElementById(id), {
    type:type,
    data:{ labels, datasets: series.map(s=>({
      label:s.name,
      data: rows.map(r=>r[s.key]||0),
      backgroundColor: type==='line' ? s.color+'33' : s.color,
      borderColor: s.color,
      borderWidth: 1,
      fill: type==='line',
      tension:.3,
      borderRadius: type==='bar'?6:0,
      datalabels: type==='bar' ? DL_BAR : {display:false},
    })) },
    options:{ plugins:{legend:{position:'bottom'}}, scales:{y:{beginAtZero:true}} }
  });
}

// ---------- Combined Overview (ITR vs Punch vs RFI) ----------
function combinedChart(id, itrRows, punchRows, rfiRows, type='line'){
  const toMap = rows => Object.fromEntries((rows||[]).map(r=>[r.label, r.Total||0]));
  const itrMap = toMap(itrRows), punchMap = toMap(punchRows), rfiMap = toMap(rfiRows);
  const labels = Array.from(new Set([...Object.keys(itrMap), ...Object.keys(punchMap), ...Object.keys(rfiMap)])).sort();

  const series = [
    {name:'ITR Closed', color:'#1a8a4a', map:itrMap},
    {name:'Punch Raised', color:'#c53030', map:punchMap},
    {name:'RFI Submitted', color:'#2563eb', map:rfiMap},
  ];
  new Chart(document.getElementById(id), {
    type:type,
    data:{ labels, datasets: series.map(s=>({
      label:s.name,
      data: labels.map(l=> s.map[l]||0),
      borderColor:s.color, backgroundColor: type==='line'? s.color+'33': s.color,
      fill: type==='line', tension:.3, borderRadius: type==='bar'?6:0,
      datalabels: DL_BAR,
    })) },
    options:{ plugins:{legend:{position:'bottom'}}, scales:{y:{beginAtZero:true}} }
  });
}

combinedChart('chartCombinedDaily', ITR.daily, (PUNCH&&PUNCH.daily)||[], (RFI&&RFI.daily)||[], 'line');
combinedChart('chartCombinedMonthly', ITR.monthly, (PUNCH&&PUNCH.monthly)||[], (RFI&&RFI.monthly)||[], 'bar');

// ---------- ITR Daily ----------
multiChart('chartDaily', ITR.daily, 'line');

// ---------- ITR Weekly ----------
multiChart('chartWeekly', ITR.weekly, 'bar');

// ---------- ITR Monthly ----------
multiChart('chartMonthly', ITR.monthly, 'bar');

// ---------- E&I&T ----------
(function(){
  const c = document.getElementById('eitCards');
  ITR.eit_summary.forEach(s=>{
    const div = document.createElement('div');
    div.className='kpi';
    div.innerHTML = `<div class="val">${s.closed} / ${s.total}</div>
      <div class="lbl">${s.label} — ${s.pct}%</div>
      <div class="progress-bar"><div class="progress-fill" style="width:${s.pct}%"></div></div>`;
    c.appendChild(div);
  });
  new Chart(document.getElementById('chartEIT'), {
    type:'bar',
    data:{ labels: ITR.eit_summary.map(s=>s.label),
      datasets:[
        {label:'Closed', data: ITR.eit_summary.map(s=>s.closed), backgroundColor:'#1a8a4a', datalabels: DL_BAR},
        {label:'Total', data: ITR.eit_summary.map(s=>s.total), backgroundColor:'#d9d0c1', datalabels: {display:false}}
      ]},
    options:{ plugins:{legend:{position:'bottom'}}, scales:{y:{beginAtZero:true}} }
  });
})();

// ---------- Progress by Milestone ----------
(function(){
  const c = document.getElementById('msCards');
  (ITR.milestone_summary||[]).forEach(s=>{
    const div = document.createElement('div');
    div.className='kpi';
    div.innerHTML = `<div class="val">${s.closed} / ${s.total}</div>
      <div class="lbl">🎯 ${s.label} — ${s.pct}%</div>
      <div class="progress-bar"><div class="progress-fill" style="width:${s.pct}%"></div></div>`;
    c.appendChild(div);
  });
  new Chart(document.getElementById('chartMilestone'), {
    type:'bar',
    data:{ labels: (ITR.milestone_summary||[]).map(s=>s.label),
      datasets: MS_STATUS_ORDER.map(st=>({
        label: st,
        data: (ITR.milestone_summary||[]).map(s=> (s.status&&s.status[st]) || 0),
        backgroundColor: MS_STATUS_COLORS[st],
        datalabels: DL_STACK
      }))
    },
    options:{ plugins:{legend:{position:'bottom'}}, scales:{ x:{stacked:true}, y:{stacked:true, beginAtZero:true} } }
  });
})();

// ---------- Today Closures by Milestone ----------
(function(){
  const tm = ITR.today_milestone || [];
  const total = ITR.today_milestone_total || 0;
  document.getElementById('todayMsTotal').innerText = total;
  const body = document.getElementById('todayMsBody');
  if (!body) return;
  const msMap = {};
  tm.forEach(r => {
    if (!msMap[r.milestone]) msMap[r.milestone] = {E:0, I:0, T:0};
    msMap[r.milestone][r.disc] = r.count;
  });
  let html = '';
  Object.keys(msMap).sort().forEach(ms => {
    const d = msMap[ms];
    const rowTotal = d.E + d.I + d.T;
    html += `<tr>
      <td style="text-align:left;border:1px solid var(--border);font-weight:bold;color:var(--red);background:var(--card3);">${ms}</td>
      <td style="text-align:center;border:1px solid var(--border);background:#FCE4D6;color:var(--red);">${d.E || '-'}</td>
      <td style="text-align:center;border:1px solid var(--border);background:#FFF2CC;color:var(--red);">${d.I || '-'}</td>
      <td style="text-align:center;border:1px solid var(--border);background:#D9D9D9;color:var(--red);">${d.T || '-'}</td>
      <td style="text-align:center;border:1px solid var(--border);font-weight:bold;color:var(--red);background:var(--card3);">${rowTotal}</td>
    </tr>`;
  });
  body.innerHTML = html;
})();

// ---------- Subsystem Summary ----------
(function(){
  const data = ITR.subsystem_summary || [];
  const body = document.getElementById('subsystemBody');
  if (!body || !data.length) return;
  const discBg = {'E':'#FCE4D6','I':'#FFF2CC','T':'#D9D9D9'};
  window._subsystemData = data;
  function render(filter) {
    let filtered = data;
    if (filter) {
      const q = filter.toLowerCase();
      filtered = data.filter(r => r.subsystem.toLowerCase().includes(q) || r.discipline.toLowerCase().includes(q));
    }
    const subs = [...new Set(filtered.map(r => r.subsystem))];
    let html = '';
    subs.forEach(sub => {
      const rows = filtered.filter(r => r.subsystem === sub);
      rows.forEach((r, i) => {
        const bg = discBg[r.disc] || 'transparent';
        const isFull = r.total > 0 && (r.pct >= 100 || r.closed === r.total);
        const rowBg = isFull ? 'linear-gradient(90deg,#7cd49b,#a9e6c2)' : '';
        const rowBorder = isFull ? '#4aa872' : 'var(--border)';
        const done = isFull ? '#0d3a20' : '';
        const pctColor = isFull ? '#0d3a20' : (r.pct >= 50 ? 'var(--green)' : 'var(--red)');
        html += `<tr style="background:${rowBg};">
          <td style="text-align:left;border:1px solid ${rowBorder};font-weight:bold;color:${done || 'var(--text)'};">${i === 0 ? sub : ''}</td>
          <td style="text-align:center;border:1px solid ${rowBorder};background:${isFull ? '#5fc287' : bg};color:${done || 'var(--red)'};">${r.discipline}</td>
          <td style="text-align:center;border:1px solid ${rowBorder};color:${done || 'var(--text)'};">${r.total}</td>
          <td style="text-align:center;border:1px solid ${rowBorder};color:${done || 'var(--green)'};font-weight:700;">${r.closed}</td>
          <td style="text-align:center;border:1px solid ${rowBorder};color:${done || 'var(--red)'};">${r.open}</td>
          <td style="text-align:center;border:1px solid ${rowBorder};color:${pctColor};font-weight:700;">${r.pct}%</td>
        </tr>`;
      });
    });
    body.innerHTML = html;
  }
  render('');
  document.getElementById('subsystemSearchInput').addEventListener('input', function(){ render(this.value); });
})();

function exportSubsystemToExcel(){
  const data = window._subsystemData || ITR.subsystem_summary || [];
  if(!data.length) return;
  let csv = 'Subsystem,Discipline,Total,Closed,Open,% Closed\n';
  data.forEach(r => {
    csv += `"${r.subsystem}","${r.discipline}",${r.total},${r.closed},${r.open},${r.pct}%\n`;
  });
  const blob = new Blob(['\uFEFF' + csv], {type:'text/csv;charset=utf-8;'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'Subsystem_Summary_CPP_AGI.csv';
  a.click();
}

// ---------- EIT ITR Description Table (Excel-style) ----------
(function(){
  if(!EIT_TABLE) return;
  document.getElementById('eitCutoffLbl').innerText = 'Cut-off: ' + EIT_TABLE.cutoff;
  document.getElementById('eitWeekColLbl').innerText = EIT_TABLE.week_label;

  const typeClass = {'E-Electrical':'type-E','I-Instrumentation':'type-I','T-Telecom':'type-T'};
  const tbody = document.getElementById('eitDescTableBody');
  let html = '';

  EIT_TABLE.order.forEach(type=>{
    const rowsOfType = EIT_TABLE.rows.filter(r=>r.type===type);
    rowsOfType.forEach((r,i)=>{
      const balCls = r.balance>0 ? 'balance-pos' : 'balance-zero';
      html += `<tr class="${typeClass[type]}">`;
      if(i===0){
        html += `<td class="type-cell" rowspan="${rowsOfType.length}">${type}</td>`;
      }
      html += `<td class="desc-cell">${r.desc}</td>
        <td>${r.total.toLocaleString()}</td>
        <td>${r.approved.toLocaleString()}</td>
        <td class="${balCls}">${r.balance.toLocaleString()}</td>
        <td>
          <div style="position:relative;background:#d9d9d9;border-radius:4px;height:18px;min-width:70px;">
            <div style="position:absolute;left:0;top:0;height:100%;width:${r.pct}%;background:linear-gradient(90deg,#70ad47,#a9d18e);border-radius:4px;"></div>
            <div style="position:relative;font-size:11px;font-weight:700;line-height:18px;">${r.pct}%</div>
          </div>
        </td>
        <td>${r.closed_wk.toLocaleString()}</td>
      </tr>`;
    });
    const t = EIT_TABLE.totals[type];
    html += `<tr class="eit-total-row">
      <td colspan="2">${type} — Total</td>
      <td>${t.total.toLocaleString()}</td>
      <td>${t.approved.toLocaleString()}</td>
      <td class="${t.balance>0?'balance-pos':''}">${t.balance.toLocaleString()}</td>
      <td>${t.pct}%</td>
      <td>${t.closed_wk.toLocaleString()}</td>
    </tr>`;
  });
  tbody.innerHTML = html;

  const g = EIT_TABLE.grand;
  document.getElementById('eitDescTableFoot').innerHTML = `<tr class="eit-total-row">
    <td colspan="2">GRAND TOTAL (E+I+T)</td>
    <td>${g.total.toLocaleString()}</td>
    <td>${g.approved.toLocaleString()}</td>
    <td class="${g.balance>0?'balance-pos':''}">${g.balance.toLocaleString()}</td>
    <td>${g.pct}%</td>
    <td>${g.closed_wk.toLocaleString()}</td>
  </tr>`;
})();

function exportEitTableToExcel(){
  if(!EIT_TABLE || typeof XLSX === 'undefined'){ alert('Excel library not loaded — check your internet connection.'); return; }
  const aoa = [
    ['EACOP PIPELINE PROJECT — ITR Description Table (E/I/T)'],
    ['Person in charge: EIT - Mohamed Abd Elgawad', '', '', 'Cut-off:', EIT_TABLE.cutoff, '', EIT_TABLE.week_label],
    [],
    ['Type','ITR Description','Total Task','Approved by EACOP','Balance','%Closed','Closed This Week']
  ];
  EIT_TABLE.order.forEach(type=>{
    EIT_TABLE.rows.filter(r=>r.type===type).forEach(r=>{
      aoa.push([type, r.desc, r.total, r.approved, r.balance, r.pct/100, r.closed_wk]);
    });
    const t = EIT_TABLE.totals[type];
    aoa.push([type+' - Total','', t.total, t.approved, t.balance, t.pct/100, t.closed_wk]);
  });
  const g = EIT_TABLE.grand;
  aoa.push(['GRAND TOTAL','', g.total, g.approved, g.balance, g.pct/100, g.closed_wk]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{wch:18},{wch:42},{wch:12},{wch:16},{wch:10},{wch:10},{wch:14}];
  ws['!merges'] = [{s:{r:0,c:0},e:{r:0,c:6}}];
  // % format for the %Closed column
  for(let R=4; R<aoa.length; R++){
    const cellRef = XLSX.utils.encode_cell({r:R,c:5});
    if(ws[cellRef]) ws[cellRef].z = '0.0%';
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'EIT ITRs');
  const today = EIT_TABLE.cutoff;
  XLSX.writeFile(wb, `EIT_ITR_Description_${today}.xlsx`);
}

// ---------- Cable CMT (from Pre_Com_Cable_ITR_Tracker) ----------
(function(){
  const TRACKER = CABLE_TRACKER;
  if(!TRACKER) return;
  const CABLE = TRACKER.by_disc;
  const SCOPE = TRACKER.by_scope;
  const bg = {'E-Electrical':'#FCE4D6','I-Instrumentation':'#FFF2CC','T-Telecom':'#D9D9D9'};
  // Table 1: By discipline
  document.getElementById('cableCmtOrigBody').innerHTML = CABLE.map(r => {
    const b = bg[r.disc] || 'transparent';
    return `<tr style="background:${b};">
      <td style="text-align:left;border:1px solid var(--border);font-weight:bold;">${r.disc}</td>
      <td style="text-align:center;border:1px solid var(--border);">${r.total}</td>
      <td style="text-align:center;border:1px solid var(--border);color:var(--green);font-weight:700;">${r.cmt_close}</td>
      <td style="text-align:center;border:1px solid var(--border);${r.cmt_open>0?'color:var(--red);font-weight:700':''}">${r.cmt_open}</td>
      <td style="text-align:center;border:1px solid var(--border);color:var(--green);font-weight:700;">${r.static_close}</td>
      <td style="text-align:center;border:1px solid var(--border);${r.static_open>0?'color:var(--red);font-weight:700':''}">${r.static_open}</td>
      <td style="text-align:center;border:1px solid #d4d4d4;">
        <div style="background:#d9d9d9;border-radius:4px;height:18px;min-width:60px;position:relative;">
          <div style="position:absolute;left:0;top:0;height:100%;width:${r.pct_cmt}%;background:linear-gradient(90deg,#70ad47,#a9d18e);border-radius:4px;"></div>
          <div style="position:relative;font-size:11px;font-weight:700;line-height:18px;">${r.pct_cmt}%</div>
        </div>
      </td>
      <td style="text-align:center;border:1px solid #d4d4d4;">
        <div style="background:#d9d9d9;border-radius:4px;height:18px;min-width:60px;position:relative;">
          <div style="position:absolute;left:0;top:0;height:100%;width:${r.pct_static}%;background:linear-gradient(90deg,#70ad47,#a9d18e);border-radius:4px;"></div>
          <div style="position:relative;font-size:11px;font-weight:700;line-height:18px;">${r.pct_static}%</div>
        </div>
      </td>
    </tr>`;
  }).join('');
  const t = CABLE.reduce((a,r) => ({total:a.total+r.total, cmt_close:a.cmt_close+r.cmt_close, cmt_open:a.cmt_open+r.cmt_open, static_close:a.static_close+r.static_close, static_open:a.static_open+r.static_open}), {total:0,cmt_close:0,cmt_open:0,static_close:0,static_open:0});
  t.pct_cmt = t.total ? Math.round(t.cmt_close/t.total*100) : 0;
  t.pct_static = t.total ? Math.round(t.static_close/t.total*100) : 0;
  document.getElementById('cableCmtOrigFoot').innerHTML = `<tr style="background:var(--card3);color:var(--text);font-weight:800;border-top:2px solid var(--gold);">
    <td style="border:1px solid var(--border);padding:7px;text-align:left;">TOTAL CABLE / CMT</td>
    <td style="border:1px solid var(--border);text-align:center;">${t.total}</td>
    <td style="border:1px solid var(--border);text-align:center;">${t.cmt_close}</td>
    <td style="border:1px solid var(--border);text-align:center;${t.cmt_open>0?'color:var(--red)':''}">${t.cmt_open}</td>
    <td style="border:1px solid var(--border);text-align:center;">${t.static_close}</td>
    <td style="border:1px solid var(--border);text-align:center;${t.static_open>0?'color:var(--red)':''}">${t.static_open}</td>
    <td style="border:1px solid var(--border);text-align:center;">${t.pct_cmt}%</td>
    <td style="border:1px solid var(--border);text-align:center;">${t.pct_static}%</td>
  </tr>`;
  window.__cableTrackerByDisc = CABLE;
  // Table 2: By scope — stacked E/I/T per scope
  if(!document.getElementById('cableScopeBody')) return;
  const scopeBg = {'E':'#FCE4D6','I':'#FFF2CC','T':'#D9D9D9'};
  let scopeHtml = '';
  const scopeNames = [...new Set(SCOPE.map(r => r.scope))];
  scopeNames.forEach(sc => {
    const rowsOf = SCOPE.filter(r => r.scope === sc);
    rowsOf.forEach((r, i) => {
      const bg = scopeBg[r.disc] || 'transparent';
      scopeHtml += `<tr>
        <td style="text-align:left;border:1px solid var(--border);font-weight:bold;">${i === 0 ? sc : ''}</td>
        <td style="text-align:center;border:1px solid var(--border);background:${bg};">${r.discipline}</td>
        <td style="text-align:center;border:1px solid var(--border);background:${bg};">${r.total}</td>
        <td style="text-align:center;border:1px solid var(--border);background:${bg};color:var(--green);font-weight:700;">${r.cmt_close}</td>
        <td style="text-align:center;border:1px solid var(--border);background:${bg};color:var(--green);font-weight:700;">${r.static_close}</td>
      </tr>`;
    });
    // Scope subtotal
    const st = rowsOf[0];
    scopeHtml += `<tr style="background:var(--card3);color:var(--text);font-weight:800;border-top:2px solid var(--gold);">
      <td style="border:1px solid var(--border);padding:7px;text-align:left;">${sc} — Total</td>
      <td style="border:1px solid var(--border);text-align:center;"></td>
      <td style="border:1px solid var(--border);text-align:center;">${st._scope_total}</td>
      <td style="border:1px solid var(--border);text-align:center;">${st._scope_cmt}</td>
      <td style="border:1px solid var(--border);text-align:center;">${st._scope_static}</td>
    </tr>`;
  });
  document.getElementById('cableScopeBody').innerHTML = scopeHtml;
  window.__cableScopeData = SCOPE;
})();

// ---------- CMT & QC Punch Summary ----------// ---------- CMT & QC Punch Summary ----------
(function(){
  const CMT_QC = CMT_QC_PUNCH;
  if(!CMT_QC || !CMT_QC.length) return;

  // Table 1: Summary
  const summaryBg = {'CMT Punch (A,B,C)':'#FCE4D6','QC Punch (A,B,C)':'#FFF2CC'};
  document.getElementById('cmtQcSummaryBody').innerHTML = CMT_QC[0].rows.map(r =>
    `<tr style="background:${summaryBg[r[0]]||'transparent'}">
      <td style="font-weight:bold;text-align:left;border:1px solid var(--border);">${r[0]}</td>
      <td style="text-align:center;border:1px solid var(--border);">${r[1]}</td>
      <td style="text-align:center;border:1px solid var(--border);color:var(--green);font-weight:700;">${r[2]}</td>
      <td style="text-align:center;border:1px solid var(--border);color:var(--red);font-weight:700;">${r[3]}</td>
      <td style="text-align:center;border:1px solid var(--border);">${r[4]}</td>
      <td style="text-align:center;border:1px solid var(--border);color:var(--green);font-weight:700;">${r[5]}</td>
      <td style="text-align:center;border:1px solid var(--border);color:var(--red);font-weight:700;">${r[6]}</td>
    </tr>`
  ).join('');

  // Totals row
  const t = CMT_QC[0].rows.reduce((a,r) => ({
    a1:a.a1+(+r[1]||0), a2:a.a2+(+r[2]||0), a3:a.a3+(+r[3]||0),
    a4:a.a4+(+r[4]||0), a5:a.a5+(+r[5]||0), a6:a.a6+(+r[6]||0)
  }), {a1:0,a2:0,a3:0,a4:0,a5:0,a6:0});
  document.getElementById('cmtQcSummaryBody').innerHTML +=
    `<tr style="background:var(--card3);color:var(--text);font-weight:800;border-top:2px solid var(--gold);">
      <td style="border:1px solid var(--border);padding:7px;text-align:left;">TOTAL</td>
      <td style="border:1px solid var(--border);text-align:center;">${t.a1}</td>
      <td style="border:1px solid var(--border);text-align:center;">${t.a2}</td>
      <td style="border:1px solid var(--border);text-align:center;color:var(--red);">${t.a3}</td>
      <td style="border:1px solid var(--border);text-align:center;">${t.a4}</td>
      <td style="border:1px solid var(--border);text-align:center;">${t.a5}</td>
      <td style="border:1px solid var(--border);text-align:center;color:var(--red);">${t.a6}</td>
    </tr>`;

  // Table 2: Per-Asset Detail
  const detailRows = CMT_QC[1].rows;
  window.__cmtQcDetailData = detailRows;

  function parseStatus(s){
    if(!s || s === 'No Data' || s === '') return [];
    return s.split('|').map(p => p.trim()).filter(p => p).map(p => {
      const parts = p.split('=');
      if(parts.length === 2){
        const cat = parts[0].trim();
        const st  = parts[1].trim();
        return {category: cat, status: st};
      }
      // Handle combined like "A=Open B=Open"
      const items = p.split(/ (?=[A-Z]=)/);
      return items.map(it => {
        const sp = it.split('=');
        return sp.length === 2 ? {category: sp[0].trim(), status: sp[1].trim()} : null;
      }).filter(Boolean);
    }).flat();
  }

  function renderCmtQcDetail(q){
    const query = (q||'').trim().toLowerCase();
    let rows = detailRows;
    if(query) rows = rows.filter(r =>
      r[0].toLowerCase().includes(query) || r[1].toLowerCase().includes(query) ||
      r[5].toLowerCase().includes(query) || r[7].toLowerCase().includes(query)
    );
    const statusColor = {'Open':'#c53030','Closed':'#1a8a4a','':'#9a8d7c'};
    const limitedRows = rows.slice(0, 800);
    document.getElementById('cmtQcDetailBody').innerHTML = limitedRows.map(r => {
      const cmtStatus = parseStatus(r[4]);
      const qcStatus  = parseStatus(r[6]);
      const cmtColors = {'Open':'#ff5252','Closed':'#69f0ae','':'#8a93ad'};
      const qcColors  = {'Open':'#d32f2f','Closed':'#2e7d32','':'#8a93ad'};
      const cmtB = cmtStatus.map(s =>
        `<span style="display:inline-block;padding:2px 8px;margin:1px;border-radius:2px;font-size:11px;font-weight:800;
          background:${cmtColors[s.status]||'#8a93ad'}22;color:${cmtColors[s.status]||'#8a93ad'};
          border:1px solid ${cmtColors[s.status]||'#8a93ad'}44;">${s.category}</span>`
      ).join('') || '<span style="color:var(--muted);font-size:11px;">-</span>';
      const qcB = qcStatus.map(s =>
        `<span style="display:inline-block;padding:2px 10px;margin:1px;border-radius:10px;font-size:11px;font-weight:800;
          background:${qcColors[s.status]||'#8a93ad'}22;color:${qcColors[s.status]||'#8a93ad'};
          border:1px solid ${qcColors[s.status]||'#8a93ad'}44;">${s.category}</span>`
      ).join('') || '<span style="color:var(--muted);font-size:11px;">-</span>';
      const hasOpen = (s) => s.some(x => x.status === 'Open');
      const allClosed = (s) => s.length > 0 && s.every(x => x.status === 'Closed');
      const cmtOpen = hasOpen(cmtStatus); const qcOpen = hasOpen(qcStatus);
      const cmtClosed = allClosed(cmtStatus); const qcClosed = allClosed(qcStatus);
      let rowBg = '', assetColor = 'var(--gold)';
      if(cmtOpen || qcOpen){ rowBg = 'rgba(197,48,48,.06)'; assetColor = '#c53030'; }
      else if((cmtClosed && cmtStatus.length) || (qcClosed && qcStatus.length)){ rowBg = 'rgba(26,138,74,.06)'; assetColor = '#1a8a4a'; }
      return `<tr style="border-bottom:1px solid var(--border);background:${rowBg};">
        <td style="padding:5px;font-weight:700;color:${assetColor};">${r[0]||'-'}</td>
        <td style="padding:5px;font-size:11px;color:var(--muted);">${r[1]||'-'}</td>
        <td style="padding:5px;text-align:center;font-size:11px;">${r[2]||'-'}</td>
        <td style="padding:5px;text-align:center;">${r[3]||'-'}</td>
        <td style="padding:5px;text-align:center;">${cmtB}</td>
        <td style="padding:5px;font-size:11px;color:var(--muted);max-width:280px;white-space:pre-wrap;">${(r[5]||'').slice(0,120)}${(r[5]||'').length>120?'...':''}</td>
        <td style="padding:5px;text-align:center;">${qcB}</td>
        <td style="padding:5px;font-size:11px;color:var(--muted);max-width:280px;">${(r[7]||'').slice(0,120)}${(r[7]||'').length>120?'...':''}</td>
      </tr>`;
    }).join('');
    document.getElementById('cmtQcDetailCount').innerText = rows.length > 800
      ? `Showing first 800 of ${rows.length} record(s) — use search to filter`
      : `Showing ${rows.length} record(s)`;
  }

  document.getElementById('cmtQcSearchInput').addEventListener('input', e => renderCmtQcDetail(e.target.value));
  renderCmtQcDetail('');
})();

function exportCmtQcDetailExcel(){
  const cmtqc = window.__cmtQcDetailData;
  if(!cmtqc || !cmtqc.length){ alert("No Per-Asset Detail data available"); return; }
  let html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
    '<head><meta charset="UTF-8">' +
    '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Sheet1</x:Name>' +
    '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
    '<style>td,th{font-family:Arial,sans-serif;font-size:12px;vertical-align:middle;}.hdr{background:#ED7D31;color:#000;font-weight:bold;text-align:center;border:1px solid #9c4a14;padding:8px;}.bg_e{background:#FCE4D6;}.bg_i{background:#FFF2CC;}.bg_t{background:#D9D9D9;}</style>' +
    '</head><body>' +
    '<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;min-width:800px;">' +
    '<tr><th colspan="8" class="hdr" style="font-size:14px;">EACOP PIPELINE PROJECT &mdash; CMT & QC Punch — Per-Asset Detail</th></tr>' +
    '<tr>' +
    '  <th class="hdr" style="text-align:left;">Asset</th>' +
    '  <th class="hdr" style="text-align:left;">Subsystem</th>' +
    '  <th class="hdr">Priority</th>' +
    '  <th class="hdr">Discipline</th>' +
    '  <th class="hdr">CMT (A,B,C)</th>' +
    '  <th class="hdr" style="text-align:left;">CMT Description</th>' +
    '  <th class="hdr">QC (A,B,C)</th>' +
    '  <th class="hdr" style="text-align:left;">QC Description</th>' +
    '</tr>';
  const bgMap = {'Electrical (E)':'bg_e','Instrumentation (I)':'bg_i','Telecom (T)':'bg_t'};
  cmtqc.forEach(r => {
    const cls = bgMap[r[3]] || '';
    html += '<tr class="' + cls + '">' +
      '<td style="border:1px solid #d4d4d4;font-weight:bold;">' + (r[0]||'') + '</td>' +
      '<td style="border:1px solid #d4d4d4;">' + (r[1]||'') + '</td>' +
      '<td style="border:1px solid #d4d4d4;text-align:center;">' + (r[2]||'') + '</td>' +
      '<td style="border:1px solid #d4d4d4;text-align:center;">' + (r[3]||'') + '</td>' +
      '<td style="border:1px solid #d4d4d4;text-align:center;">' + (r[4]||'') + '</td>' +
      '<td style="border:1px solid #d4d4d4;">' + (r[5]||'') + '</td>' +
      '<td style="border:1px solid #d4d4d4;text-align:center;">' + (r[6]||'') + '</td>' +
      '<td style="border:1px solid #d4d4d4;">' + (r[7]||'') + '</td>' +
    '</tr>';
  });
  html += '</table></body></html>';
  const blob = new Blob(['\ufeff' + html], {type:'application/vnd.ms-excel;charset=utf-8'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = 'CMT_QC_Per_Asset_Detail.xls';
  a.click();
  URL.revokeObjectURL(url);
}

function exportCableCmtExcelOrig(){
  const d = window.__cableTrackerByDisc;
  if(!d){ alert('No data'); return; }
  let html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
    '<head><meta charset="UTF-8">' +
    '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Sheet1</x:Name>' +
    '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
    '<style>td,th{font-family:Arial,sans-serif;font-size:12px;vertical-align:middle;}.hdr{background:#ED7D31;color:#000;font-weight:bold;text-align:center;border:1px solid #9c4a14;padding:8px;}</style>' +
    '</head><body>' +
    '<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;min-width:600px;">' +
    '<tr><th colspan="8" class="hdr" style="font-size:14px;">EACOP PIPELINE PROJECT &mdash; Cable / CMT Status by Discipline</th></tr>' +
    '<tr>' +
    '  <th class="hdr" style="text-align:left;">Discipline</th>' +
    '  <th class="hdr">Total</th>' +
    '  <th class="hdr">CMT Close</th>' +
    '  <th class="hdr">CMT Open</th>' +
    '  <th class="hdr">Static Close</th>' +
    '  <th class="hdr">Static Open</th>' +
    '  <th class="hdr">% CMT</th>' +
    '  <th class="hdr">% Static</th>' +
    '</tr>';
  const bgMap = {'E-Electrical':'#FCE4D6','I-Instrumentation':'#FFF2CC','T-Telecom':'#D9D9D9'};
  d.forEach(r => {
    const bg = bgMap[r.disc]||'#fff';
    html += '<tr style="background:' + bg + ';">' +
      '<td style="border:1px solid #d4d4d4;text-align:left;font-weight:bold;">' + r.disc + '</td>' +
      '<td style="border:1px solid #d4d4d4;text-align:center;">' + r.total + '</td>' +
      '<td style="border:1px solid #d4d4d4;text-align:center;">' + r.cmt_close + '</td>' +
      '<td style="border:1px solid #d4d4d4;text-align:center;">' + r.cmt_open + '</td>' +
      '<td style="border:1px solid #d4d4d4;text-align:center;">' + r.static_close + '</td>' +
      '<td style="border:1px solid #d4d4d4;text-align:center;">' + r.static_open + '</td>' +
      '<td style="border:1px solid #d4d4d4;text-align:center;">' + r.pct_cmt + '%</td>' +
      '<td style="border:1px solid #d4d4d4;text-align:center;">' + r.pct_static + '%</td>' +
    '</tr>';
  });
  const t = d.reduce((a,r) => ({total:a.total+r.total,cmt_close:a.cmt_close+r.cmt_close,cmt_open:a.cmt_open+r.cmt_open,static_close:a.static_close+r.static_close,static_open:a.static_open+r.static_open}), {total:0,cmt_close:0,cmt_open:0,static_close:0,static_open:0});
  t.pct_cmt = t.total ? Math.round(t.cmt_close/t.total*100) : 0;
  t.pct_static = t.total ? Math.round(t.static_close/t.total*100) : 0;
  html += '<tr style="background:#404040;color:#fff;font-weight:800;">' +
    '<td style="border:1px solid #222;padding:7px;text-align:left;">TOTAL CABLE / CMT</td>' +
    '<td style="border:1px solid #222;text-align:center;">' + t.total + '</td>' +
    '<td style="border:1px solid #222;text-align:center;">' + t.cmt_close + '</td>' +
    '<td style="border:1px solid #222;text-align:center;">' + t.cmt_open + '</td>' +
    '<td style="border:1px solid #222;text-align:center;">' + t.static_close + '</td>' +
    '<td style="border:1px solid #222;text-align:center;">' + t.static_open + '</td>' +
    '<td style="border:1px solid #222;text-align:center;">' + t.pct_cmt + '%</td>' +
    '<td style="border:1px solid #222;text-align:center;">' + t.pct_static + '%</td>' +
  '</tr>' +
  '</table></body></html>';
  downloadExcel(html, 'Cable_CMT_Discipline');
}

function exportCableScopeExcel(){
  const d = window.__cableScopeData;
  if(!d){ alert('No data'); return; }
  const scopeBg = {'E':'#FCE4D6','I':'#FFF2CC','T':'#D9D9D9'};
  let html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
    '<head><meta charset="UTF-8">' +
    '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Sheet1</x:Name>' +
    '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
    '<style>td,th{font-family:Arial,sans-serif;font-size:12px;vertical-align:middle;}.hdr{background:#ED7D31;color:#000;font-weight:bold;text-align:center;border:1px solid #9c4a14;padding:8px;}</style>' +
    '</head><body>' +
    '<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;min-width:600px;">' +
    '<tr><th colspan="5" class="hdr" style="font-size:14px;">EACOP PIPELINE PROJECT &mdash; CMT & Static Test Summary by Scope (E / I / T)</th></tr>' +
    '<tr>' +
    '  <th class="hdr" style="text-align:left;">Scope</th>' +
    '  <th class="hdr">Discipline</th>' +
    '  <th class="hdr">Total</th>' +
    '  <th class="hdr">CMT Close</th>' +
    '  <th class="hdr">Static Close</th>' +
    '</tr>';
  const scopeNames = [...new Set(d.map(r => r.scope))];
  scopeNames.forEach(sc => {
    const rowsOf = d.filter(r => r.scope === sc);
    rowsOf.forEach(r => {
      const bg = scopeBg[r.disc] || '#fff';
      html += '<tr style="background:' + bg + ';">' +
        '<td style="border:1px solid #d4d4d4;text-align:left;font-weight:bold;">' + r.scope + '</td>' +
        '<td style="border:1px solid #d4d4d4;text-align:center;">' + r.discipline + '</td>' +
        '<td style="border:1px solid #d4d4d4;text-align:center;">' + r.total + '</td>' +
        '<td style="border:1px solid #d4d4d4;text-align:center;color:#1d6f42;font-weight:bold;">' + r.cmt_close + '</td>' +
        '<td style="border:1px solid #d4d4d4;text-align:center;color:#1d6f42;font-weight:bold;">' + r.static_close + '</td>' +
      '</tr>';
    });
    const st = rowsOf[0];
    html += '<tr style="background:#404040;color:#fff;font-weight:bold;">' +
      '<td style="border:1px solid #222;text-align:left;padding:7px;">' + sc + ' — Total</td>' +
      '<td style="border:1px solid #222;text-align:center;"></td>' +
      '<td style="border:1px solid #222;text-align:center;">' + st._scope_total + '</td>' +
      '<td style="border:1px solid #222;text-align:center;">' + st._scope_cmt + '</td>' +
      '<td style="border:1px solid #222;text-align:center;">' + st._scope_static + '</td>' +
    '</tr>';
  });
  html += '</table></body></html>';
  downloadExcel(html, 'CMT_Static_Scope_Summary');
}


function exportEitClosureExcel(){
  const d = window.__eitClosureData;
  if(!d){ alert('No data'); return; }
  if(typeof XLSX === 'undefined'){ alert('Excel library not loaded'); return; }
  const aoa = [
    ['EIT — Closed Status & Backlog'],
    ['Discipline','ITR Description','Total Tasks','Closed','Balance','% Closed','Closed This Week'],
  ];
  d.rows.forEach(r => {
    aoa.push([r.type, r.desc, r.total, r.approved, r.balance, r.pct/100, r.closed_wk]);
  });
  d.order.forEach(type => {
    const rowsOfType = d.rows.filter(r => r.type === type);
    const t = rowsOfType.reduce((a,r) => ({total:a.total+r.total, approved:a.approved+r.approved, balance:a.balance+r.balance, closed_wk:a.closed_wk+r.closed_wk}), {total:0,approved:0,balance:0,closed_wk:0});
    aoa.push([type+' TOTAL','', t.total, t.approved, t.balance, t.total ? t.approved/t.total : 0, t.closed_wk]);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{wch:18},{wch:42},{wch:12},{wch:10},{wch:10},{wch:10},{wch:18}];
  ws['!merges'] = [{s:{r:0,c:0},e:{r:0,c:6}}];
  for(let R=2; R<aoa.length; R++){
    const cellRef = XLSX.utils.encode_cell({r:R,c:5});
    if(ws[cellRef]) ws[cellRef].z = '0.0%';
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'EIT Closure');
  const today = EIT_TABLE ? EIT_TABLE.cutoff : new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `EIT_Closure_Status_${today}.xlsx`);
}

// ---------- Excel export helpers ----------
function buildExcelHeader(title){
  return '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
  '<head><meta charset="UTF-8">' +
  '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Sheet1</x:Name>' +
  '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
  '<style>' +
  '  td,th{font-family:Arial,sans-serif;font-size:12px;vertical-align:middle;}' +
  '  .hdr{background:#ED7D31;color:#000;font-weight:bold;text-align:center;border:1px solid #9c4a14;padding:8px;}' +
  '  .tot{background:#404040;color:#fff;font-weight:bold;}' +
  '</style>' +
  '</head><body>' +
  '<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;min-width:600px;">' +
  '<tr><th colspan="5" class="hdr" style="font-size:14px;">EACOP PIPELINE PROJECT &mdash; ' + title + '</th></tr>' +
  '<tr>' +
  '  <th class="hdr" style="text-align:left;">Discipline</th>' +
  '  <th class="hdr">Total</th>' +
  '  <th class="hdr">Open / Balance</th>' +
  '  <th class="hdr">Closed / Approved</th>' +
  '  <th class="hdr">% Closed</th>' +
  '</tr>';
}
function downloadExcel(html, filename){
  const blob = new Blob(['\ufeff' + html], {type:'application/vnd.ms-excel;charset=utf-8'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = filename + '.xls';
  a.click();
  URL.revokeObjectURL(url);
}

// ---------- 1. Punch Closing Status Table ----------
(function(){
  if(!PUNCH) return;
  const discs = ['Electrical (E)', 'Instrumentation (I)', 'Telecom (T)'];
  const rows = discs.map(d => ({
    disc: d,
    total: PUNCH.disc_counts[d] || 0,
    open: PUNCH.disc_counts_open[d] || 0,
    closed: PUNCH.disc_counts_closed[d] || 0,
  }));
  rows.forEach(r => r.pct = r.total ? Math.round(r.closed/r.total*100) : 0);

  const discColors = {'Electrical (E)':'#FCE4D6','Instrumentation (I)':'#FFF2CC','Telecom (T)':'#D9D9D9'};
  document.getElementById('punchStatusBody').innerHTML = rows.map(r => `
    <tr style="background:${discColors[r.disc]||'transparent'}">
      <td style="font-weight:bold;text-align:center;border:1px solid var(--border);">${r.disc}</td>
      <td style="text-align:center;border:1px solid var(--border);">${r.total}</td>
      <td style="text-align:center;border:1px solid var(--border);color:var(--red);font-weight:700;">${r.open}</td>
      <td style="text-align:center;border:1px solid var(--border);color:var(--green);font-weight:700;">${r.closed}</td>
      <td style="text-align:center;border:1px solid var(--border);">
        <div style="background:#d9d9d9;border-radius:4px;height:18px;min-width:60px;position:relative;">
          <div style="position:absolute;left:0;top:0;height:100%;width:${r.pct}%;background:linear-gradient(90deg,#70ad47,#a9d18e);border-radius:4px;"></div>
          <div style="position:relative;font-size:11px;font-weight:700;line-height:18px;">${r.pct}%</div>
        </div>
      </td>
    </tr>`).join('');

  const t = rows.reduce((a,r) => ({total:a.total+r.total, open:a.open+r.open, closed:a.closed+r.closed}), {total:0,open:0,closed:0});
  t.pct = t.total ? Math.round(t.closed/t.total*100) : 0;
  document.getElementById('punchStatusFoot').innerHTML = `<tr style="background:var(--card3);color:var(--text);font-weight:800;border-top:2px solid var(--gold);">
    <td style="border:1px solid var(--border);padding:7px;text-align:center;">GRAND TOTAL</td>
    <td style="border:1px solid var(--border);text-align:center;">${t.total}</td>
    <td style="border:1px solid var(--border);text-align:center;color:var(--red);">${t.open}</td>
    <td style="border:1px solid var(--border);text-align:center;">${t.closed}</td>
    <td style="border:1px solid var(--border);text-align:center;">${t.pct}%</td>
  </tr>`;

  window.__punchStatusData = {rows, total: t};
})();

function exportPunchStatusExcel(){
  const d = window.__punchStatusData;
  if(!d){ alert('No data'); return; }
  const discColors = {'Electrical (E)':'#FCE4D6','Instrumentation (I)':'#FFF2CC','Telecom (T)':'#D9D9D9'};
  let html = buildExcelHeader('Punch Closing Status &mdash; By Discipline');
  d.rows.forEach(r => {
    const bg = discColors[r.disc]||'#fff';
    html += `<tr style="background:${bg};">
      <td style="border:1px solid #bbb;font-weight:bold;text-align:center;">${r.disc}</td>
      <td style="border:1px solid #bbb;text-align:center;">${r.total}</td>
      <td style="border:1px solid #bbb;text-align:center;color:#C00000;font-weight:bold;">${r.open}</td>
      <td style="border:1px solid #bbb;text-align:center;color:#1d6f42;font-weight:bold;">${r.closed}</td>
      <td style="border:1px solid #bbb;text-align:center;">${r.pct}%</td>
    </tr>`;
  });
  html += `<tr style="background:#404040;color:#fff;font-weight:bold;">
    <td style="border:1px solid #222;text-align:center;padding:7px;">GRAND TOTAL</td>
    <td style="border:1px solid #222;text-align:center;">${d.total.total}</td>
    <td style="border:1px solid #222;text-align:center;color:#ff9b9b;">${d.total.open}</td>
    <td style="border:1px solid #222;text-align:center;">${d.total.closed}</td>
    <td style="border:1px solid #222;text-align:center;">${d.total.pct}%</td>
  </tr>`;
  downloadExcel(html, 'Punch_Closing_Status');
}

// ---------- 3,4,5. E / I / T Description Status Tables ----------
(function(){
  if(!EIT_DESC) return;
  function pct(a,b){ return b ? ((a/b*100).toFixed(1)) : 0; }
  const config = [
    {bodyId:'elecDescBody', headBg:'#2563eb', rowBg:'#f0f5fb', label:'E - Electrical', border:'#c8d6e5'},
    {bodyId:'instDescBody', headBg:'#7c3aed', rowBg:'#f2fbe8', label:'I - Instrumentation', border:'#c5e0b3'},
    {bodyId:'teleDescBody', headBg:'#0891b2', rowBg:'#f0f0f0', label:'T - Telecom', border:'#cccccc'},
  ];

  const colHeaders = `
    <th style="background:${'#2c2416'};color:#f5f0e8;padding:10px 8px;font-size:12px;font-weight:700;text-align:left;border:1px solid #6b5e4d;letter-spacing:0.5px;">Description</th>
    <th style="background:${'#2c2416'};color:#f5f0e8;padding:10px 8px;font-size:12px;font-weight:700;text-align:center;border:1px solid #6b5e4d;letter-spacing:0.5px;">Total Tasks</th>
    <th style="background:${'#2c2416'};color:#f5f0e8;padding:10px 8px;font-size:12px;font-weight:700;text-align:center;border:1px solid #6b5e4d;letter-spacing:0.5px;">Closed</th>
    <th style="background:${'#2c2416'};color:#f5f0e8;padding:10px 8px;font-size:12px;font-weight:700;text-align:center;border:1px solid #6b5e4d;letter-spacing:0.5px;">Pending</th>
    <th style="background:${'#2c2416'};color:#f5f0e8;padding:10px 8px;font-size:12px;font-weight:700;text-align:center;border:1px solid #6b5e4d;letter-spacing:0.5px;">% Completion</th>
    <th style="background:${'#2c2416'};color:#f5f0e8;padding:10px 8px;font-size:12px;font-weight:700;text-align:center;border:1px solid #6b5e4d;letter-spacing:0.5px;">Open A-Punch</th>`;

  config.forEach(cfg => {
    const rows = EIT_DESC[cfg.bodyId === 'elecDescBody' ? 'E' : cfg.bodyId === 'instDescBody' ? 'I' : 'T'] || [];
    if(!rows.length) return;
    let total = 0, closed = 0;
    let html = `<tr>${colHeaders}</tr>`;
    rows.forEach((r, idx) => {
      total += r.total; closed += r.closed;
      const pending = r.total - r.closed;
      const p = pct(r.closed, r.total);
      const altBg = idx % 2 === 0 ? cfg.rowBg : '#ffffff';
      html += `<tr style="background:${altBg};transition:background 0.2s;" onmouseover="this.style.background='${cfg.headBg}22'" onmouseout="this.style.background='${altBg}'">
        <td style="border:1px solid ${cfg.border};padding:8px 6px;font-weight:600;text-align:left;font-size:12px;color:#2c2416;">${r.desc}</td>
        <td style="border:1px solid ${cfg.border};text-align:center;padding:8px 6px;font-size:12px;font-weight:600;">${r.total}</td>
        <td style="border:1px solid ${cfg.border};text-align:center;padding:8px 6px;font-size:12px;"><span style="display:inline-block;padding:2px 10px;border-radius:8px;background:rgba(29,111,66,0.12);color:#1d6f42;font-weight:700;">${r.closed}</span></td>
        <td style="border:1px solid ${cfg.border};text-align:center;padding:8px 6px;font-size:12px;">${pending > 0 ? `<span style="display:inline-block;padding:2px 10px;border-radius:8px;background:rgba(192,0,0,0.1);color:#C00000;font-weight:700;">${pending}</span>` : '<span style="color:#1d6f42;">0</span>'}</td>
        <td style="border:1px solid ${cfg.border};text-align:center;padding:8px 6px;font-size:12px;">
          <div style="display:flex;align-items:center;gap:6px;justify-content:center;">
            <div style="flex:1;max-width:70px;height:6px;background:#e0e0e0;border-radius:3px;overflow:hidden;">
              <div style="height:100%;width:${p}%;background:${p >= 80 ? '#1d6f42' : p >= 50 ? '#f39c12' : '#C00000'};border-radius:3px;transition:width 0.5s;"></div>
            </div>
            <span style="font-weight:700;font-size:11px;color:${p >= 80 ? '#1d6f42' : p >= 50 ? '#f39c12' : '#C00000'};">${p}%</span>
          </div>
        </td>
        <td style="border:1px solid ${cfg.border};text-align:center;padding:8px 6px;font-size:12px;color:var(--muted);">-</td>
      </tr>`;
    });
    const tPct = pct(closed, total);
    html += `<tr style="background:#2c2416;color:#fff;font-weight:800;">
      <td style="border:1px solid #6b5e4d;padding:10px 8px;text-align:left;font-size:13px;">TOTAL ${cfg.label}</td>
      <td style="border:1px solid #6b5e4d;text-align:center;padding:10px 8px;font-size:13px;">${total}</td>
      <td style="border:1px solid #6b5e4d;text-align:center;padding:10px 8px;font-size:13px;color:#22c55e;">${closed}</td>
      <td style="border:1px solid #6b5e4d;text-align:center;padding:10px 8px;font-size:13px;${total-closed > 0 ? 'color:#f87171;' : ''}">${total - closed}</td>
      <td style="border:1px solid #6b5e4d;text-align:center;padding:10px 8px;font-size:13px;">
        <div style="display:flex;align-items:center;gap:6px;justify-content:center;">
          <div style="flex:1;max-width:70px;height:8px;background:#6b5e4d;border-radius:4px;overflow:hidden;">
            <div style="height:100%;width:${tPct}%;background:${tPct >= 80 ? '#22c55e' : tPct >= 50 ? '#f59e0b' : '#f87171'};border-radius:4px;"></div>
          </div>
          <span style="font-weight:700;font-size:12px;color:#fff;">${tPct}%</span>
        </div>
      </td>
      <td style="border:1px solid #6b5e4d;text-align:center;padding:10px 8px;font-size:13px;">-</td>
    </tr>`;
    document.getElementById(cfg.bodyId).innerHTML = html;
    if(!window.__eitDescStore) window.__eitDescStore = {};
    const codeLetter = cfg.bodyId === 'elecDescBody' ? 'E' : cfg.bodyId === 'instDescBody' ? 'I' : 'T';
    window.__eitDescStore[codeLetter] = {rows, total: {total, closed, balance: total-closed, pct: total ? Math.round(closed/total*100) : 0}, label: cfg.label, color: cfg.rowBg};
  });
})();

function exportDescExcel(code){
  if(!EIT_DESC) { alert('No data'); return; }
  const rows = EIT_DESC[code] || [];
  if(!rows.length) { alert('No data for ' + code); return; }
  const labels = {'E':'E - Electrical','I':'I - Instrumentation','T':'T - Telecom'};
  const colors = {'E':'#FCE4D6','I':'#E2EFDA','T':'#D6DCE4'};
  const bg = colors[code] || '#fff';
  const label = labels[code] || code;
  let html = buildExcelHeader(label + ' - Status & Backlog');
  // Update title
  html = html.replace(`<tr><th colspan="5" class="hdr" style="font-size:14px;">EACOP PIPELINE PROJECT &mdash; ${label} - Status & Backlog</th></tr>`,
    `<tr><th colspan="5" class="hdr" style="font-size:14px;">EACOP PIPELINE PROJECT &mdash; ${label}</th></tr>`);
  rows.forEach(r => {
    html += `<tr style="background:${bg};">
      <td style="border:1px solid #bbb;text-align:left;font-weight:bold;">${r.desc}</td>
      <td style="border:1px solid #bbb;text-align:center;">${r.total}</td>
      <td style="border:1px solid #bbb;text-align:center;color:#1d6f42;font-weight:bold;">${r.closed}</td>
      <td style="border:1px solid #bbb;text-align:center;${r.balance>0?'color:#C00000;font-weight:bold':''}">${r.balance}</td>
      <td style="border:1px solid #bbb;text-align:center;">${r.pct}%</td>
    </tr>`;
  });
  const t = rows.reduce((a,r) => ({total:a.total+r.total, closed:a.closed+r.closed, balance:a.balance+r.balance}), {total:0,closed:0,balance:0});
  html += `<tr style="background:#404040;color:#fff;font-weight:bold;">
    <td style="border:1px solid #222;text-align:left;padding:7px;">TOTAL ${label}</td>
    <td style="border:1px solid #222;text-align:center;">${t.total}</td>
    <td style="border:1px solid #222;text-align:center;">${t.closed}</td>
    <td style="border:1px solid #222;text-align:center;${t.balance>0?'color:#ff9b9b':''}">${t.balance}</td>
    <td style="border:1px solid #222;text-align:center;">${t.pct}%</td>
  </tr>`;
  downloadExcel(html, code + '_Description_Status');
}

// ---------- Punch List ----------
if(PUNCH){
  // status donut
  const statusLabels = Object.keys(PUNCH.status_counts);
  new Chart(document.getElementById('chartPunchStatus'), {
    type:'doughnut',
    data:{ labels: statusLabels, datasets:[{ data: statusLabels.map(k=>PUNCH.status_counts[k]),
      backgroundColor: palette, datalabels: DL_PIE }] },
    options:{ plugins:{legend:{position:'bottom'}} }
  });

  multiChart('chartPunchDaily', PUNCH.daily, 'line');
  multiChart('chartPunchWeekly', PUNCH.weekly, 'bar');
  multiChart('chartPunchMonthly', PUNCH.monthly, 'bar');

  const catLabels = Object.keys(PUNCH.cat_counts);
  new Chart(document.getElementById('chartPunchCat'), {
    type:'pie',
    data:{ labels: catLabels, datasets:[{ data: catLabels.map(k=>PUNCH.cat_counts[k]), backgroundColor:palette, datalabels: DL_PIE }] },
    options:{ plugins:{legend:{position:'bottom'}} }
  });

  const discLabels = Object.keys(PUNCH.disc_counts);
  new Chart(document.getElementById('chartPunchDisc'), {
    type:'pie',
    data:{ labels: discLabels, datasets:[{ data: discLabels.map(k=>PUNCH.disc_counts[k]), backgroundColor:palette, datalabels: DL_PIE }] },
    options:{ plugins:{legend:{position:'bottom'}} }
  });

  new Chart(document.getElementById('chartPunchSub'), {
    type:'bar',
    data:{ labels: PUNCH.top_subsystems.map(r=>r.label),
      datasets:[{ label:'Punch Items', data: PUNCH.top_subsystems.map(r=>r.count), backgroundColor:'#7c3aed', borderRadius:6, datalabels: {...DL_BAR, anchor:'end', align:'right'} }]},
    options:{ indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{beginAtZero:true}} }
  });
  // ---- Recent Punches Table (sec-punch-recent) ----
  (function(){
    const catColors = {'A':'#c53030','B':'#c8940a','C':'#2563eb','':' #9a8d7c'};
    const statusColor = {'Open':'#c53030','Closed':'#1a8a4a','Title':'#9a8d7c'};
    let activeCat = 'All';

    const allCats = ['All', ...Object.keys(PUNCH.cat_counts).sort()];
    const btnContainer = document.getElementById('punchCatBtns');
    allCats.forEach(cat=>{
      const btn = document.createElement('button');
      btn.textContent = cat === 'All' ? '📋 All' : `Cat ${cat}`;
      btn.dataset.cat = cat;
      const col = cat==='All' ? '#1a8a4a' : (catColors[cat]||'#9a8d7c');
      btn.style.cssText = `padding:6px 14px;border-radius:20px;border:2px solid ${col};
        background:${cat==='All'?col+'33':'transparent'};color:${col};font-weight:700;
        font-size:12px;cursor:pointer;transition:.15s;`;
      btn.addEventListener('click',()=>{
        activeCat = cat;
        document.querySelectorAll('#punchCatBtns button').forEach(b=>{
          const bc = b.dataset.cat==='All' ? '#1a8a4a' : (catColors[b.dataset.cat]||'#9a8d7c');
          b.style.background = b.dataset.cat===cat ? bc+'33' : 'transparent';
        });
        renderPunchTable();
      });
      btnContainer.appendChild(btn);
    });

    function renderPunchTable(){
      const q = (document.getElementById('punchSearchInput').value||'').trim().toLowerCase();
      let rows = PUNCH.recent;
      if(activeCat !== 'All') rows = rows.filter(r=>(r.cat||'').toString().trim()===activeCat);
      if(q) rows = rows.filter(r=>
        (r.plid||'').toLowerCase().includes(q) ||
        (r.subsystem||'').toLowerCase().includes(q) ||
        (r.desc||'').toLowerCase().includes(q) ||
        (r.rfi_no||'').toLowerCase().includes(q)
      );
      const catBg = {'A':'rgba(197,48,48,.08)','B':'rgba(200,148,10,.08)','C':'rgba(37,99,235,.08)'};
      document.getElementById('punchRecentBody').innerHTML = rows.map(r=>{
        const cat = (r.cat||'').toString().trim();
        const sc = statusColor[r.status]||'#9a8d7c';
        const cc = catColors[cat]||'#9a8d7c';
        const bg = catBg[cat]||'';
        return `<tr style="border-bottom:1px solid var(--border);background:${bg};">
          <td style="padding:7px;color:var(--gold);font-weight:700;">${r.plid||'-'}</td>
          <td style="padding:7px;text-align:center;">
            <span style="display:inline-block;padding:2px 10px;border-radius:12px;font-size:12px;
              font-weight:800;background:${cc}22;color:${cc};border:1px solid ${cc}44;">
              ${cat||'-'}
            </span>
          </td>
          <td style="padding:7px;">${r.disc||'-'}</td>
          <td style="padding:7px;text-align:center;">
            <span style="display:inline-block;padding:2px 9px;border-radius:10px;font-size:11px;
              font-weight:700;background:${sc}22;color:${sc};border:1px solid ${sc}44;">
              ${r.status||'-'}
            </span>
          </td>
          <td style="padding:7px;color:var(--blue);">${r.rfi_no||'-'}</td>
          <td style="padding:7px;color:var(--muted);font-size:11px;">${r.subsystem||'-'}</td>
          <td style="padding:7px;">${r.date||'-'}</td>
          <td style="padding:7px;color:var(--muted);font-size:12px;">${r.desc||'-'}</td>
        </tr>`;
      }).join('');
      document.getElementById('punchRecentCount').innerText =
        `Showing ${rows.length} record(s) — Category filter: ${activeCat}`;
    }

    document.getElementById('punchSearchInput').addEventListener('input', renderPunchTable);
    renderPunchTable();
  })();

} else {
  ['sec-punch-kpi','sec-punch-daily','sec-punch-weekly','sec-punch-monthly','sec-punch-breakdown','sec-punch-subsystems','sec-punch-recent']
    .forEach(id=>{
      document.getElementById(id).innerHTML = '<div class="section-title">⚠️ Punch List Register file not found in Downloads</div>';
    });
}

// ---------- RFI / Inspection Register ----------
if(RFI){
  // KPI cards
  (function(){
    const c = document.getElementById('rfiKpiCards');
    const kpis = [
      {icon:'📋', val:RFI.total_assets, lbl:'Total Assets in Register', cls:''},
      {icon:'📝', val:RFI.total_rfi, lbl:'Total RFIs Submitted', cls:'gold'},
      {icon:'✅', val:(RFI.status_counts['Accepted']||0)+(RFI.status_counts['Accepted with Punch']||0), lbl:'Accepted (incl. with Punch)', cls:'teal'},
      {icon:'🟡', val:RFI.status_counts['Open']||0, lbl:'Open RFIs', cls:'pink'},
      {icon:'⛔', val:RFI.status_counts['Hold by EACOP']||0, lbl:'Hold by EACOP', cls:'blue'},
    ];
    kpis.forEach(k=>{
      const div = document.createElement('div');
      div.className = 'kpi ' + k.cls;
      div.innerHTML = `<div class="icon">${k.icon}</div><div class="val">${k.val}</div><div class="lbl">${k.lbl}</div>`;
      c.appendChild(div);
    });
  })();

  document.getElementById('rfiDailyTotal').innerText = RFI.daily_total;
  document.getElementById('rfiWeeklyTotal').innerText = RFI.weekly_total;
  document.getElementById('rfiMonthlyTotal').innerText = RFI.monthly_total;

  // overall status donut
  const statusLabels = Object.keys(RFI.status_counts);
  new Chart(document.getElementById('chartRfiStatus'), {
    type:'doughnut',
    data:{ labels: statusLabels, datasets:[{ data: statusLabels.map(k=>RFI.status_counts[k]), backgroundColor: palette, datalabels: DL_PIE }] },
    options:{ plugins:{legend:{position:'bottom'}} }
  });

  // status by discipline (stacked bar)
  new Chart(document.getElementById('chartRfiStatusDisc'), {
    type:'bar',
    data:{
      labels: RFI.status_by_disc.map(r=>r.label),
      datasets: statusLabels.map((s,i)=>({
        label:s, data: RFI.status_by_disc.map(r=>r[s]||0), backgroundColor: palette[i % palette.length], datalabels: DL_STACK
      }))
    },
    options:{ plugins:{legend:{position:'bottom'}}, scales:{ x:{stacked:true}, y:{stacked:true, beginAtZero:true} } }
  });

  // daily/weekly/monthly trends (E/I/T + Total)
  multiChart('chartRfiDaily', RFI.daily, 'line');
  multiChart('chartRfiWeekly', RFI.weekly, 'bar');
  multiChart('chartRfiMonthly', RFI.monthly, 'bar');

  // inspection type
  const typeLabels = Object.keys(RFI.type_counts);
  new Chart(document.getElementById('chartRfiType'), {
    type:'pie',
    data:{ labels: typeLabels, datasets:[{ data: typeLabels.map(k=>RFI.type_counts[k]), backgroundColor: palette, datalabels: DL_PIE }] },
    options:{ plugins:{legend:{position:'bottom'}} }
  });

  // top subsystems
  new Chart(document.getElementById('chartRfiSub'), {
    type:'bar',
    data:{ labels: RFI.top_subsystems.map(r=>r.label),
      datasets:[{ label:'RFI Count', data: RFI.top_subsystems.map(r=>r.count), backgroundColor:'#2563eb', borderRadius:6, datalabels: {...DL_BAR, anchor:'end', align:'right'} }]},
    options:{ indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{beginAtZero:true}} }
  });

  // ---- Recent RFI table with search ----
  function renderRfiTable(rows){
    const statusColor = {
      'Accepted':'#16a085', 'Accepted with Punch':'#f4b942',
      'Open':'#c53030', 'Hold by EACOP':'#2563eb', 'No RFI Yet':'#999', 'Other':'#999'
    };
    document.getElementById('rfiTableBody').innerHTML = rows.map(r=>`
      <tr style="border-bottom:1px solid var(--border);">
        <td style="padding:7px;">${r.asset}</td>
        <td style="padding:7px;">${r.rfi_no||'-'}</td>
        <td style="padding:7px;">${r.disc}</td>
        <td style="padding:7px;color:${statusColor[r.status]||'var(--text)'};font-weight:600;">${r.status}</td>
        <td style="padding:7px;">${r.date}</td>
      </tr>`).join('');
  }
  renderRfiTable(RFI.recent);
  document.getElementById('rfiSearch').addEventListener('input', e=>{
    const q = e.target.value.trim().toLowerCase();
    const filtered = !q ? RFI.recent : RFI.recent.filter(r =>
      r.asset.toLowerCase().includes(q) || (r.rfi_no||'').toLowerCase().includes(q));
    renderRfiTable(filtered);
  });
} else {
  ['sec-rfi-kpi','sec-rfi-status','sec-rfi-daily','sec-rfi-weekly','sec-rfi-monthly','sec-rfi-type','sec-rfi-subsystems','sec-rfi-recent','sec-export-rfi']
    .forEach(id=>{
      document.getElementById(id).innerHTML = '<div class="section-title">⚠️ Inspection Register file not found</div>';
    });
}

// ---- RFI Inspection Summary table (export section 7) ----
(function(){
  const data = RFI && RFI.inspection_summary;
  const body = document.getElementById('rfiInspectionBody');
  const foot = document.getElementById('rfiInspectionFoot');
  if(!data || !data.rows || !data.rows.length){
    if(body) body.innerHTML = '<tr><td colspan="5" style="color:#999;text-align:center;padding:20px;">No inspection data available</td></tr>';
    return;
  }
  const rows = data.rows;
  body.innerHTML = rows.map(r => {
    return '<tr>' +
      '<td style="border:1px solid var(--border);text-align:center;font-weight:bold;background:' + (r.discipline=='E'?'#FCE4D6':r.discipline=='I'?'#FFF2CC':'#D9D9D9') + ';">' + r.discipline + '</td>' +
      '<td style="border:1px solid var(--border);text-align:center;">' + r.assets + '</td>' +
      '<td style="border:1px solid var(--border);text-align:center;background:var(--green-bg);font-weight:700;font-size:14px;color:' + (r.laying_accepted>0?'var(--green)':'var(--text3)') + ';">' + (r.laying_accepted>0?'✅ ':'❌ ') + r.laying_accepted + '/' + r.laying_submitted + '</td>' +
      '<td style="border:1px solid var(--border);text-align:center;background:var(--green-bg);font-weight:700;font-size:14px;color:' + (r.testing_accepted>0?'var(--green)':'var(--text3)') + ';">' + (r.testing_accepted>0?'✅ ':'❌ ') + r.testing_accepted + '/' + r.testing_submitted + '</td>' +
      '<td style="border:1px solid var(--border);text-align:center;background:var(--green-bg);font-weight:700;font-size:14px;color:' + (r.term_accepted>0?'var(--green)':'var(--text3)') + ';">' + (r.term_accepted>0?'✅ ':'❌ ') + r.term_accepted + '/' + r.term_submitted + '</td>' +
    '</tr>';
  }).join('');

  const t = data.totals;
  foot.innerHTML = '<tr style="background:var(--card3);color:var(--text);font-weight:800;border-top:2px solid var(--gold);">' +
    '<td style="border:1px solid var(--border);padding:7px;text-align:center;">GRAND TOTAL</td>' +
    '<td style="border:1px solid var(--border);text-align:center;">' + t.total_assets + '</td>' +
    '<td style="border:1px solid var(--border);text-align:center;color:var(--green);">✅ ' + t.laying_accepted + '</td>' +
    '<td style="border:1px solid var(--border);text-align:center;color:var(--green);">✅ ' + t.testing_accepted + '</td>' +
    '<td style="border:1px solid var(--border);text-align:center;color:var(--green);">✅ ' + t.term_accepted + '</td>' +
  '</tr>';
})();

function exportRfiInspectionExcel(){
  const data = RFI && RFI.inspection_summary;
  if(!data || !data.rows || !data.rows.length){ alert('No data'); return; }
  let html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
    '<head><meta charset="UTF-8">' +
    '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Sheet1</x:Name>' +
    '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
    '<style>td,th{font-family:Arial,sans-serif;font-size:12px;vertical-align:middle;}.hdr{background:#ED7D31;color:#000;font-weight:bold;text-align:center;border:1px solid #9c4a14;padding:8px;}</style>' +
    '</head><body>' +
    '<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;min-width:500px;">' +
    '<tr><th colspan="5" class="hdr" style="font-size:14px;">EACOP PIPELINE PROJECT &mdash; RFI Inspection Summary &mdash; By Discipline</th></tr>' +
    '<tr><th class="hdr">Discipline</th><th class="hdr">Assets</th><th class="hdr">LAYING RFI ✓</th><th class="hdr">TESTING RFI ✓</th><th class="hdr">TERMINATION RFI ✓</th></tr>';
  data.rows.forEach(r => {
    html += '<tr>' +
      '<td style="border:1px solid #bbb;text-align:center;font-weight:bold;">' + r.discipline + '</td>' +
      '<td style="border:1px solid #bbb;text-align:center;">' + r.assets + '</td>' +
      '<td style="border:1px solid #bbb;text-align:center;background:#e8f5e9;font-weight:bold;">' + (r.laying_accepted>0?'✅ ':'❌ ') + r.laying_accepted + '/' + r.laying_submitted + '</td>' +
      '<td style="border:1px solid #bbb;text-align:center;background:#e8f5e9;font-weight:bold;">' + (r.testing_accepted>0?'✅ ':'❌ ') + r.testing_accepted + '/' + r.testing_submitted + '</td>' +
      '<td style="border:1px solid #bbb;text-align:center;background:#e8f5e9;font-weight:bold;">' + (r.term_accepted>0?'✅ ':'❌ ') + r.term_accepted + '/' + r.term_submitted + '</td>' +
    '</tr>';
  });
  const t = data.totals;
  html += '<tr style="background:#404040;color:#fff;font-weight:bold;">' +
    '<td style="border:1px solid #222;text-align:center;">GRAND TOTAL</td>' +
    '<td style="border:1px solid #222;text-align:center;">' + t.total_assets + '</td>' +
    '<td style="border:1px solid #222;text-align:center;">✅ ' + t.laying_accepted + '</td>' +
    '<td style="border:1px solid #222;text-align:center;">✅ ' + t.testing_accepted + '</td>' +
    '<td style="border:1px solid #222;text-align:center;">✅ ' + t.term_accepted + '</td>' +
  '</tr></table></body></html>';
  downloadExcel(html, 'RFI_Inspection_Summary');
}

// ---- RFI Detail table (per asset from cable tracker) ----
(function(){
  const ALL = CABLE_TRACKER && CABLE_TRACKER.detail;
  const body = document.getElementById('rfiDetailBody');
  const count = document.getElementById('rfiDetailCount');
  const search = document.getElementById('rfiDetailSearch');
  if(!ALL || !ALL.length){
    if(body) body.innerHTML = '<tr><td colspan="8" style="color:#999;text-align:center;padding:20px;">No cable tracker data</td></tr>';
    return;
  }
  const discBg = {'E':'#FCE4D6','I':'#FFF2CC','T':'#D9D9D9'};
  function render(data){
    const shown = data.slice(0, 800);
    body.innerHTML = shown.map(r =>
      '<tr>' +
        '<td style="border:1px solid var(--border);padding:3px 4px;">' + r.asset + '</td>' +
        '<td style="border:1px solid var(--border);padding:3px 4px;">' + (r.subsystem||'') + '</td>' +
        '<td style="border:1px solid var(--border);padding:3px 4px;text-align:center;background:' + (discBg[r.disc]||'') + ';font-weight:600;">' + r.disc + '</td>' +
        '<td style="border:1px solid var(--border);padding:3px 4px;">' + r.desc.slice(0,60) + '</td>' +
        '<td style="border:1px solid var(--border);padding:3px 4px;">' + r.scope.slice(0,25) + '</td>' +
        '<td style="border:1px solid var(--border);padding:3px 4px;text-align:center;font-size:13px;font-weight:700;color:' + (r.laying_rfi?'var(--green)':'var(--text3)') + ';">' + (r.laying_rfi?'✅':'❌') + '</td>' +
        '<td style="border:1px solid var(--border);padding:3px 4px;text-align:center;font-size:13px;font-weight:700;color:' + (r.testing_rfi?'var(--green)':'var(--text3)') + ';">' + (r.testing_rfi?'✅':'❌') + '</td>' +
        '<td style="border:1px solid var(--border);padding:3px 4px;text-align:center;font-size:13px;font-weight:700;color:' + (r.term_rfi?'var(--green)':'var(--text3)') + ';">' + (r.term_rfi?'✅':'❌') + '</td>' +
      '</tr>'
    ).join('');
    count.innerText = data.length > 800
      ? 'Showing first 800 of ' + data.length + ' / ' + ALL.length + ' assets — use search to filter'
      : 'Showing ' + data.length + ' / ' + ALL.length + ' assets';
  }
  render(ALL);
    if(search){
    search.addEventListener('input', function(){
      const q = this.value.trim().toLowerCase();
      const filtered = !q ? ALL : ALL.filter(r =>
        r.asset.toLowerCase().includes(q) ||
        (r.subsystem||'').toLowerCase().includes(q) ||
        (r.desc||'').toLowerCase().includes(q) ||
        (r.scope||'').toLowerCase().includes(q) ||
        (r.disc||'').toLowerCase().includes(q));
      render(filtered);
    });
  }
})();

function exportRfiDetailExcel(){
  const d = CABLE_TRACKER && CABLE_TRACKER.detail;
  if(!d || !d.length){ alert('No data'); return; }
  let html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
    '<head><meta charset="UTF-8">' +
    '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Sheet1</x:Name>' +
    '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
    '<style>td,th{font-family:Arial,sans-serif;font-size:11px;vertical-align:middle;}.hdr{background:#ED7D31;color:#000;font-weight:bold;text-align:center;border:1px solid #9c4a14;padding:6px;}</style>' +
    '</head><body>' +
    '<table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse;">' +
    '<tr><th colspan="8" class="hdr" style="font-size:13px;">EACOP PIPELINE PROJECT &mdash; RFI Detail by Asset (Cable Tracker) &mdash; ' + d.length + ' items</th></tr>' +
    '<tr><th class="hdr" style="text-align:left;">Asset Tag</th><th class="hdr" style="text-align:left;">Subsystem</th><th class="hdr">Disc</th><th class="hdr" style="text-align:left;">Description</th><th class="hdr">Scope</th><th class="hdr">LAYING RFI</th><th class="hdr">TESTING RFI</th><th class="hdr">TERMINATION RFI</th></tr>';
  d.forEach(r => {
    html += '<tr>' +
      '<td style="border:1px solid #ccc;padding:3px 5px;">' + r.asset + '</td>' +
      '<td style="border:1px solid #ccc;padding:3px 5px;">' + (r.subsystem||'') + '</td>' +
      '<td style="border:1px solid #ccc;padding:3px 5px;text-align:center;">' + r.disc + '</td>' +
      '<td style="border:1px solid #ccc;padding:3px 5px;">' + r.desc + '</td>' +
      '<td style="border:1px solid #ccc;padding:3px 5px;text-align:center;">' + r.scope + '</td>' +
      '<td style="border:1px solid #ccc;padding:3px 5px;text-align:center;font-weight:bold;">' + (r.laying_rfi?'✅':'❌') + '</td>' +
      '<td style="border:1px solid #ccc;padding:3px 5px;text-align:center;font-weight:bold;">' + (r.testing_rfi?'✅':'❌') + '</td>' +
      '<td style="border:1px solid #ccc;padding:3px 5px;text-align:center;font-weight:bold;">' + (r.term_rfi?'✅':'❌') + '</td>' +
    '</tr>';
  });
  html += '</table></body></html>';
  downloadExcel(html, 'RFI_Detail_by_Asset');
}

// ---------- Punch Tracking & Closure ----------
if(PUNCH){
  (function(){
    const c = document.getElementById('punchTrackingKpis');
    const open = PUNCH.status_counts['Open']||0;
    const closed = Object.entries(PUNCH.status_counts).filter(([k])=>k!=='Open').reduce((a,[,v])=>a+v,0);
    const total = PUNCH.total;
    const pct = total ? Math.round(closed/total*100) : 0;
    const kpis = [
      {icon:'📌', val:total, lbl:'Total Punch Items', cls:''},
      {icon:'🟡', val:open, lbl:'Open (Backlog)', cls:'pink'},
      {icon:'✅', val:closed, lbl:'Closed / Resolved', cls:'teal'},
      {icon:'📈', val:pct+'%', lbl:'Closure Rate', cls:'gold'},
    ];
    kpis.forEach(k=>{
      const div = document.createElement('div');
      div.className='kpi '+k.cls;
      div.innerHTML = `<div class="icon">${k.icon}</div><div class="val">${k.val}</div><div class="lbl">${k.lbl}</div>`;
      c.appendChild(div);
    });
  })();

  // Open backlog by subsystem (reuse top_subsystems as proxy for backlog)
  new Chart(document.getElementById('chartPunchBacklog'), {
    type:'bar',
    data:{ labels: PUNCH.top_subsystems.map(r=>r.label),
      datasets:[{ label:'Open Punch Items', data: PUNCH.top_subsystems.map(r=>r.count), backgroundColor:'#c53030', borderRadius:6, datalabels:{...DL_BAR, anchor:'end', align:'right'} }]},
    options:{ indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{beginAtZero:true}} }
  });

  // Punch by discipline - open vs closed
  const discLabels = Object.keys(PUNCH.disc_counts);
  new Chart(document.getElementById('chartPunchClosure'), {
    type:'bar',
    data:{
      labels: discLabels,
      datasets:[
        {label:'Open', data: discLabels.map(k=>PUNCH.disc_counts_open[k]||0), backgroundColor:'#c53030', datalabels: DL_BAR},
        {label:'Closed', data: discLabels.map(k=>PUNCH.disc_counts_closed[k]||0), backgroundColor:'#1a8a4a', datalabels: DL_BAR},
      ]
    },
    options:{ plugins:{legend:{position:'bottom'}}, scales:{y:{beginAtZero:true}} }
  });
} else {
  document.getElementById('sec-punch-tracking').innerHTML =
    '<div class="section-title">⚠️ Punch List Register file not found</div>';
}

// ---------- Universal Search ----------
document.getElementById('searchLibSize').innerText = Object.keys(SI).length.toLocaleString();

const statusColors = {
  'Closed':'#1a8a4a','To be completed':'#c8940a','Submitted':'#2563eb',
  'Accepted':'#1a8a4a','Accepted with Punch':'#c8940a','Open':'#c53030',
  'Hold by EACOP':'#2563eb','No RFI Yet':'#9a8d7c','Other':'#9a8d7c','Rejected':'#c53030',
};
const discColors = {'E':'#2563eb','I':'#7c3aed','T':'#0891b2'};
const MS_LOOKUP = Object.fromEntries((ITR.milestone_summary||[]).map(s=>[s.raw, s]));
let msSearchCharts = [];

function badge(text, color){
  return `<span style="display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;
    background:${color}22;color:${color};border:1px solid ${color}44;">${text}</span>`;
}

function renderSearchResults(query){
  const box = document.getElementById('searchResults');
  msSearchCharts.forEach(ch=>{ try{ ch.destroy(); }catch(e){} });
  msSearchCharts = [];
  if(!query || query.length < 3){
    box.innerHTML = '<div style="color:var(--muted);padding:20px;text-align:center;">Type at least 3 characters (Asset Tag / Task ID / RFI No / Subsystem / Milestone)...</div>';
    return;
  }
  const q = query.trim().toLowerCase();
  const seen = new Set();
  let matchedTags = [];

  function addTag(t){ if(t && SI[t] && !seen.has(t)){ seen.add(t); matchedTags.push(t); } }

  // 1. Exact Asset Tag
  for(const tag in SI){ if(tag.toLowerCase()===q){ addTag(tag); break; } }

  // 2. Task ID exact -> get asset tag
  if(TID_MAP[q.toUpperCase()]) addTag(TID_MAP[q.toUpperCase()]);

  // 3. RFI No partial
  for(const rn in RFI_MAP) if(rn.toLowerCase().includes(q)) RFI_MAP[rn].forEach(addTag);

  // 4. Subsystem partial
  for(const s in SUB_MAP) if(s.toLowerCase().includes(q)) SUB_MAP[s].forEach(addTag);

  // 5. Milestone partial
  const matchedMilestones = [];
  for(const p in PRI_MAP){
    if(p.toLowerCase().includes(q)){
      PRI_MAP[p].forEach(addTag);
      if(MS_LOOKUP[p] && !matchedMilestones.some(m=>m.raw===p)) matchedMilestones.push(MS_LOOKUP[p]);
    }
  }

  // 6. Partial Asset Tag or Task ID
  for(const tag in SI){
    if(seen.has(tag)) continue;
    const srcs = SI[tag];
    if(tag.toLowerCase().includes(q) ||
       srcs.itr.some(it => it.id && it.id.toLowerCase().includes(q))){
      addTag(tag);
    }
    if(matchedTags.length >= 50) break;
  }

  if(!matchedTags.length && !matchedMilestones.length){
    box.innerHTML = '<div style="color:var(--accent);padding:20px;text-align:center;">⚠️ No results found.</div>';
    return;
  }

  const msSummaryHtml = matchedMilestones.map((s,idx)=>`
    <div style="border:1px solid var(--border);border-radius:12px;padding:14px 16px;margin-bottom:14px;background:var(--panel2);">
      <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;">
        <div style="font-size:15px;font-weight:800;color:var(--gold);">🎯 ${s.label} — Progress (CPP AGI)</div>
        <div style="font-size:13px;color:var(--text);font-weight:700;">${s.closed} / ${s.total} Closed — ${s.pct}%</div>
      </div>
      <div class="progress-bar"><div class="progress-fill" style="width:${s.pct}%"></div></div>
      <div style="margin-top:12px;"><canvas id="msSearchChart${idx}" height="90"></canvas></div>
    </div>`).join('');

  box.innerHTML = msSummaryHtml + `<div style="color:var(--muted);font-size:12px;margin-bottom:10px;">
    Found ${matchedTags.length} result(s)</div>` +
  matchedTags.map(tag=>{
    const srcs = SI[tag];
    const totalTasks  = srcs.itr.length;
    const closedTasks = srcs.itr.filter(it=>it.closed).length;
    const pendingTasks= totalTasks - closedTasks;
    const first = srcs.itr[0] || {};
    const discColor = discColors[first.d] || '#9a8d7c';

    const itrRows = srcs.itr.map(it=>`
      <tr style="border-bottom:1px solid var(--border);background:${it.closed?'rgba(0,230,118,.04)':'rgba(255,77,141,.03)'}">
        <td style="padding:5px;color:var(--gold);font-weight:700;">${it.id||'-'}</td>
        <td style="padding:5px;">${it.ty||'-'}</td>
        <td style="padding:5px;">${badge(it.disc||it.d||'-', discColors[it.d]||'#9a8d7c')}</td>
        <td style="padding:5px;">
          ${it.closed
            ? '<span style="color:#1a8a4a;font-weight:800;">✅ Closed</span>'
            : `<span style="color:#c53030;font-weight:700;">⏳ ${it.st||'Pending'}</span>`}
        </td>
        <td style="padding:5px;color:${it.cd?'#1a8a4a':'var(--muted)'};">${it.cd||'-'}</td>
        <td style="padding:5px;color:var(--muted);font-size:11px;">${(it.sub||'').split(' - ')[0]||'-'}</td>
        <td style="padding:5px;">${it.ms?badge('🎯 '+it.ms,'#c8940a'):'-'}</td>
      </tr>`).join('');

    const rfiRows = srcs.rfi.map(r=>`
      <tr style="border-bottom:1px solid var(--border);">
        <td style="padding:5px;color:var(--blue);font-weight:700;">${r.rfi_no||'-'}</td>
        <td style="padding:5px;">${r.discipline||'-'}</td>
        <td style="padding:5px;">${(r.type||'-').slice(0,20)}</td>
        <td style="padding:5px;">${badge(r.status, statusColors[r.status]||'#9a8d7c')}</td>
        <td style="padding:5px;">${r.date||'-'}</td>
      </tr>`).join('');

    const punchRows = srcs.punch.map(p=>`
      <tr style="border-bottom:1px solid var(--border);">
        <td style="padding:5px;color:var(--accent);font-weight:700;">${p.plid||'-'}</td>
        <td style="padding:5px;">${p.discipline||'-'}</td>
        <td style="padding:5px;">${p.category||'-'}</td>
        <td style="padding:5px;">${badge(p.status, statusColors[p.status]||'#c53030')}</td>
        <td style="padding:5px;">${p.rfi_no||'-'}</td>
        <td style="padding:5px;">${p.date||'-'}</td>
        <td style="padding:5px;color:var(--muted);">${p.desc||'-'}</td>
      </tr>`).join('');

    return `
    <div style="border:1px solid var(--border);border-radius:12px;padding:16px;margin-bottom:14px;background:var(--panel2);">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
        <div>
          <div style="font-size:17px;font-weight:800;color:var(--teal);">📍 ${tag}</div>
          <div style="font-size:12px;margin-top:4px;">
            ${badge(first.disc||first.d||'Unknown', discColor)}
            <span style="color:#1a8a4a;font-weight:700;margin-left:6px;">✅ ${closedTasks} Closed</span>
            ${pendingTasks>0?`<span style="color:#c53030;font-weight:700;margin-left:4px;">⏳ ${pendingTasks} Pending</span>`:''}
            ${first.ms ? badge('🎯 '+first.ms,'#c8940a') : ''}
            ${first.res ? `<span style="color:var(--muted);font-size:11px;margin-left:6px;">${first.res}</span>` : ''}
            ${first.precom ? badge('🔌 '+first.precom,'#7c3aed') : ''}
          </div>
        </div>
        <div style="text-align:right;font-size:11px;color:var(--muted);">
          ${first.sub ? `<div style="color:var(--muted)">${first.sub}</div>` : ''}
          ${first.cd ? `<div style="color:var(--teal)">Closed: ${first.cd}</div>` : ''}
        </div>
      </div>

      <!-- ITR -->
      <div style="margin-bottom:12px;">
        <div style="font-weight:700;color:var(--text);font-size:13px;margin-bottom:6px;">
          🧪 ITR / Tasks <span style="color:var(--gold)">(${srcs.itr.length})</span></div>
        ${srcs.itr.length ? `<div style="overflow-x:auto;">
          <table style="width:100%;font-size:12px;border-collapse:collapse;min-width:600px;">
            <tr style="color:var(--teal);border-bottom:1px solid var(--border);">
              <th style="text-align:left;padding:5px;">Task ID</th>
              <th style="text-align:left;padding:5px;">Type</th>
              <th style="padding:5px;">Discipline</th>
              <th style="padding:5px;">State</th>
              <th style="padding:5px;">Closing Date</th>
              <th style="text-align:left;padding:5px;">Subsystem</th>
              <th style="padding:5px;">Milestone</th>
            </tr>${itrRows}
          </table></div>` : '<div style="color:var(--muted);font-size:12px;">No ITR records</div>'}
      </div>

      <!-- RFI -->
      <div style="margin-bottom:12px;">
        <div style="font-weight:700;color:var(--text);font-size:13px;margin-bottom:6px;">
          📝 RFI <span style="color:var(--blue)">(${srcs.rfi.length})</span></div>
        ${srcs.rfi.length ? `<div style="overflow-x:auto;">
          <table style="width:100%;font-size:12px;border-collapse:collapse;min-width:500px;">
            <tr style="color:var(--teal);border-bottom:1px solid var(--border);">
              <th style="text-align:left;padding:5px;">RFI No</th>
              <th style="padding:5px;">Discipline</th>
              <th style="padding:5px;">Type</th>
              <th style="padding:5px;">Status</th>
              <th style="padding:5px;">Date</th>
            </tr>${rfiRows}
          </table></div>` : '<div style="color:var(--muted);font-size:12px;">No RFI submitted yet</div>'}
      </div>

      <!-- Punch -->
      <div>
        <div style="font-weight:700;color:var(--text);font-size:13px;margin-bottom:6px;">
          📌 Punch List <span style="color:var(--accent)">(${srcs.punch.length})</span></div>
        ${srcs.punch.length ? `<div style="overflow-x:auto;">
          <table style="width:100%;font-size:12px;border-collapse:collapse;min-width:700px;">
            <tr style="color:var(--teal);border-bottom:1px solid var(--border);">
              <th style="text-align:left;padding:5px;">PL ID</th>
              <th style="padding:5px;">Discipline</th>
              <th style="padding:5px;">Category</th>
              <th style="padding:5px;">Status</th>
              <th style="padding:5px;">RFI No</th>
              <th style="padding:5px;">Raised Date</th>
              <th style="text-align:left;padding:5px;">Description</th>
            </tr>${punchRows}
          </table></div>` : '<div style="color:var(--muted);font-size:12px;">No punch items ✅</div>'}
      </div>
    </div>`;
  }).join('');

  matchedMilestones.forEach((s,idx)=>{
    const canvas = document.getElementById('msSearchChart'+idx);
    if(!canvas) return;
    const ch = new Chart(canvas, {
      type:'bar',
      data:{ labels:[s.label],
        datasets: MS_STATUS_ORDER.map(st=>({
          label: st,
          data: [(s.status&&s.status[st]) || 0],
          backgroundColor: MS_STATUS_COLORS[st],
          datalabels: DL_STACK
        }))
      },
      options:{ indexAxis:'y', plugins:{legend:{position:'bottom'}}, scales:{ x:{stacked:true, beginAtZero:true}, y:{stacked:true} } }
    });
    msSearchCharts.push(ch);
  });
}

renderSearchResults('');
document.getElementById('universalSearch').addEventListener('input', e=>{
  renderSearchResults(e.target.value);
});

// ---------- Extra Tabs (PS5 EIT CPP AGI Dashboard sheets) ----------
(function(){
  var pages = EIT_PAGES || {};
  var order = Object.keys(pages);
  var bar = document.getElementById('mainTabBar');
  var container = document.getElementById('tab-container');
  if(!bar || !container || !order.length) return;

  // Build tab buttons
  var pageIcons = [
    {match:/^dashboard$/i, icon:'🎯'},
    {match:/subsystem report/i, icon:'🗂️'},
    {match:/electrical/i, icon:'⚡'},
    {match:/instrumentation/i, icon:'🔧'},
    {match:/telecom/i, icon:'📡'},
    {match:/cable routes/i, icon:'🧵'},
    {match:/route chains/i, icon:'🔗'},
    {match:/panel connections/i, icon:'🖥️'},
    {match:/itr tasks/i, icon:'✅'},
    {match:/punch list/i, icon:'📌'},
    {match:/qc punch/i, icon:'🛠️'},
    {match:/inspection/i, icon:'🔍'},
    {match:/subsystem summary/i, icon:'📈'},
  ];
  function pageIcon(name){
    for(var i=0;i<pageIcons.length;i++) if(pageIcons[i].match.test(name)) return pageIcons[i].icon;
    return '📄';
  }
  function tabBtnHtml(name){
    return '<span class="t-icon">' + pageIcon(name) + '</span>' + name;
  }
  var btn = document.createElement('button');
  btn.className = 'tabbtn active';
  btn.innerHTML = '<span class="t-icon">🏠</span>Main Dashboard';
  btn.dataset.tab = 'tab-main-dashboard';
  bar.appendChild(btn);
  order.forEach(function(name, idx){
    var b = document.createElement('button');
    b.className = 'tabbtn';
    b.dataset.tab = 'tab-' + idx;
    b.innerHTML = tabBtnHtml(name);
    bar.appendChild(b);
  });

  // Build page divs (KPI row + charts + table)
  var pageCharts = {};
  var pagesCache = {};
  var builtIdx = {};
  var PAL = ['#2563eb','#7c3aed','#0891b2','#1a8a4a','#c8940a','#c53030','#ec4899','#f97316','#14b8a6','#6366f1','#84cc16','#06b6d4','#ef4444','#8b5cf6'];
  var HEADER_RE = /equipment|description|status|discipline|subsystem|task|panel|level|cable|metric|type|role|tag|state|category|punch|inspection|rfi|direction|route|asset|close|complete|number|name|unit|wire|from|to|fed|length|raised|serial/i;

  order.forEach(function(name, idx){
    var div = document.createElement('div');
    div.className = 'tabpage';
    div.id = 'tab-' + idx;
    div.innerHTML =
      '<div class="section-title">' + pageIcon(name) + ' ' + name + '</div>' +
      '<div class="kpi-row" id="pkpi-' + idx + '"></div>' +
      '<div class="chart-row" id="pcharts-' + idx + '"></div>' +
      '<div class="chart-card">' +
        '<div class="eit-page-toolbar">' +
          '<input type="text" placeholder="Search ' + name + '..." data-search="' + idx + '">' +
          '<span class="count" id="count-' + idx + '"></span>' +
        '</div>' +
        '<div class="eit-page-table-wrap" id="wrap-' + idx + '"></div>' +
      '</div>';
    container.appendChild(div);
  });

  function esc(s){
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }
  function parseNum(v){
    var n = parseFloat(String(v).replace(/[^0-9.\-]/g,''));
    return isNaN(n) ? 0 : n;
  }
  function fmt(n){ return (n>=10000 ? (n/1000).toFixed(1)+'k' : Math.round(n).toLocaleString()); }
  function kpiCard(icon,val,lbl,cls){
    return '<div class="kpi ' + (cls||'') + '"><div class="icon">' + icon + '</div><div class="val">' + val + '</div><div class="lbl">' + lbl + '</div></div>';
  }
  function progressBar(pct, extra){
    return '<div class="progress-bar" style="height:10px;background:#e7dfcf;border-radius:6px;overflow:hidden;margin-top:8px;">' +
      '<div class="progress-fill" style="height:100%;width:' + Math.min(100,Math.max(0,pct)) + '%;background:linear-gradient(90deg,' + (pct>=80?'#1a8a4a':pct>=50?'#f59e0b':'#c53030') + ',' + (pct>=80?'#22c55e':pct>=50?'#fbbf24':'#ef4444') + ');border-radius:6px;transition:width .8s;"></div></div>' + (extra||'');
  }

  // ---- generic header detection ----
  function findHeader(rows){
    for(var i=1;i<Math.min(rows.length,15);i++){
      var row = rows[i]||[];
      var ne = row.filter(function(c){ return String(c).trim()!==''; });
      if(ne.length>=3){
        var kws = ne.filter(function(c){ return HEADER_RE.test(String(c)); });
        if(kws.length>=2) return {idx:i, headers:ne.slice(0,row.length)};
      }
    }
    return null;
  }
  function toRecords(rows, hIdx, headers){
    var recs = [];
    for(var i=hIdx+1;i<rows.length;i++){
      var row = rows[i]||[];
      var vals = [];
      for(var c=0;c<headers.length;c++) vals.push(row[c]!==undefined ? String(row[c]).trim() : '');
      var ne = vals.filter(function(v){ return v!==''; });
      if(ne.length<2) continue;
      if(/^total/i.test(vals.join(' '))) continue;
      if(/^(total|subtotal)$/i.test(ne[0])) continue;
      var rec = {};
      for(var k=0;k<headers.length;k++) rec['c'+k] = vals[k];
      recs.push(rec);
    }
    return recs;
  }
  function countBy(recs, col){
    var m = {};
    recs.forEach(function(r){
      var k = r[col]||'';
      if(!k) return;
      k = k.split(/[,;]/)[0].trim();
      if(k) m[k]=(m[k]||0)+1;
    });
    return m;
  }
  function sumCol(recs, col){
    var s=0;
    recs.forEach(function(r){ s+=parseNum(r[col]); });
    return s;
  }
  function pickCol(headers, pats){
    for(var p=0;p<pats.length;p++){
      for(var i=0;i<headers.length;i++){
        if(new RegExp(pats[p],'i').test(String(headers[i]))) return i;
      }
    }
    return -1;
  }
  function isClosedish(v){
    return /(closed|complete|done|accepted)/i.test(v) && !/open/i.test(v) && !/not complete|to be completed/i.test(v);
  }

  // ---- fallback raw table ----
  var ROW_CAP = 800;
  function renderRawTable(idx, rows){
    var filtered = rows;
    if(pagesCache[idx].filter){
      var q = pagesCache[idx].filter.toLowerCase();
      filtered = rows.filter(function(r){ return r.join(' ').toLowerCase().indexOf(q)>=0; });
    }
    var total = filtered.length;
    var limited = filtered.slice(0, ROW_CAP);
    var maxCols = 0;
    limited.forEach(function(r){ if(r.length>maxCols) maxCols=r.length; });
    var html = '<table><thead><tr>';
    for(var c=0;c<maxCols;c++) html += '<th>' + (c===0?'#':'Col '+c) + '</th>';
    html += '</tr></thead><tbody>';
    limited.forEach(function(r, ri){
      html += '<tr><td>' + (ri+1) + '</td>';
      for(var c=0;c<maxCols;c++) html += '<td>' + esc(r[c]||'') + '</td>';
      html += '</tr>';
    });
    html += '</tbody></table>';
    if(total > ROW_CAP) html += '<div style="padding:10px;text-align:center;font-size:12px;color:var(--text3);">' +
      'Showing first ' + ROW_CAP.toLocaleString() + ' of ' + total.toLocaleString() + ' rows — use search to filter.</div>';
    document.getElementById('wrap-'+idx).innerHTML = html;
    document.getElementById('count-'+idx).textContent = total + ' / ' + rows.length + ' rows';
  }

  // ---- data table with real headers ----
  function renderTable(idx, rows, headers, recs){
    var filtered = recs;
    var f = pagesCache[idx].filter;
    if(f){
      var q = f.toLowerCase();
      filtered = recs.filter(function(r){
        for(var k=0;k<headers.length;k++) if(r['c'+k].toLowerCase().indexOf(q)>=0) return true;
        return false;
      });
    }
    document.getElementById('count-'+idx).textContent = filtered.length + ' / ' + rows.length + ' rows';
    if(!filtered.length){
      document.getElementById('wrap-'+idx).innerHTML = '<div style="padding:20px;color:var(--muted);">No results</div>';
      return;
    }
    var total = filtered.length;
    var limited = filtered.slice(0, ROW_CAP);
    var html = '<table><thead><tr><th>#</th>';
    headers.forEach(function(h){ html += '<th>' + esc(h) + '</th>'; });
    html += '</tr></thead><tbody>';
    limited.forEach(function(r, ri){
      html += '<tr><td>' + (ri+1) + '</td>';
      for(var k=0;k<headers.length;k++){
        var v = r['c'+k];
        var extra = '';
        if(isClosedish(v)) extra = ' style="color:#1a8a4a;font-weight:700;"';
        else if(/open|originated|to be completed|pending/i.test(v)) extra = ' style="color:#c53030;font-weight:700;"';
        html += '<td' + extra + '>' + esc(v) + '</td>';
      }
      html += '</tr>';
    });
    html += '</tbody></table>';
    if(total > ROW_CAP) html += '<div style="padding:10px;text-align:center;font-size:12px;color:var(--text3);">' +
      'Showing first ' + ROW_CAP.toLocaleString() + ' of ' + total.toLocaleString() + ' rows — use search to filter.</div>';
    document.getElementById('wrap-'+idx).innerHTML = html;
  }

  function initChart(idx, canvasId, type, labels, datasets, isBar){
    var el = document.getElementById(canvasId);
    if(!el) return null;
    var chart = new Chart(el.getContext('2d'), {
      type: type,
      data: { labels: labels, datasets: datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position:'bottom', labels:{ color:'#6b5e4d', font:{size:11}, boxWidth:12 } },
          tooltip: { backgroundColor:'#2c2416', titleColor:'#f5f0e8', bodyColor:'#f5f0e8' }
        },
        scales: isBar ? {
          x: { ticks:{ color:'#6b5e4d' }, grid:{ color:'rgba(107,94,77,0.12)' }, stacked:false },
          y: { beginAtZero:true, ticks:{ color:'#6b5e4d' }, grid:{ color:'rgba(107,94,77,0.12)' } }
        } : {}
      }
    });
    if(!pageCharts[idx]) pageCharts[idx] = [];
    pageCharts[idx].push(chart);
    return chart;
  }

  function buildTablePage(idx, name, rows){
    pagesCache[idx] = pagesCache[idx] || {filter:''};
    var hdr = findHeader(rows);
    if(!hdr){ renderRawTable(idx, rows); return; }
    var headers = hdr.headers;
    var recs = toRecords(rows, hdr.idx, headers);
    pagesCache[idx].rows = rows;
    pagesCache[idx].headers = headers;
    pagesCache[idx].recs = recs;

    var statusCol = pickCol(headers, ['status','state','workflow','close','accept','complete']);
    var subCol = pickCol(headers, ['subsystem']);
    var catCol = pickCol(headers, ['category','task type','type','role','direction','level','scope']);
    var lenCol = pickCol(headers, ['length','km']);
    var totalCol = pickCol(headers, ['total','itrs','tasks']);

    var closedCount = 0;
    if(statusCol>=0) recs.forEach(function(r){ if(isClosedish(r['c'+statusCol])) closedCount++; });
    else if(totalCol>=0 && pickCol(headers,['closed'])>=0) closedCount = sumCol(recs, 'c'+pickCol(headers,['closed']));
    var pct = recs.length ? Math.round(closedCount/recs.length*100) : 0;

    var kpis = [
      kpiCard('📊', fmt(recs.length), 'Total Records'),
      kpiCard('✅', fmt(closedCount), 'Closed / Done', 'teal'),
      kpiCard('📈', pct+'%', 'Completion', 'gold'),
    ];
    if(lenCol>=0) kpis.push(kpiCard('🧵', (sumCol(recs,'c'+lenCol)/1000).toFixed(1)+' km', 'Total ' + esc(headers[lenCol])));
    if(subCol>=0){
      var uniqSubs = Object.keys(countBy(recs, 'c'+subCol)).length;
      kpis.push(kpiCard('🗂️', fmt(uniqSubs), 'Unique Subsystems', 'blue'));
    }
    document.getElementById('pkpi-'+idx).innerHTML = '<div class="kpi-row">' + kpis.join('') + '</div>';

    var charts = '';
    var chartCards = [];
    if(statusCol>=0){
      var cm = Object.entries(countBy(recs, 'c'+statusCol)).sort(function(a,b){return b[1]-a[1];});
      var slice = cm.slice(0,8);
      chartCards.push({id:'ch-d-'+idx, title:'Breakdown by ' + esc(headers[statusCol]), type:'doughnut',
        labels: slice.map(function(e){return e[0];}),
        data: slice.map(function(e){return e[1];}),
        sum: cm.reduce(function(a,e){return a+e[1];},0)});
    }
    var barCol = subCol>=0 ? subCol : catCol;
    if(barCol>=0){
      var bm = Object.entries(countBy(recs, 'c'+barCol)).sort(function(a,b){return b[1]-a[1];}).slice(0,15);
      chartCards.push({id:'ch-b-'+idx, title:'Top 15 by ' + esc(headers[barCol]), type:'bar',
        labels: bm.map(function(e){return e[0];}),
        data: bm.map(function(e){return e[1];})});
    }
    if(totalCol>=0 && catCol>=0){
      // stacked total by category
      var tm = Object.entries(countBy(recs, 'c'+catCol)).sort(function(a,b){return b[1]-a[1];}).slice(0,8);
      chartCards.push({id:'ch-c-'+idx, title:'Records by ' + esc(headers[catCol]), type:'doughnut',
        labels: tm.map(function(e){return e[0];}),
        data: tm.map(function(e){return e[1];})});
    }
    chartCards.forEach(function(cc){
      var isBar = cc.type==='bar';
      charts += '<div class="chart-card" style="flex:' + (isBar?1.7:1) + ';min-width:280px;"><h3>' + cc.title + '</h3>' +
        '<div style="height:300px;position:relative;"><canvas id="' + cc.id + '"></canvas></div></div>';
      var cols = PAL.slice(0, cc.labels.length);
      var ds = isBar ?
        [{label:cc.title.replace(/Top 15 by /,'').replace(/Records by /,''), data:cc.data, backgroundColor:'rgba(37,99,235,0.75)', borderColor:'#2563eb', borderWidth:1, borderRadius:5}] :
        [{data:cc.data, backgroundColor:cols, borderColor:'#fffcf7', borderWidth:2, hoverOffset:6}];
      initChart(idx, cc.id, cc.type, cc.labels, ds, isBar);
    });
    if(chartCards.length) document.getElementById('pcharts-'+idx).innerHTML = charts;
    renderTable(idx, rows, headers, recs);
  }

  // ---- Dashboard sheet: KPI blocks + table blocks ----
  function buildDashboardPage(idx, rows){
    pagesCache[idx] = pagesCache[idx] || {filter:''};
    var kpis = [];
    var blocks = [];
    for(var i=0;i<rows.length;i++){
      var first = String(rows[i][0]||'').trim();
      if(/^metric$/i.test(first)){
        for(var j=i+1;j<rows.length;j++){
          var r = rows[j];
          var lbl = String(r[0]||'').trim();
          if(!lbl || r.length<2 || /^[^A-Za-z]/.test(lbl)) break;
          kpis.push({lbl:lbl, done:parseNum(r[1]), total:parseNum(r[2])});
        }
      }
      if(/^(discipline|subsystem|inspection type)$/i.test(first)){
        var hdrRow = rows[i].filter(function(c){ return String(c).trim()!==''; });
        var bRows = [];
        for(var j2=i+1;j2<rows.length;j2++){
          var r2 = rows[j2];
          var f2 = String(r2[0]||'').trim();
          if(!f2 || /^total$/i.test(f2)) break;
          if(/^[A-Z]{2,}[\s]*$/.test(f2) && r2[1]===undefined) break;
          bRows.push(r2.slice(0,hdrRow.length));
        }
        blocks.push({headers:hdrRow, rows:bRows, title:first});
      }
    }
    var kpiHtml = kpis.map(function(k){
      var p = k.total ? Math.round(k.done/k.total*100) : 0;
      return '<div class="kpi"><div class="icon">' + (p>=80?'✅':p>=50?'📈':'⏳') + '</div>' +
        '<div class="val">' + fmt(k.done) + ' / ' + fmt(k.total) + '</div>' +
        '<div class="lbl">' + esc(k.lbl) + '</div>' + progressBar(p) + '</div>';
    }).join('');
    document.getElementById('pkpi-'+idx).innerHTML = '<div class="kpi-row">' + kpiHtml + '</div>';

    var charts = '';
    blocks.forEach(function(b, bi){
      var headers = b.headers;
      var recs = toRecords(b.rows, -1, headers);
      var col = pickCol(headers, ['discipline','subsystem','inspection type']);
      if(col<0) return;
      var cm = Object.entries(countBy(recs, 'c'+col)).sort(function(a,b){return b[1]-a[1];});
      if(!cm.length) return;
      var isSub = /subsystem/i.test(String(headers[col]));
      var slice = cm.slice(0, isSub?15:8);
      var totalAll = cm.reduce(function(a,e){return a+e[1];},0);
      var isBar = isSub;
      charts += '<div class="chart-card" style="flex:' + (isBar?1.7:1) + ';min-width:280px;"><h3>' + esc(b.title) + ' (by ' + esc(headers[col]) + ')</h3>' +
        '<div style="height:300px;position:relative;"><canvas id="ch-dd-' + idx + '-' + bi + '"></canvas></div></div>';
      var ds = isBar ?
        [{label:esc(headers[col]), data:slice.map(function(e){return e[1];}), backgroundColor:'rgba(200,148,10,0.75)', borderColor:'#c8940a', borderWidth:1, borderRadius:5}] :
        [{data:slice.map(function(e){return e[1];}), backgroundColor:PAL.slice(0,slice.length), borderColor:'#fffcf7', borderWidth:2, hoverOffset:6}];
      initChart(idx, 'ch-dd-'+idx+'-'+bi, isBar?'bar':'doughnut',
        slice.map(function(e){return e[0].length>28 ? e[0].slice(0,28)+'…' : e[0];}), ds, isBar);
    });
    if(charts) document.getElementById('pcharts-'+idx).innerHTML = charts;
    document.getElementById('count-'+idx).textContent = rows.length + ' rows';
    renderRawTable(idx, rows);
  }

  // ---- Subsystem Report: per-subsystem metric cards ----
  function buildSubsystemReportPage(idx, rows){
    pagesCache[idx] = pagesCache[idx] || {filter:''};
    var blocks = [];
    var cur = null;
    for(var i=0;i<rows.length;i++){
      var first = String(rows[i][0]||'').trim();
      if(first.indexOf('▸')===0){ cur = {name:first.replace(/^▸\s*/,''), metrics:[]}; blocks.push(cur); continue; }
      if(cur && /^metric$/i.test(first)){
        for(var j=i+1;j<rows.length;j++){
          var r = rows[j];
          var lbl = String(r[0]||'').trim();
          if(!lbl) break;
          cur.metrics.push({lbl:lbl, open:parseNum(r[8]), closed:parseNum(r[9])});
        }
      }
    }
    var kpiHtml = blocks.map(function(b){
      var tot=0, cl=0;
      b.metrics.forEach(function(m){ tot+=m.open+m.closed; cl+=m.closed; });
      var p = tot ? Math.round(cl/tot*100) : 0;
      return '<div class="kpi"><div class="icon">' + (p>=80?'✅':'🗂️') + '</div>' +
        '<div class="val" style="font-size:14px;">' + esc(b.name) + '</div>' +
        '<div class="lbl">' + cl + ' / ' + tot + ' closed · ' + p + '%</div>' + progressBar(p) + '</div>';
    }).join('');
    document.getElementById('pkpi-'+idx).innerHTML = '<div class="kpi-row">' + kpiHtml + '</div>';
    document.getElementById('count-'+idx).textContent = blocks.length + ' subsystems';
    renderRawTable(idx, rows);
  }

  // ---- route dispatch ----
  function buildPage(idx, name, rows){
    if(/^dashboard$/i.test(name)) return buildDashboardPage(idx, rows);
    if(/subsystem report/i.test(name)) return buildSubsystemReportPage(idx, rows);
    return buildTablePage(idx, name, rows);
  }

  order.forEach(function(name, idx){
    document.querySelector('[data-search="' + idx + '"]').addEventListener('input', function(e){
      if(!builtIdx['tab-'+idx]){ builtIdx['tab-'+idx] = true; buildPage(idx, name, pages[name] || []); }
      pagesCache[idx] = pagesCache[idx] || {};
      pagesCache[idx].filter = e.target.value;
      var pc = pagesCache[idx];
      if(pc.recs) renderTable(idx, pc.rows, pc.headers, pc.recs);
      else renderRawTable(idx, pages[order[idx]] || []);
    });
  });

  // Tab switching (lazy build: a page is built only the first time it is opened)
  function switchTab(id){
    document.querySelectorAll('.tabbtn').forEach(function(b){ b.classList.toggle('active', b.dataset.tab === id); });
    document.querySelectorAll('.tabpage').forEach(function(p){ p.classList.toggle('active', p.id === id); });
    var num = parseInt(id.split('-')[1], 10);
    if(!isNaN(num) && !builtIdx[id] && order[num]){
      builtIdx[id] = true;
      buildPage(num, order[num], pages[order[num]] || []);
    }
    (pageCharts[num]||[]).forEach(function(c){ try{ c.resize(); }catch(e){} });
  }
  document.querySelectorAll('.tabbtn').forEach(function(b){
    b.addEventListener('click', function(){ switchTab(this.dataset.tab); });
  });
  switchTab('tab-main-dashboard');
})();
</script>
</body>
</html>
"""


def build_html(itr_data, punch_data, rfi_data, search_index, eit_table_data, cmt_qc_punch_data, cable_ov_data, cable_tracker_data, output_path):
    html = HTML_TEMPLATE
    html = html.replace('__ITR_JSON__', json.dumps(itr_data, ensure_ascii=False))
    html = html.replace('__PUNCH_JSON__', json.dumps(punch_data, ensure_ascii=False) if punch_data else 'null')
    html = html.replace('__RFI_JSON__', json.dumps(rfi_data, ensure_ascii=False) if rfi_data else 'null')
    html = html.replace('__NOW__', itr_data['now'])
    html = html.replace('__TODAY_LABEL__', itr_data['today_label'])
    html = html.replace('__TOTAL_CLOSED__', str(itr_data['total_closed_project']))
    html = html.replace('__TOTAL_TASKS__', str(itr_data['total_project_tasks']))
    html = html.replace('__HOURLY_TOTAL__', str(itr_data['hourly_closed_eit']))
    html = html.replace('__HOURLY_SUBMITTED__', str(itr_data['hourly_submitted']))
    html = html.replace('__WEEKLY_TOTAL__', str(itr_data['weekly_total']))
    html = html.replace('__MONTHLY_TOTAL__', str(itr_data['monthly_total']))
    html = html.replace('__PUNCH_TOTAL__', str(punch_data['total']) if punch_data else '0')
    html = html.replace('__PUNCH_DAILY_TOTAL__', str(punch_data['daily_total']) if punch_data else '0')
    html = html.replace('__PUNCH_WEEKLY_TOTAL__', str(punch_data['weekly_total']) if punch_data else '0')
    html = html.replace('__PUNCH_MONTHLY_TOTAL__', str(punch_data['monthly_total']) if punch_data else '0')
    html = html.replace('__RFI_TOTAL__', str(rfi_data['total_rfi']) if rfi_data else '0')
    html = html.replace('__SEARCH_INDEX_JSON__', json.dumps(search_index, ensure_ascii=False))
    html = html.replace('__EIT_TABLE_JSON__', json.dumps(eit_table_data, ensure_ascii=False) if eit_table_data else 'null')
    html = html.replace('__EIT_DESC_JSON__', json.dumps(eit_table_data.get('eit_desc', {}), ensure_ascii=False) if eit_table_data else 'null')
    html = html.replace('__CMT_QC_PUNCH_JSON__', json.dumps(cmt_qc_punch_data, ensure_ascii=False) if cmt_qc_punch_data else 'null')
    html = html.replace('__CABLE_OV_JSON__', json.dumps(cable_ov_data, ensure_ascii=False) if cable_ov_data else 'null')
    html = html.replace('__CABLE_TRACKER_JSON__', json.dumps(cable_tracker_data, ensure_ascii=False) if cable_tracker_data else 'null')

    # ---- Extra pages from PS5 EIT CPP AGI Dashboard.xlsx ----
    eit_pages = {}
    try:
        eit_xlsx = os.path.join(DOWNLOADS, 'PS5 EIT CPP AGI Dashboard.xlsx')
        if os.path.exists(eit_xlsx):
            xls = pd.ExcelFile(eit_xlsx)
            skip_sheets = {'data - punch list', 'data - qc punch register', 'data - inspection', 'subsystem summary combined'}
            for sheet in xls.sheet_names:
                if sheet.strip().lower() in skip_sheets:
                    continue
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
                rows = []
                for _, row in df.iterrows():
                    vals = []
                    for v in row.tolist():
                        if v is None:
                            vals.append('')
                        elif isinstance(v, float) and v.is_integer():
                            vals.append(str(int(v)))
                        elif hasattr(v, 'date'):
                            vals.append(str(v.date()))
                        else:
                            s = str(v).strip()
                            vals.append(s if s != 'nan' else '')
                    while vals and vals[-1] == '':
                        vals.pop()
                    if any(vals):
                        rows.append(vals)
                eit_pages[sheet] = rows
            print(f"  Extra pages loaded: {len(eit_pages)} sheets from PS5 EIT CPP AGI Dashboard.xlsx")
    except Exception as e:
        print(f"  Extra pages skipped: {e}")
    html = html.replace('__EIT_PAGES_JSON__', json.dumps(eit_pages, ensure_ascii=False))

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Dashboard saved: {output_path}")


def build_cmt_qc_punch_data():
    """Read CMT_QC_Punch_Summary.xls and return parsed JSON data."""
    path = os.path.join(DOWNLOADS, 'CMT_QC_Punch_Summary.xls')
    if not os.path.exists(path):
        print('  CMT_QC_Punch_Summary.xls: NOT FOUND')
        return None
    try:
        from html.parser import HTMLParser
        class _CMTQC(HTMLParser):
            def __init__(self):
                super().__init__(); self.tables = []; self.in_t = False; self.in_r = False
                self.in_h = False; self.in_d = False; self.row = []; self.cell = ''; self.data = []; self.hdrs = []; self.hdr = True
            def handle_starttag(self, tag, attrs):
                if tag == 'table': self.in_t = True; self.data = []; self.hdrs = []; self.hdr = True
                elif tag == 'tr' and self.in_t: self.in_r = True; self.row = []
                elif tag in ('th','td') and self.in_r: self.in_h = (tag=='th'); self.in_d = (tag=='td'); self.cell = ''
                elif tag == 'br' and (self.in_d or self.in_h): self.cell += ' | '
            def handle_endtag(self, tag):
                if tag == 'table':
                    if self.data and self.hdrs: self.tables.append({'headers':self.hdrs,'rows':self.data})
                    self.in_t = False
                elif tag == 'tr' and self.in_r:
                    if self.hdr and self.row: self.hdrs = list(self.row); self.hdr = False
                    elif not self.hdr and self.row: self.data.append(list(self.row))
                    self.in_r = False
                elif tag in ('th','td'):
                    if self.in_h: self.hdrs.append(self.cell.strip()); self.in_h = False
                    elif self.in_d: self.row.append(self.cell.strip()); self.in_d = False
            def handle_data(self, data):
                if self.in_d or self.in_h: self.cell += data

        with open(path, 'r', encoding='utf-8-sig') as f:
            raw = f.read()
        parser = _CMTQC()
        parser.feed(raw)
        clean = []
        for t in parser.tables:
            rows = [[' '.join(c.replace('\n',' ').replace('\r',' ').strip().split()) for c in r] for r in t['rows']]
            clean.append({'headers': t['headers'], 'rows': rows})
        return clean
    except Exception as e:
        print(f'  CMT_QC_Punch_Summary.xls: Error - {e}')
        return None


# ====================================================================
#  MAIN
# ====================================================================# ====================================================================
#  MAIN
# ====================================================================
def main():
    import argparse
    parser = argparse.ArgumentParser(description='PS5 Project Dashboard Generator')
    parser.add_argument('--date', type=str, default=None,
                        help='Override today date (YYYY-MM-DD). Default: max date in data')
    args = parser.parse_args()

    today_override = None
    if args.date:
        try:
            datetime.datetime.strptime(args.date, '%Y-%m-%d')
            today_override = args.date
        except ValueError:
            print(f"[ERROR] Invalid date format: {args.date}. Use YYYY-MM-DD")
            return

    print("=" * 60)
    print("   PS5 Project Dashboard Generator")
    print("=" * 60)

    ov_path = find_file(['ovtasks'])
    punch_path = find_file(['punch', 'list', 'register']) or find_file(['punch', 'list'])
    rfi_path = find_file(['inspection', 'register'])

    if not ov_path:
        print(f"[ERROR] ovTasks file not found in {DOWNLOADS}")
        input("Press Enter...")
        return

    print(f"ovTasks file       : {ov_path}")
    print(f"Punch List         : {punch_path if punch_path else 'NOT FOUND'}")
    print(f"Inspection Register: {rfi_path if rfi_path else 'NOT FOUND'}")

    itr_data = build_itr_data(ov_path, today_override)
    eit_table_data = build_itr_breakdown_table(ov_path)
    punch_data = build_punch_data(punch_path) if punch_path else None
    rfi_data = build_inspection_data(rfi_path) if rfi_path else None
    cmt_qc_punch_data = build_cmt_qc_punch_data()
    cable_ov_data = build_cable_ov_data(ov_path)
    cable_tracker_data = build_cable_tracker_data()

    master_path = find_file(['master', 'tracker', 'eit']) or find_file(['PS5 Master tracker'])
    inspection_path = find_file(['inspection', 'register'])
    print("\nBuilding universal search index (Asset Tag -> ITR / RFI / Punch)...")
    search_index = build_search_index(ov_path, punch_path, rfi_path)

    build_html(itr_data, punch_data, rfi_data, search_index, eit_table_data, cmt_qc_punch_data, cable_ov_data, cable_tracker_data, OUTPUT_HTML)

    # نسخة تانية باسم PS5_Project_Dashboard.html (للمشاركة المباشرة)
    import shutil
    shutil.copy(OUTPUT_HTML, OUTPUT_HTML2)
    print(f"  Also saved: {OUTPUT_HTML2}")

    # ---- Save SMS data (fallback for phone) ----
    sms_out = os.path.join(os.path.dirname(__file__) or '.', 'sms_data.json')
    today = itr_data.get('today_label', datetime.datetime.now().strftime('%Y-%m-%d'))
    eit_summary = itr_data.get('eit_summary', [])
    sms_data = {
        'hourly_closed_eit': itr_data.get('hourly_closed_eit', 0),
        'today_submitted_assets': itr_data.get('today_submitted_assets', 0),
        'today_closed': itr_data.get('hourly_closed_eit', 0),
        'today_submitted': itr_data.get('today_submitted_assets', 0),
        'total_closed': itr_data.get('total_closed_eit', 0),
        'total_open': itr_data.get('total_tasks_eit', 0) - itr_data.get('total_closed_eit', 0),
        'eit_summary': eit_summary,
        'date': today,
    }
    for d in eit_summary:
        code = d['label'][0]
        sms_data[code.lower()] = {'total': d['total'], 'closed': d['closed'], 'open': d['total'] - d['closed']}
    with open(sms_out, 'w') as f:
        json.dump(sms_data, f)
    print(f"  SMS data saved: {sms_out}")

    # ---- Auto upload to GitHub via git (handles large files) ----
    token_file = os.path.join(os.path.dirname(__file__) or '.', 'github_token.txt')
    if os.path.exists(token_file):
        try:
            token = open(token_file, 'r').read().strip()
            if token:
                print("\n  Uploading to GitHub...")
                import subprocess
                base = os.path.dirname(os.path.abspath(__file__)) or '.'
                files_to_commit = ['index.html', 'sms_data.json', 'phone_sms.py', 'test_sms.py', 'cpp_agi_dashboard.py', 'update_dashboard.bat']
                # ensure local repo exists
                if not os.path.exists(os.path.join(base, '.git')):
                    subprocess.run(['git', 'init'], cwd=base, capture_output=True)
                    subprocess.run(['git', 'remote', 'add', 'origin', 'https://github.com/mohamedgawad1/ps5-dashboard.git'], cwd=base, capture_output=True)
                # set token in URL for this push only
                push_url = f"https://{token}@github.com/mohamedgawad1/ps5-dashboard.git"
                cmds = [
                    ['git', 'add'] + files_to_commit,
                    ['git', '-c', 'http.postBuffer=104857600', 'commit', '-m', f'auto update {datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}', '--no-verify'],
                    ['git', '-c', 'http.postBuffer=104857600', 'push', '--force', push_url, 'HEAD:main'],
                ]
                for cmd in cmds:
                    r = subprocess.run(cmd, cwd=base, capture_output=True, text=True, timeout=300)
                    err = (r.stderr or '').strip()
                    if r.returncode != 0 and 'nothing to commit' not in err and 'Everything up-to-date' not in err:
                        if cmd[0] == 'git' and cmd[1] == 'push':
                            print(f"  git push output: {err[-300:]}")
                            raise RuntimeError(err[-300:])
                print("  Uploaded via git: index.html, sms_data.json")
        except Exception as e:
            print(f"  Upload failed: {e}")

    print("\n" + "=" * 60)
    print("  Dashboard built successfully!")
    print(f"  File: {OUTPUT_HTML}")
    print("=" * 60)


if __name__ == "__main__":
    main()
