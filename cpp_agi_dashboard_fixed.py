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
import json
import webbrowser

try:
    import pandas as pd
except ImportError:
    os.system("pip install pandas openpyxl --break-system-packages")
    import pandas as pd


# ====================================================================
#  إعدادات
# ====================================================================
DOWNLOADS = "/mnt/user-data/uploads"
OUTPUT_HTML = "/mnt/user-data/outputs/index.html"
OUTPUT_HTML2 = "/mnt/user-data/outputs/PS5_Project_Dashboard.html"

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
            continue  # ملف مؤقت من Excel (lock file) - تجاهله
        low = f.lower()
        if low.endswith('.xlsx') and all(k.lower() in low for k in prefix_keywords):
            return os.path.join(DOWNLOADS, f)
    return None


# ====================================================================
#  2) ITR Closures من ovTasks
# ====================================================================
def build_itr_data(excel_path):
    df = pd.read_excel(excel_path, sheet_name='Exported from SC')

    total_project_tasks = len(df)
    closed = df[df['Task State'] == 'Closed'].copy()
    closed['Closing Date'] = pd.to_datetime(closed['Closing Date'])
    closed = closed.dropna(subset=['Closing Date'])

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

    now = closed['Closing Date'].max()
    total_closed_project = len(closed)

    def pivot_disc(data, col):
        if data.empty:
            return []
        p = data.groupby([col, 'disc']).size().unstack(fill_value=0)
        for d in ['E', 'I', 'T']:
            if d not in p.columns:
                p[d] = 0
        p['Total'] = p[['E', 'I', 'T']].sum(axis=1) + p.get('Other', 0)
        p = p[['E', 'I', 'T', 'Total']]
        return p.reset_index().rename(columns={col: 'label'}).to_dict('records')

    # ---- Hourly (Today) ----
    today_df = closed[closed['Closing Date'].dt.date == now.date()].copy()
    today_df['hour'] = today_df['Closing Date'].dt.hour.apply(lambda h: f"{h:02d}:00")
    hourly = pivot_disc(today_df, 'hour')
    hourly_total = int(today_df.shape[0])

    # ---- Daily (Last 30 Days) ----
    closed['date'] = closed['Closing Date'].dt.strftime('%Y-%m-%d')
    daily = pivot_disc(closed, 'date')
    daily = sorted(daily, key=lambda r: r['label'])[-30:]
    daily_total = hourly_total

    # ---- Weekly ----
    closed['week'] = closed['Closing Date'].dt.to_period('W').apply(
        lambda r: r.start_time.strftime('%Y-%m-%d'))
    weekly = pivot_disc(closed, 'week')
    weekly = sorted(weekly, key=lambda r: r['label'])[-16:]

    week_start = (now - pd.Timedelta(days=now.weekday())).normalize()
    weekly_total = int((closed['Closing Date'] >= week_start).sum())

    # ---- Monthly ----
    closed['month'] = closed['Closing Date'].dt.strftime('%Y-%m')
    monthly = pivot_disc(closed, 'month')
    monthly = sorted(monthly, key=lambda r: r['label'])

    monthly_total = int(((closed['Closing Date'].dt.year == now.year) &
                          (closed['Closing Date'].dt.month == now.month)).sum())


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

    return {
        'now': now.strftime('%Y-%m-%d %H:%M'),
        'today_label': now.strftime('%Y-%m-%d'),
        'hourly': hourly,
        'hourly_total': hourly_total,
        'daily': daily,
        'daily_total': daily_total,
        'weekly': weekly,
        'weekly_total': weekly_total,
        'monthly': monthly,
        'monthly_total': monthly_total,
        'total_closed_project': total_closed_project,
        'total_project_tasks': total_project_tasks,
        'eit_summary': eit_summary,
        'milestone_summary': milestone_summary,
    }


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
    df['status'] = df['status'].fillna('Open')

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

    # disc_counts بالأسماء الكاملة الموحدة (Electrical (E) / Instrumentation (I) / Telecom (T))
    DISC_FULL_LABEL = {'E': 'Electrical (E)', 'I': 'Instrumentation (I)', 'T': 'Telecom (T)', 'Other': 'Other'}
    disc_counts = (df['disc'].map(DISC_FULL_LABEL).fillna('Other')
                   .value_counts().to_dict())

    def pivot_disc(data, col):
        if data.empty:
            return []
        p = data.groupby([col, 'disc']).size().unstack(fill_value=0)
        for d in ['E', 'I', 'T']:
            if d not in p.columns:
                p[d] = 0
        p['Total'] = p[['E', 'I', 'T']].sum(axis=1) + p.get('Other', 0)
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

    df = pd.read_excel(excel_path, sheet_name='PS-5 EIT INSPECTION REGISTER', header=4)
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

    df['status_norm'] = df['Status Of RFI '].apply(norm_status)

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

    type_col = 'INSPECTION TYPE (PULLING/GLANDING & TERMINATION/CABLE TESTING/INSTALLATION)'
    df['type_norm'] = df[type_col].apply(norm_type)

    total_assets = len(df)
    total_rfi_submitted = int(df['QC RFI No'].notna().sum())

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
        recent_records.append({
            'asset': str(r.get(table_col, '')),
            'rfi_no': str(r.get('QC RFI No', '') or ''),
            'disc': r.get('disc', ''),
            'status': r.get('status_norm', ''),
            'date': r['Inspection Date'].strftime('%Y-%m-%d'),
        })

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
    }


# ====================================================================
#  3c) DPR Summary File (EIT SUMMERY sheet)
# ====================================================================
def build_eit_subsystem_data(excel_path):
    if not excel_path:
        return None

    df = pd.read_excel(excel_path, sheet_name='EIT SUMMERY', header=None, skiprows=3)
    df = df.iloc[:, :10]
    df.columns = ['subsystem', 'E_itrs', 'E_closed', 'E_remain',
                   'I_itrs', 'I_closed', 'I_remain',
                   'T_itrs', 'T_closed', 'T_remain']
    df = df.dropna(subset=['subsystem'])
    df = df[df['subsystem'].astype(str).str.startswith('PS5')]
    for c in df.columns[1:]:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(int)

    totals = {
        'E': {'itrs': int(df['E_itrs'].sum()), 'closed': int(df['E_closed'].sum())},
        'I': {'itrs': int(df['I_itrs'].sum()), 'closed': int(df['I_closed'].sum())},
        'T': {'itrs': int(df['T_itrs'].sum()), 'closed': int(df['T_closed'].sum())},
    }
    for d in totals.values():
        d['remain'] = d['itrs'] - d['closed']
        d['pct'] = round(d['closed'] / d['itrs'] * 100, 1) if d['itrs'] else 0

    # Top 15 subsystems by total ITRs remaining (biggest remaining work)
    df['total_remain'] = df['E_remain'] + df['I_remain'] + df['T_remain']
    df['total_itrs'] = df['E_itrs'] + df['I_itrs'] + df['T_itrs']
    top = df.sort_values('total_itrs', ascending=False).head(15)

    table_rows = []
    for _, r in top.iterrows():
        table_rows.append({
            'subsystem': r['subsystem'],
            'E_itrs': int(r['E_itrs']), 'E_closed': int(r['E_closed']), 'E_remain': int(r['E_remain']),
            'I_itrs': int(r['I_itrs']), 'I_closed': int(r['I_closed']), 'I_remain': int(r['I_remain']),
            'T_itrs': int(r['T_itrs']), 'T_closed': int(r['T_closed']), 'T_remain': int(r['T_remain']),
        })

    return {'totals': totals, 'table': table_rows}


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
    df = pd.read_excel(ov_path, sheet_name='Exported from SC')

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
            'closed': is_closed,          # ✅ هل مغلق فعلاً؟ من ovTasks مباشرة
            'ty':    str(r['Task Type (Name)'])   if pd.notna(r['Task Type (Name)'])   else '',
            'd':     disc_norm_map.get(raw_disc, raw_disc[:1] if raw_disc else ''),
            'disc':  disc_norm_map.get(raw_disc, raw_disc),
            'cd':    pd.to_datetime(closing).strftime('%Y-%m-%d') if pd.notna(closing) else '',
            'sub':   str(r['Systemization - Subsystem (Summary)']) if pd.notna(r.get('Systemization - Subsystem (Summary)')) else '',
            'ms':    str(r['Subsystem Priority']) if pd.notna(r.get('Subsystem Priority')) else '',
            'res':   str(r['Responsible Company (Summary)']) if pd.notna(r.get('Responsible Company (Summary)')) else '',
        })

    # ---- Inspection Register ----
    if rfi_path:
        rdf = pd.read_excel(rfi_path, sheet_name='PS-5 EIT INSPECTION REGISTER', header=4)
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

        type_col = 'INSPECTION TYPE (PULLING/GLANDING & TERMINATION/CABLE TESTING/INSTALLATION)'
        for _, r in rdf.iterrows():
            tag = r['Asset - Tag']
            insp_date = pd.to_datetime(r.get('Inspection Date'), errors='coerce')
            add(tag, 'rfi', {
                'rfi_no': str(r.get('QC RFI No', '') or ''),
                'status': norm_status(r.get('Status Of RFI ')),
                'discipline': str(r.get('Discipline', '') or ''),
                'type': str(r.get(type_col, '') or ''),
                'date': insp_date.strftime('%Y-%m-%d') if pd.notna(insp_date) else '',
            })

    # ---- Punch List ----
    if punch_path:
        # خريطة سريعة: RFI No -> [Asset Tags] من فهرس RFI المبني مسبقاً
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
            }
            tags = rfi_to_tags.get(rfi_no)
            if tags:
                for tag in tags:
                    index[tag]['punch'].append(record)
            else:
                add(f"[Unlinked] {rfi_no or record['plid']}", 'punch', record)

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
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en" dir="ltr">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="900">
<title>PS5 Project Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
:root{
  --bgmain:#0b1120; --panel:#131a2b; --panel2:#1a2236;
  --accent:#ff4d8d; --gold:#ffb627; --teal:#00e0c6; --blue:#3da8ff; --purple:#9b59f6;
  --text:#e8ecf7; --muted:#8a93ad; --border:#232c44;
}
*{box-sizing:border-box;font-family:'Segoe UI',Tahoma,Arial,sans-serif;margin:0;padding:0;}
body{
  background:radial-gradient(circle at 20% 0%, #16203a 0%, #0b1120 55%);
  color:var(--text);display:flex;
}

/* ---------- Sidebar ---------- */
.sidebar{
  width:240px;min-height:100vh;
  background:#0e1626;
  color:#fff;padding:24px 18px;position:sticky;top:0;height:100vh;overflow-y:auto;
  border-right:1px solid var(--border);
}
.sidebar h2{font-size:17px;margin-bottom:18px;display:flex;align-items:center;gap:8px;
  letter-spacing:.5px;color:var(--teal);}
.sidebar label{
  display:flex;align-items:center;gap:8px;padding:9px 8px;border-radius:8px;
  cursor:pointer;font-size:13px;margin-bottom:3px;transition:.15s;color:var(--muted);
}
.sidebar label:hover{background:rgba(255,255,255,.06);color:#fff;}
.sidebar input{accent-color:var(--teal);width:16px;height:16px;}
.sidebar .grp{margin-top:20px;font-size:11px;opacity:.6;text-transform:uppercase;
  letter-spacing:1.5px;border-bottom:1px solid var(--border);padding-bottom:4px;color:var(--gold);}

/* ---------- Main ---------- */
.main{flex:1;padding:24px 32px;}
.header{
  background:linear-gradient(120deg,#161f38 0%, #0e1626 100%);
  color:#fff;padding:22px 30px;border-radius:16px;margin-bottom:24px;
  display:flex;justify-content:space-between;align-items:center;
  border:1px solid var(--border);position:relative;overflow:hidden;
}
.header::after{
  content:"";position:absolute;inset:0;
  background:radial-gradient(circle at 90% 10%, rgba(0,224,198,.15), transparent 60%);
}
.header h1{font-size:24px;letter-spacing:.5px;}
.header .sub{font-size:12px;opacity:.7;margin-top:4px;color:var(--muted);}

.section{display:none;}
.section.active{display:block;}

.kpi-row{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:22px;}
.kpi{
  flex:1;min-width:170px;background:var(--panel);border-radius:14px;padding:20px;
  text-align:center;border:1px solid var(--border);
  border-top:4px solid var(--purple);position:relative;overflow:hidden;
  transition:transform .2s, border-color .2s;
}
.kpi:hover{transform:translateY(-4px);border-color:var(--teal);}
.kpi.gold{border-top-color:var(--gold);}
.kpi.teal{border-top-color:var(--teal);}
.kpi.pink{border-top-color:var(--accent);}
.kpi.blue{border-top-color:var(--blue);}
.kpi .icon{font-size:24px;margin-bottom:6px;}
.kpi .val{font-size:32px;font-weight:800;color:#fff;letter-spacing:.5px;}
.kpi .lbl{font-size:12px;color:var(--muted);margin-top:4px;}

.chart-row{display:flex;gap:18px;flex-wrap:wrap;margin-bottom:22px;}
.chart-card{
  flex:1;min-width:340px;background:var(--panel);border-radius:14px;padding:20px;
  border:1px solid var(--border);
}
.chart-card h3{font-size:15px;color:var(--teal);margin-bottom:12px;font-weight:700;}
canvas{max-height:300px;}

.section-title{
  font-size:19px;color:#fff;font-weight:800;margin:8px 0 16px;
  border-left:6px solid var(--accent);padding-left:12px;
  display:flex;align-items:center;gap:8px;
}
.progress-bar{height:10px;background:#0e1626;border-radius:6px;overflow:hidden;margin-top:8px;border:1px solid var(--border);}
.progress-fill{height:100%;background:linear-gradient(90deg,var(--blue),var(--teal));}
.footer{text-align:center;padding:16px;color:var(--muted);font-size:12px;}

/* tables on dark theme */
table{color:var(--text);}
thead tr{background:var(--panel2) !important;}
tbody tr{border-bottom:1px solid var(--border) !important;}
input#rfiSearch{background:var(--panel2);border:1px solid var(--border) !important;color:#fff;}
input#rfiSearch::placeholder{color:var(--muted);}
</style>
</head>
<body>

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

  <div class="grp">Punch List</div>
  <label><input type="checkbox" data-target="sec-punch-kpi" checked> Punch List Summary</label>
  <label><input type="checkbox" data-target="sec-punch-daily" checked> Daily</label>
  <label><input type="checkbox" data-target="sec-punch-weekly" checked> Weekly</label>
  <label><input type="checkbox" data-target="sec-punch-monthly" checked> Monthly</label>
  <label><input type="checkbox" data-target="sec-punch-breakdown" checked> Breakdown (Category/Discipline)</label>
  <label><input type="checkbox" data-target="sec-punch-subsystems" checked> Top Subsystems</label>
  <label><input type="checkbox" data-target="sec-punch-recent" checked> Recent Punches (Details)</label>

  <div class="grp">RFI / Inspection Register</div>
  <label><input type="checkbox" data-target="sec-rfi-kpi" checked> RFI Summary</label>
  <label><input type="checkbox" data-target="sec-rfi-status" checked> Status Breakdown</label>
  <label><input type="checkbox" data-target="sec-rfi-daily" checked> Daily</label>
  <label><input type="checkbox" data-target="sec-rfi-weekly" checked> Weekly</label>
  <label><input type="checkbox" data-target="sec-rfi-monthly" checked> Monthly</label>
  <label><input type="checkbox" data-target="sec-rfi-type" checked> Inspection Type</label>
  <label><input type="checkbox" data-target="sec-rfi-subsystems" checked> Top Subsystems</label>
  <label><input type="checkbox" data-target="sec-rfi-recent" checked> Recent RFIs (Details)</label>

  <div class="grp">DPR Summary (EIT)</div>
  <label><input type="checkbox" data-target="sec-dpr-eit" checked> E/I/T Progress &amp; Top Subsystems</label>
  <label><input type="checkbox" data-target="sec-punch-tracking" checked> Punch Tracking &amp; Closure</label>
</div>

<!-- ================= MAIN ================= -->
<div class="main">
  <div class="header">
    <div>
      <h1>📊 PS5 — EACOP Project Dashboard</h1>
      <div class="sub">ITR Closures & Punch List | Data last updated: __NOW__</div>
    </div>
    <div style="font-size:38px;">⚙️</div>
  </div>

  <!-- ===== Universal Search ===== -->
  <div class="section active" id="sec-search">
    <div class="section-title">🔍 Universal Asset / Task ID Search — ITR + RFI + Punch (Combined, No Conflicts)</div>
    <div class="chart-card">
      <input id="universalSearch" type="text" placeholder="Type Asset Tag or Task ID (e.g. PS5-25-DM-3100A-CH01 or T-00232-0952)..."
        style="width:100%;padding:12px 16px;border:1px solid var(--border);border-radius:8px;font-size:14px;
        background:var(--panel2);color:#fff;margin-bottom:6px;">
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
      <div class="kpi"><div class="icon">✅</div><div class="val">__TOTAL_CLOSED__</div><div class="lbl">Total ITRs Closed (Whole Project)</div></div>
      <div class="kpi gold"><div class="icon">📋</div><div class="val">__TOTAL_TASKS__</div><div class="lbl">Total Project Tasks</div></div>
      <div class="kpi pink"><div class="icon">📌</div><div class="val">__PUNCH_TOTAL__</div><div class="lbl">Total Punch List Items</div></div>
      <div class="kpi blue"><div class="icon">📝</div><div class="val">__RFI_TOTAL__</div><div class="lbl">Total RFIs Submitted</div></div>
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

  <!-- ===== DPR: EIT ITRs Backlog Tracking ===== -->
  <div class="section active" id="sec-dpr-eit">
    <div class="section-title">📊 EIT ITRs Backlog Tracking (Project-wide)</div>
    <div class="kpi-row" id="dprEitKpis"></div>
    <div class="chart-row">
      <div class="chart-card" style="flex:2;">
        <h3>Top 15 Subsystems — Total ITRs (E+I+T)</h3>
        <canvas id="chartDprEitTop"></canvas>
      </div>
    </div>
    <div class="chart-card">
      <h3>Backlog Table — Top 15 Subsystems</h3>
      <div style="max-height:420px;overflow:auto;">
        <table style="width:100%;border-collapse:collapse;font-size:12px;">
          <thead>
            <tr style="background:var(--panel2);color:var(--teal);">
              <th style="padding:6px;text-align:left;">Subsystem</th>
              <th style="padding:6px;">E Total</th><th style="padding:6px;">E Closed</th><th style="padding:6px;">E Remain</th>
              <th style="padding:6px;">I Total</th><th style="padding:6px;">I Closed</th><th style="padding:6px;">I Remain</th>
              <th style="padding:6px;">T Total</th><th style="padding:6px;">T Closed</th><th style="padding:6px;">T Remain</th>
            </tr>
          </thead>
          <tbody id="dprEitTableBody"></tbody>
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



<script>
const ITR = __ITR_JSON__;
const PUNCH = __PUNCH_JSON__;
const RFI = __RFI_JSON__;
const DPR_EIT = __DPR_EIT_JSON__;
const SEARCH_INDEX = __SEARCH_INDEX_JSON__;
const SI       = SEARCH_INDEX.index   || SEARCH_INDEX;
const RFI_MAP  = SEARCH_INDEX.rfi_map || {};
const SUB_MAP  = SEARCH_INDEX.sub_map || {};
const PRI_MAP  = SEARCH_INDEX.pri_map || {};
const TID_MAP  = SEARCH_INDEX.tid_map || {};
const palette = ['#9b59f6','#3da8ff','#ff4d8d','#ffb627','#00e0c6','#7c8cff','#ff8a3d'];
const MS_STATUS_COLORS = {'Closed':'#00e0c6','Submitted':'#3da8ff','To be completed':'#ffb627','Other':'#8a93ad'};
const MS_STATUS_ORDER = ['Closed','Submitted','To be completed','Other'];

// ---------- Global Chart.js + datalabels setup ----------
Chart.register(ChartDataLabels);
Chart.defaults.font.family = "'Segoe UI', Tahoma, Arial, sans-serif";
Chart.defaults.color = '#8a93ad';
Chart.defaults.borderColor = '#232c44';
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
    {key:'E', name:'Electrical (E)', color:'#9b59f6'},
    {key:'I', name:'Instrumentation (I)', color:'#00e0c6'},
    {key:'T', name:'Telecom (T)', color:'#ff4d8d'},
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
    {name:'ITR Closed', color:'#9b59f6', map:itrMap},
    {name:'Punch Raised', color:'#ff4d8d', map:punchMap},
    {name:'RFI Submitted', color:'#3da8ff', map:rfiMap},
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
        {label:'Closed', data: ITR.eit_summary.map(s=>s.closed), backgroundColor:'#00e0c6', datalabels: DL_BAR},
        {label:'Total', data: ITR.eit_summary.map(s=>s.total), backgroundColor:'#2a3450', datalabels: {display:false}}
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
      datasets:[{ label:'Punch Items', data: PUNCH.top_subsystems.map(r=>r.count), backgroundColor:'#6a2c91', borderRadius:6, datalabels: {...DL_BAR, anchor:'end', align:'right'} }]},
    options:{ indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{beginAtZero:true}} }
  });
} else {
  ['sec-punch-kpi','sec-punch-daily','sec-punch-weekly','sec-punch-monthly','sec-punch-breakdown','sec-punch-subsystems']
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
      datasets:[{ label:'RFI Count', data: RFI.top_subsystems.map(r=>r.count), backgroundColor:'#2980ff', borderRadius:6, datalabels: {...DL_BAR, anchor:'end', align:'right'} }]},
    options:{ indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{beginAtZero:true}} }
  });

  // ---- Recent RFI table with search ----
  function renderRfiTable(rows){
    const statusColor = {
      'Accepted':'#16a085', 'Accepted with Punch':'#f4b942',
      'Open':'#ff4d8d', 'Hold by EACOP':'#2980ff', 'No RFI Yet':'#999', 'Other':'#999'
    };
    document.getElementById('rfiTableBody').innerHTML = rows.map(r=>`
      <tr style="border-bottom:1px solid #eee;">
        <td style="padding:7px;">${r.asset}</td>
        <td style="padding:7px;">${r.rfi_no||'-'}</td>
        <td style="padding:7px;">${r.disc}</td>
        <td style="padding:7px;color:${statusColor[r.status]||'#333'};font-weight:600;">${r.status}</td>
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
  ['sec-rfi-kpi','sec-rfi-status','sec-rfi-daily','sec-rfi-weekly','sec-rfi-monthly','sec-rfi-type','sec-rfi-subsystems','sec-rfi-recent']
    .forEach(id=>{
      document.getElementById(id).innerHTML = '<div class="section-title">⚠️ Inspection Register file not found</div>';
    });
}

// ---------- DPR: EIT ITRs Backlog Tracking ----------
if(DPR_EIT){
  (function(){
    const c = document.getElementById('dprEitKpis');
    const map = {E:{name:'Electrical (E)',color:'gold'}, I:{name:'Instrumentation (I)',color:'teal'}, T:{name:'Telecom (T)',color:'pink'}};
    Object.entries(DPR_EIT.totals).forEach(([k,v])=>{
      const div = document.createElement('div');
      div.className = 'kpi ' + map[k].color;
      div.innerHTML = `<div class="val">${v.closed} / ${v.itrs}</div>
        <div class="lbl">${map[k].name} — ${v.pct}% Closed (Remain: ${v.remain})</div>
        <div class="progress-bar"><div class="progress-fill" style="width:${v.pct}%"></div></div>`;
      c.appendChild(div);
    });
  })();

  new Chart(document.getElementById('chartDprEitTop'), {
    type:'bar',
    data:{
      labels: DPR_EIT.table.map(r=>r.subsystem),
      datasets:[
        {label:'Electrical', data: DPR_EIT.table.map(r=>r.E_itrs), backgroundColor:'#ffb627'},
        {label:'Instrumentation', data: DPR_EIT.table.map(r=>r.I_itrs), backgroundColor:'#00e0c6'},
        {label:'Telecom', data: DPR_EIT.table.map(r=>r.T_itrs), backgroundColor:'#ff4d8d'},
      ]
    },
    options:{ indexAxis:'y', plugins:{legend:{position:'bottom'}}, scales:{x:{stacked:true, beginAtZero:true}, y:{stacked:true}} }
  });

  document.getElementById('dprEitTableBody').innerHTML = DPR_EIT.table.map(r=>`
    <tr style="border-bottom:1px solid var(--border);">
      <td style="padding:6px;">${r.subsystem}</td>
      <td style="padding:6px;text-align:center;">${r.E_itrs}</td><td style="padding:6px;text-align:center;">${r.E_closed}</td><td style="padding:6px;text-align:center;color:var(--accent);">${r.E_remain}</td>
      <td style="padding:6px;text-align:center;">${r.I_itrs}</td><td style="padding:6px;text-align:center;">${r.I_closed}</td><td style="padding:6px;text-align:center;color:var(--accent);">${r.I_remain}</td>
      <td style="padding:6px;text-align:center;">${r.T_itrs}</td><td style="padding:6px;text-align:center;">${r.T_closed}</td><td style="padding:6px;text-align:center;color:var(--accent);">${r.T_remain}</td>
    </tr>`).join('');
} else {
  document.getElementById('sec-dpr-eit').innerHTML =
    "<div class=\"section-title\">⚠️ DPR Summary file not found (needs sheet 'EIT SUMMERY')</div>";
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
      datasets:[{ label:'Open Punch Items', data: PUNCH.top_subsystems.map(r=>r.count), backgroundColor:'#ff4d8d', borderRadius:6, datalabels:{...DL_BAR, anchor:'end', align:'right'} }]},
    options:{ indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{beginAtZero:true}} }
  });

  // Punch by discipline - open vs closed (closed approximated as 0 since data has no closed yet)
  const discLabels = Object.keys(PUNCH.disc_counts);
  new Chart(document.getElementById('chartPunchClosure'), {
    type:'bar',
    data:{
      labels: discLabels,
      datasets:[
        {label:'Open', data: discLabels.map(k=>PUNCH.disc_counts[k]), backgroundColor:'#ff4d8d', datalabels: DL_BAR},
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
  'Closed':'#00e0c6','To be completed':'#ffb627','Submitted':'#3da8ff',
  'Accepted':'#00e0c6','Accepted with Punch':'#ffb627','Open':'#ff4d8d',
  'Hold by EACOP':'#3da8ff','No RFI Yet':'#8a93ad','Other':'#8a93ad','Rejected':'#ff4d8d',
};
const discColors = {'E':'#ffb627','I':'#00e0c6','T':'#ff4d8d'};
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
        <div style="font-size:13px;color:#fff;font-weight:700;">${s.closed} / ${s.total} Closed — ${s.pct}%</div>
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
    const discColor = discColors[first.d] || '#8a93ad';

    const itrRows = srcs.itr.map(it=>`
      <tr style="border-bottom:1px solid var(--border);background:${it.closed?'rgba(0,230,118,.04)':'rgba(255,77,141,.03)'}">
        <td style="padding:5px;color:var(--gold);font-weight:700;">${it.id||'-'}</td>
        <td style="padding:5px;">${it.ty||'-'}</td>
        <td style="padding:5px;">${badge(it.disc||it.d||'-', discColors[it.d]||'#8a93ad')}</td>
        <td style="padding:5px;">
          ${it.closed
            ? '<span style="color:#00e676;font-weight:800;">✅ Closed</span>'
            : `<span style="color:#ff4d8d;font-weight:700;">⏳ ${it.st||'Pending'}</span>`}
        </td>
        <td style="padding:5px;color:${it.cd?'#00e0c6':'var(--muted)'};">${it.cd||'-'}</td>
        <td style="padding:5px;color:var(--muted);font-size:11px;">${(it.sub||'').split(' - ')[0]||'-'}</td>
        <td style="padding:5px;">${it.ms?badge('🎯 '+it.ms,'#ffb627'):'-'}</td>
      </tr>`).join('');

    const rfiRows = srcs.rfi.map(r=>`
      <tr style="border-bottom:1px solid var(--border);">
        <td style="padding:5px;color:var(--blue);font-weight:700;">${r.rfi_no||'-'}</td>
        <td style="padding:5px;">${r.discipline||'-'}</td>
        <td style="padding:5px;">${(r.type||'-').slice(0,20)}</td>
        <td style="padding:5px;">${badge(r.status, statusColors[r.status]||'#8a93ad')}</td>
        <td style="padding:5px;">${r.date||'-'}</td>
      </tr>`).join('');

    const punchRows = srcs.punch.map(p=>`
      <tr style="border-bottom:1px solid var(--border);">
        <td style="padding:5px;color:var(--accent);font-weight:700;">${p.plid||'-'}</td>
        <td style="padding:5px;">${p.discipline||'-'}</td>
        <td style="padding:5px;">${p.category||'-'}</td>
        <td style="padding:5px;">${badge(p.status, statusColors[p.status]||'#ff4d8d')}</td>
        <td style="padding:5px;">${p.rfi_no||'-'}</td>
        <td style="padding:5px;">${p.date||'-'}</td>
      </tr>`).join('');

    return `
    <div style="border:1px solid var(--border);border-radius:12px;padding:16px;margin-bottom:14px;background:var(--panel2);">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
        <div>
          <div style="font-size:17px;font-weight:800;color:var(--teal);">📍 ${tag}</div>
          <div style="font-size:12px;margin-top:4px;">
            ${badge(first.disc||first.d||'Unknown', discColor)}
            <span style="color:#00e676;font-weight:700;margin-left:6px;">✅ ${closedTasks} Closed</span>
            ${pendingTasks>0?`<span style="color:#ff4d8d;font-weight:700;margin-left:4px;">⏳ ${pendingTasks} Pending</span>`:''}
            ${first.ms ? badge('🎯 '+first.ms,'#ffb627') : ''}
            ${first.res ? `<span style="color:var(--muted);font-size:11px;margin-left:6px;">${first.res}</span>` : ''}
          </div>
        </div>
        <div style="text-align:right;font-size:11px;color:var(--muted);">
          ${first.sub ? `<div style="color:var(--muted)">${first.sub}</div>` : ''}
          ${first.cd ? `<div style="color:var(--teal)">Closed: ${first.cd}</div>` : ''}
        </div>
      </div>

      <!-- ITR -->
      <div style="margin-bottom:12px;">
        <div style="font-weight:700;color:#fff;font-size:13px;margin-bottom:6px;">
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
        <div style="font-weight:700;color:#fff;font-size:13px;margin-bottom:6px;">
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
        <div style="font-weight:700;color:#fff;font-size:13px;margin-bottom:6px;">
          📌 Punch List <span style="color:var(--accent)">(${srcs.punch.length})</span></div>
        ${srcs.punch.length ? `<div style="overflow-x:auto;">
          <table style="width:100%;font-size:12px;border-collapse:collapse;min-width:500px;">
            <tr style="color:var(--teal);border-bottom:1px solid var(--border);">
              <th style="text-align:left;padding:5px;">PL ID</th>
              <th style="padding:5px;">Discipline</th>
              <th style="padding:5px;">Category</th>
              <th style="padding:5px;">Status</th>
              <th style="padding:5px;">RFI No</th>
              <th style="padding:5px;">Raised Date</th>
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
</script>
</body>
</html>
"""


def build_html(itr_data, punch_data, rfi_data, dpr_eit_data, search_index, output_path):
    html = HTML_TEMPLATE
    html = html.replace('__ITR_JSON__', json.dumps(itr_data, ensure_ascii=False))
    html = html.replace('__PUNCH_JSON__', json.dumps(punch_data, ensure_ascii=False) if punch_data else 'null')
    html = html.replace('__RFI_JSON__', json.dumps(rfi_data, ensure_ascii=False) if rfi_data else 'null')
    html = html.replace('__NOW__', itr_data['now'])
    html = html.replace('__TODAY_LABEL__', itr_data['today_label'])
    html = html.replace('__TOTAL_CLOSED__', str(itr_data['total_closed_project']))
    html = html.replace('__TOTAL_TASKS__', str(itr_data['total_project_tasks']))
    html = html.replace('__HOURLY_TOTAL__', str(itr_data['hourly_total']))
    html = html.replace('__WEEKLY_TOTAL__', str(itr_data['weekly_total']))
    html = html.replace('__MONTHLY_TOTAL__', str(itr_data['monthly_total']))
    html = html.replace('__PUNCH_TOTAL__', str(punch_data['total']) if punch_data else '0')
    html = html.replace('__PUNCH_DAILY_TOTAL__', str(punch_data['daily_total']) if punch_data else '0')
    html = html.replace('__PUNCH_WEEKLY_TOTAL__', str(punch_data['weekly_total']) if punch_data else '0')
    html = html.replace('__PUNCH_MONTHLY_TOTAL__', str(punch_data['monthly_total']) if punch_data else '0')
    html = html.replace('__RFI_TOTAL__', str(rfi_data['total_rfi']) if rfi_data else '0')
    html = html.replace('__DPR_EIT_JSON__', json.dumps(dpr_eit_data, ensure_ascii=False) if dpr_eit_data else 'null')
    html = html.replace('__SEARCH_INDEX_JSON__', json.dumps(search_index, ensure_ascii=False))

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Dashboard saved: {output_path}")


# ====================================================================
#  MAIN
# ====================================================================
def main():
    print("=" * 60)
    print("   PS5 Project Dashboard Generator")
    print("=" * 60)

    ov_path = find_file(['ovtasks'])
    punch_path = find_file(['punch', 'list'])
    rfi_path = find_file(['inspection', 'register'])
    dpr_path = find_file(['completions', 'dpr'])

    if not ov_path:
        print(f"[ERROR] ovTasks file not found in {DOWNLOADS}")
        input("Press Enter...")
        return

    print(f"ovTasks file       : {ov_path}")
    print(f"Punch List         : {punch_path if punch_path else 'NOT FOUND'}")
    print(f"Inspection Register: {rfi_path if rfi_path else 'NOT FOUND'}")
    print(f"DPR Summary file   : {dpr_path if dpr_path else 'NOT FOUND'}")

    itr_data = build_itr_data(ov_path)
    punch_data = build_punch_data(punch_path) if punch_path else None
    rfi_data = build_inspection_data(rfi_path) if rfi_path else None
    dpr_eit_data = build_eit_subsystem_data(dpr_path) if dpr_path else None

    print("\nBuilding universal search index (Asset Tag -> ITR / RFI / Punch)...")
    search_index = build_search_index(ov_path, punch_path, rfi_path)

    build_html(itr_data, punch_data, rfi_data, dpr_eit_data, search_index, OUTPUT_HTML)

    # نسخة تانية باسم PS5_Project_Dashboard.html (للمشاركة المباشرة)
    import shutil
    shutil.copy(OUTPUT_HTML, OUTPUT_HTML2)
    print(f"  Also saved: {OUTPUT_HTML2}")

    print("\n" + "=" * 60)
    print("  Dashboard built successfully!")
    print(f"  File: {OUTPUT_HTML}")
    print("=" * 60)

    # webbrowser.open('file://' + OUTPUT_HTML)
    # input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
