"""
Daily Data Updater - CPP AGI EIT
Optimized: skips PDF scanning (use scan_pdfs.py when new PDFs added)
"""
import re, os, json, shutil, urllib.parse, pandas as pd
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS = os.path.join(os.path.expanduser('~'), 'Downloads')
LOG_FILE = os.path.join(BASE_DIR, 'update_log.txt')
WIRING_DIR = os.path.join(BASE_DIR, 'WIRING - MASTER')

def log(msg):
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(line + '\n')

def clean_tag(t):
    if pd.isna(t): return ''
    return str(t).replace(' ', '').replace('\n', '')

def get_base_tag(tag):
    return re.sub(r'-(CL|CH|CD|CJ|CS|CT|CC|CB|IJ|IE|IG|DM|GU|ISO)\d*$', '', tag)

def find_pdf_link(tag, asset_tag_map):
    if not tag: return ''
    # 1. Exact match
    if tag in asset_tag_map:
        pdf = asset_tag_map[tag].replace('\\', '/')
        if os.path.exists(os.path.join(WIRING_DIR, pdf.replace('/', os.sep))):
            return f'/pdf/{urllib.parse.quote(pdf)}'
    # 2. Base name match
    base = get_base_tag(tag)
    if base != tag and base in asset_tag_map:
        pdf = asset_tag_map[base].replace('\\', '/')
        if os.path.exists(os.path.join(WIRING_DIR, pdf.replace('/', os.sep))):
            return f'/pdf/{urllib.parse.quote(pdf)}'
    # 3. Try with common suffixes stripped
    stripped = re.sub(r'-(CL|CH|CD|CJ|CS|CT|CC|CB)\d+$', '', tag)
    if stripped != tag and stripped in asset_tag_map:
        pdf = asset_tag_map[stripped].replace('\\', '/')
        if os.path.exists(os.path.join(WIRING_DIR, pdf.replace('/', os.sep))):
            return f'/pdf/{urllib.parse.quote(pdf)}'
    # 4. Try matching by prefix (first 3 parts: PS5-XX-XX)
    parts = tag.split('-')
    if len(parts) >= 4:
        prefix = '-'.join(parts[:4])
        for k, v in asset_tag_map.items():
            if k.startswith(prefix) and k != tag:
                pdf = v.replace('\\', '/')
                if os.path.exists(os.path.join(WIRING_DIR, pdf.replace('/', os.sep))):
                    return f'/pdf/{urllib.parse.quote(pdf)}'
                break
    return ''

def copy_new_pdfs():
    log("Copying new PDFs from Downloads...")
    if not os.path.exists(DOWNLOADS): return 0
    dl = set(f for f in os.listdir(DOWNLOADS) if f.lower().endswith('.pdf'))
    wm = set(f for f in os.listdir(WIRING_DIR) if f.lower().endswith('.pdf'))
    skip = ['Mohamed', 'CV', 'attendance', 'PUNCH', 'Punchlist', 'oil.pdf',
            'Task Datasheet', 'Document_Control', '1784709', '1786128']
    n = 0
    for f in sorted(dl - wm):
        if any(p.lower() in f.lower() for p in skip): continue
        try:
            shutil.copy2(os.path.join(DOWNLOADS, f), os.path.join(WIRING_DIR, f))
            n += 1
        except: pass
    log(f"  Copied {n} new PDFs")
    return n

def main():
    log("=" * 50)
    log("Starting daily update...")
    os.chdir(BASE_DIR)

    with open('asset_tag_mapping.json') as f:
        asset_tag_map = json.load(f)

    copy_new_pdfs()

    # Try to copy locked files to _temp first
    for f in ['PS5 Master tracker EIT Combined.xlsx', 'PS5 EIT CPP AGI Dashboard.xlsx']:
        src = os.path.join(BASE_DIR, f)
        dst = os.path.join(BASE_DIR, '_temp', f)
        try:
            shutil.copy2(src, dst)
        except PermissionError:
            if not os.path.exists(dst):
                raise

    xls_master = pd.ExcelFile(os.path.join('_temp', 'PS5 Master tracker EIT Combined.xlsx'))
    xls_dash = pd.ExcelFile(os.path.join('_temp', 'PS5 EIT CPP AGI Dashboard.xlsx'))

    # Subsystem mapping
    log("Building subsystem mapping...")
    asset_sub = {}
    all_subs = set()

    df_itm = pd.read_excel(xls_master, 'CPP AGI ITR Master', header=None)
    for i in range(7, len(df_itm)):
        tag = clean_tag(df_itm.iloc[i, 2]) if pd.notna(df_itm.iloc[i, 2]) else ''
        sub = str(df_itm.iloc[i, 6]).strip() if pd.notna(df_itm.iloc[i, 6]) else ''
        if tag.startswith('PS5-') and sub.startswith('PS5-'):
            asset_sub[tag] = sub
            all_subs.add(sub)

    for sheet in ['Subsystem Summary (E+I+T)', 'Subsystem Summary (Combined)']:
        df = pd.read_excel(xls_master, sheet, header=None)
        for i in range(1, len(df)):
            s = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ''
            if s.startswith('PS5-') and ' - ' in s: all_subs.add(s)

    df_cr = pd.read_excel(xls_dash, 'Cable Routes', header=2)
    for _, r in df_cr.iterrows():
        sub = str(r['Subsystem']).strip() if pd.notna(r['Subsystem']) else ''
        if sub.startswith('PS5-'):
            all_subs.add(sub)
            for col in ['From Tag', 'To Tag', 'Cable Tag']:
                t = clean_tag(r[col]) if pd.notna(r[col]) else ''
                if t.startswith('PS5-'): asset_sub[t] = sub

    df_pc = pd.read_excel(xls_dash, 'Panel Connections', header=2)
    for _, r in df_pc.iterrows():
        sub = str(r['Subsystem']).strip() if pd.notna(r['Subsystem']) else ''
        if sub.startswith('PS5-'):
            all_subs.add(sub)
            for col in ['Panel Tag', 'Connected Tag']:
                t = clean_tag(r[col]) if pd.notna(r[col]) else ''
                if t.startswith('PS5-'): asset_sub[t] = sub
    log(f"  {len(all_subs)} subsystems, {len(asset_sub)} mappings")

    # ovTasks
    log("Reading ovTasks (10K+ tags)...")
    ovtasks_src = os.path.join(BASE_DIR, 'ovTasks_TestsPlanned_1369.xlsx')
    ovtasks_dst = os.path.join(BASE_DIR, '_temp', 'ovTasks_TestsPlanned_1369.xlsx')
    try:
        shutil.copy2(ovtasks_src, ovtasks_dst)
    except PermissionError:
        if not os.path.exists(ovtasks_dst):
            raise
    df_ov = pd.read_excel(ovtasks_dst, 'Exported from SC',
                          usecols=[1, 7, 8, 11, 21, 28], header=0)
    ovtasks = {}
    for _, r in df_ov.iterrows():
        tag = clean_tag(r.get('Asset - Tag', ''))
        if not tag.startswith('PS5-'): continue
        if tag not in ovtasks:
            ovtasks[tag] = {
                'subsystem': str(r.get('Systemization - Subsystem (Summary)', '')).strip() if pd.notna(r.get('Systemization - Subsystem (Summary)')) else '',
                'discipline': str(r.get('Discipline', '')).strip() if pd.notna(r.get('Discipline')) else '',
                'description': str(r.get('Asset - Description', '')).strip() if pd.notna(r.get('Asset - Description')) else '',
            }
    log(f"  {len(ovtasks)} unique tags")

    # Glanding & Termination RFI (col 13, 14) + match to CPP-RFI PDFs
    log("Reading Glanding RFI + matching CPP-RFI PDFs...")
    df_ir = pd.read_excel(os.path.join('_temp', 'PS-5 INSPECTION REGISTER.xlsx'),
                          'PS-5 EIT INSPECTION REGISTER', header=None)
    rfi_pdfs = [f for f in os.listdir(WIRING_DIR) if f.startswith('CPP-RFI')]
    gl_map = {}
    rfi_matched = 0
    # Build normalized RFI PDF lookup for flexible matching
    rfi_pdf_norm = {}
    for f in rfi_pdfs:
        norm = f.replace('.pdf', '').replace('.PDF', '').replace(' ', '').replace('-', '').replace('_', '').lower()
        rfi_pdf_norm[norm] = f
    # Also keep last 4 digits index
    rfi_pdf_by_num = {}
    import re as _re
    for f in rfi_pdfs:
        num_m = _re.search(r'(\d{4})', f)
        if num_m:
            num = num_m.group(1)
            rfi_pdf_by_num[num] = f

    for i in range(6, len(df_ir)):
        tag = clean_tag(df_ir.iloc[i, 1]) if pd.notna(df_ir.iloc[i, 1]) else ''
        if not tag.startswith('PS5-'): continue
        gl_rfi = str(df_ir.iloc[i, 13]).strip() if pd.notna(df_ir.iloc[i, 13]) else ''
        gl_st = str(df_ir.iloc[i, 14]).strip() if pd.notna(df_ir.iloc[i, 14]) else ''
        if not gl_rfi or gl_rfi == 'nan': continue
        rfi_clean = gl_rfi.replace('/', '').replace(' ', '')
        rfi_norm = rfi_clean.replace('-', '').replace('_', '').lower()
        found_pdf = None
        # 1. Exact match
        for f in rfi_pdfs:
            if rfi_clean == f.replace('.pdf', '').replace('.PDF', '').replace(' ', ''):
                found_pdf = f
                break
        # 2. Normalized match (ignore dashes, case)
        if not found_pdf:
            if rfi_norm in rfi_pdf_norm:
                found_pdf = rfi_pdf_norm[rfi_norm]
        # 3. Last 4 digits match
        if not found_pdf:
            num_m = _re.search(r'(\d{4})$', rfi_clean)
            if num_m:
                num = num_m.group(1)
                if num in rfi_pdf_by_num:
                    found_pdf = rfi_pdf_by_num[num]
        rfi_pdf = f'/pdf/{found_pdf}' if found_pdf else ''
        if found_pdf: rfi_matched += 1
        gl_map[tag] = {'rfi': gl_rfi, 'status': gl_st, 'rfi_pdf': rfi_pdf}
    log(f"  Glanding: {len(gl_map)} | Matched to PDF: {rfi_matched}")

    # Cable schedules
    log("Parsing cable schedules...")
    cables = []
    for sheet in ['Electrical Cable Schedule', 'Instrument Cable Schedule', 'Telecom Cable Schedule']:
        df = pd.read_excel(xls_master, sheet, header=None)
        n = 0
        for i in range(14, len(df)):
            cr = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ''
            if not cr or not cr.replace(' ', '').startswith('PS5-'): continue
            ft = str(df.iloc[i, 10]).strip() if pd.notna(df.iloc[i, 10]) else ''
            fd = str(df.iloc[i, 12]).strip() if pd.notna(df.iloc[i, 12]) else ''
            tt = str(df.iloc[i, 17]).strip() if pd.notna(df.iloc[i, 17]) else ''
            td = str(df.iloc[i, 19]).strip() if pd.notna(df.iloc[i, 19]) else ''
            cables.append({'cable': clean_tag(cr), 'from': clean_tag(ft), 'fd': fd if fd != 'nan' else '',
                           'to': clean_tag(tt), 'td': td if td != 'nan' else ''})
            n += 1
        log(f"  {sheet}: {n}")

    df_th = pd.read_excel(xls_master, 'Trace Heating - Cable Schedule', header=None)
    n = 0
    for i in range(3, len(df_th)):
        cr = str(df_th.iloc[i, 1]).strip() if pd.notna(df_th.iloc[i, 1]) else ''
        if not cr or not cr.replace(' ', '').startswith('PS5-'): continue
        ft = str(df_th.iloc[i, 10]).strip() if pd.notna(df_th.iloc[i, 10]) else ''
        fd = str(df_th.iloc[i, 12]).strip() if pd.notna(df_th.iloc[i, 12]) else ''
        tt = str(df_th.iloc[i, 17]).strip() if pd.notna(df_th.iloc[i, 17]) else ''
        td = str(df_th.iloc[i, 19]).strip() if pd.notna(df_th.iloc[i, 19]) else ''
        cables.append({'cable': clean_tag(cr), 'from': clean_tag(ft), 'fd': fd if fd != 'nan' else '',
                       'to': clean_tag(tt), 'td': td if td != 'nan' else ''})
        n += 1
    log(f"  Trace Heating: {n}")

    # Build data.json
    log("Building data.json...")
    data = []
    cable_tags = set()

    for c in cables:
        cable_tags.add(c['cable'])
        sub = ''
        for t in [c['from'], c['to'], c['cable']]:
            if t in asset_sub: sub = asset_sub[t]; break

        from_str = f"{c['from']} - {c['fd']}" if c['from'] and c['fd'] else c['from']
        to_str = f"{c['to']} - {c['td']}" if c['to'] and c['td'] else c['to']

        link = ''
        for t in [c['from'], c['to'], c['cable']]:
            link = find_pdf_link(t, asset_tag_map)
            if link: break

        base = re.sub(r'-CL\d+$|-CH\d+$|-CD\d+$', '', c['cable'])
        gl = gl_map.get(c['cable'], gl_map.get(base))

        data.append({'subsystem': sub, 'asset_tag': c['cable'], 'from': from_str, 'to': to_str,
                     'link': link, 'rfi': gl['rfi'] if gl else '', 'rfi_status': gl['status'] if gl else '',
                     'rfi_pdf': gl['rfi_pdf'] if gl else ''})

    for tag, info in ovtasks.items():
        if tag in cable_tags: continue
        link = find_pdf_link(tag, asset_tag_map)
        sub = info['subsystem'] or asset_sub.get(tag, '')
        base = re.sub(r'-CL\d+$|-CH\d+$|-CD\d+$', '', tag)
        gl = gl_map.get(tag, gl_map.get(base))
        data.append({'subsystem': sub, 'asset_tag': tag, 'from': '', 'to': '',
                     'link': link, 'rfi': gl['rfi'] if gl else '', 'rfi_status': gl['status'] if gl else '',
                     'rfi_pdf': gl['rfi_pdf'] if gl else '',
                     'discipline': info['discipline'], 'description': info['description']})

    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)
    shutil.copy('data.json', os.path.join('mobile_app', 'data.json'))

    w_pdf = sum(1 for d in data if d['link'])
    w_rfi = sum(1 for d in data if d['rfi'])
    w_rfi_pdf = sum(1 for d in data if d.get('rfi_pdf'))
    log(f"DONE! {len(data)} assets | {len(cables)} cables | {len(ovtasks)} ovtasks | "
        f"{w_pdf} w/PDF | {w_rfi} w/RFI | {w_rfi_pdf} w/RFI-PDF | {len(all_subs)} subs")

if __name__ == '__main__':
    main()
