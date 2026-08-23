"""
Daily Data Updater - CPP AGI E&I&T
Scans ALL PDFs from Downloads recursively + rebuilds data.json
"""
import re, os, json, shutil, urllib.parse, pandas as pd
from datetime import datetime
from pathlib import Path

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS = Path.home() / "Downloads"
LOG_FILE = os.path.join(BASE_DIR, 'update_log.txt')
WIRING_DIR = os.path.join(BASE_DIR, 'WIRING - MASTER')

SKIP_PATTERNS = ['mohamed', 'cv', 'attendance', 'punch', 'punchlist',
                 'oil.pdf', 'task datasheet', 'document_control',
                 '1784709', '1786128', 'resume', 'letter', 'contract',
                 'invoice', 'receipt', 'photo', 'image', 'screenshot']

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
    if tag in asset_tag_map:
        pdf = asset_tag_map[tag].replace('\\', '/')
        if os.path.exists(os.path.join(WIRING_DIR, pdf.replace('/', os.sep))):
            return f'/pdf/{urllib.parse.quote(pdf)}'
    base = get_base_tag(tag)
    if base != tag and base in asset_tag_map:
        pdf = asset_tag_map[base].replace('\\', '/')
        if os.path.exists(os.path.join(WIRING_DIR, pdf.replace('/', os.sep))):
            return f'/pdf/{urllib.parse.quote(pdf)}'
    stripped = re.sub(r'-(CL|CH|CD|CJ|CS|CT|CC|CB)\d+$', '', tag)
    if stripped != tag and stripped in asset_tag_map:
        pdf = asset_tag_map[stripped].replace('\\', '/')
        if os.path.exists(os.path.join(WIRING_DIR, pdf.replace('/', os.sep))):
            return f'/pdf/{urllib.parse.quote(pdf)}'
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

def copy_all_pdfs_from_downloads():
    """Recursively scan ALL subfolders in Downloads for PDFs and copy to WIRING-MASTER."""
    log("Scanning ALL Downloads subfolders for PDFs...")
    if not DOWNLOADS.exists():
        return 0
    wm_names = set(f for f in os.listdir(WIRING_DIR) if f.lower().endswith('.pdf'))
    copied = 0
    skipped = 0
    for pdf_path in DOWNLOADS.rglob("*.pdf"):
        if pdf_path.is_dir():
            continue
        fname = pdf_path.name
        if fname.lower() in {n.lower() for n in wm_names}:
            continue
        if any(p in fname.lower() for p in SKIP_PATTERNS):
            continue
        if not fname.lower().startswith('cpp-rfi'):
            continue
        try:
            shutil.copy2(str(pdf_path), os.path.join(WIRING_DIR, fname))
            log(f"  COPIED [{pdf_path.parent.name}]: {fname}")
            wm_names.add(fname)
            copied += 1
        except Exception as e:
            log(f"  ERROR copying {fname}: {e}")
    log(f"  Copied {copied} new PDFs from Downloads")
    return copied

def main():
    log("=" * 50)
    log("Starting daily update...")
    os.chdir(BASE_DIR)

    with open('asset_tag_mapping.json') as f:
        asset_tag_map = json.load(f)

    copy_all_pdfs_from_downloads()

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

    log("Reading Inspection Register RFI (col 8 + col 13)...")
    df_ir = pd.read_excel(os.path.join('_temp', 'PS-5 INSPECTION REGISTER.xlsx'),
                          'PS-5 EIT INSPECTION REGISTER', header=None)
    rfi_pdfs = [f for f in os.listdir(WIRING_DIR) if f.startswith('CPP-RFI')]
    gl_map = {}
    rfi_matched = 0
    rfi_pdf_norm = {}
    for f in rfi_pdfs:
        norm = f.replace('.pdf', '').replace('.PDF', '').replace(' ', '').replace('-', '').replace('_', '').lower()
        rfi_pdf_norm[norm] = f
    rfi_pdf_by_num = {}
    for f in rfi_pdfs:
        num_m = re.search(r'(\d{4})', f)
        if num_m:
            num = num_m.group(1)
            rfi_pdf_by_num[num] = f

    def find_rfi_pdf(rfi_text):
        rfi_clean = rfi_text.replace('/', '').replace(' ', '')
        rfi_norm = rfi_clean.replace('-', '').replace('_', '').lower()
        for f in rfi_pdfs:
            if rfi_clean == f.replace('.pdf', '').replace('.PDF', '').replace(' ', ''):
                return f
        if rfi_norm in rfi_pdf_norm:
            return rfi_pdf_norm[rfi_norm]
        num_m = re.search(r'(\d{4})$', rfi_clean)
        if num_m:
            num = num_m.group(1)
            if num in rfi_pdf_by_num:
                return rfi_pdf_by_num[num]
        return None

    for i in range(6, len(df_ir)):
        tag = clean_tag(df_ir.iloc[i, 1]) if pd.notna(df_ir.iloc[i, 1]) else ''
        if not tag.startswith('PS5-'): continue
        rfi_8 = str(df_ir.iloc[i, 8]).strip() if pd.notna(df_ir.iloc[i, 8]) else ''
        if not rfi_8 or rfi_8 == 'nan': rfi_8 = ''
        gl_rfi = str(df_ir.iloc[i, 13]).strip() if pd.notna(df_ir.iloc[i, 13]) else ''
        if not gl_rfi or gl_rfi == 'nan': gl_rfi = ''
        gl_st = str(df_ir.iloc[i, 14]).strip() if pd.notna(df_ir.iloc[i, 14]) else ''
        best_rfi = rfi_8 or gl_rfi
        if not best_rfi: continue
        found_pdf = find_rfi_pdf(best_rfi)
        rfi_pdf = f'/pdf/{found_pdf}' if found_pdf else ''
        if found_pdf: rfi_matched += 1
        if tag not in gl_map or (rfi_pdf and not gl_map[tag].get('rfi_pdf')):
            gl_map[tag] = {'rfi': best_rfi, 'status': gl_st, 'rfi_pdf': rfi_pdf}
    log(f"  Inspection: {len(gl_map)} assets | Matched to PDF: {rfi_matched}")

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
