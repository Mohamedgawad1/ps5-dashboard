import time, os, pandas as pd
files = {
    'ovTasks': 'ovTasks_TestsPlanned_1369.xlsx',
    'PunchL': 'PS-5 EIT PUNCH LIST REGISTER.xlsx',
    'Insp':   'PS-5 INSPECTION REGISTER.xlsx',
    'Master': 'PS5 Master tracker EIT Combined.xlsx',
    'EITDI':  'PS5 EIT CPP AGI Dashboard.xlsx',
}
for k, f in files.items():
    p = os.path.join(r'C:\Users\mylap\OneDrive\Desktop\dashboard', f)
    if not os.path.exists(p):
        print(f'{k}: MISSING')
        continue
    t0 = time.time()
    try:
        xl = pd.ExcelFile(p)
        t1 = time.time()
        print(f'{k}: {os.path.basename(p)} sheets={xl.sheet_names[:3]} read={t1-t0:.1f}s')
    except Exception as e:
        print(f'{k}: ERROR {type(e).__name__} {e}')