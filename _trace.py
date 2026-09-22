import time, sys, pandas as pd
sys.path.insert(0, r'C:\Users\mylap\OneDrive\Desktop\dashboard')
import cpp_agi_dashboard as m

PATH = r'C:\Users\mylap\OneDrive\Desktop\dashboard\ovTasks_TestsPlanned_1369.xlsx'
t0 = time.time()

print("1: safe_read_excel...", flush=True)
df = m.safe_read_excel(PATH, sheet_name='Exported from SC')
print(f"   done {time.time()-t0:.1f}s rows={len(df)}", flush=True)

t0 = time.time()
print("2: closed filter...", flush=True)
closed = df[df['Task State'] == 'Closed'].copy()
print("   ok", flush=True)

t0 = time.time()
print("3: to_datetime...", flush=True)
closed['Closing Date'] = pd.to_datetime(closed['Closing Date'])
print(f"   done {time.time()-t0:.1f}s", flush=True)

t0 = time.time()
print("4: dropna...", flush=True)
closed = closed.dropna(subset=['Closing Date'])
print("   ok", flush=True)

t0 = time.time()
print("5: submitted...", flush=True)
submitted = df[df['Task State'] == 'Submitted'].copy()
submitted['Closing Date'] = pd.to_datetime(submitted['Closing Date'])
print("   ok", flush=True)

t0 = time.time()
print("6: now...", flush=True)
now = pd.Timestamp(closed['Closing Date'].max())
print(f"   now={now}", flush=True)

t0 = time.time()
print("7: today_closed filter...", flush=True)
today_closed = closed[closed['Closing Date'].dt.date == now.date()].copy()
print(f"   done {time.time()-t0:.1f}s n={len(today_closed)}", flush=True)

t0 = time.time()
print("8: hourly...", flush=True)
today_closed['hour'] = today_closed['Closing Date'].dt.hour.apply(lambda h: f"{h:02d}:00")
print("   ok", flush=True)

t0 = time.time()
print("9: pivot...", flush=True)
p = today_closed.groupby(['hour', 'disc']).size().unstack(fill_value=0)
print(f"   done {time.time()-t0:.1f}s", flush=True)

import threading
for name in ['main']:
    for t in threading.enumerate():
        if t.name == name:
            print("ALIVE", flush=True)

print("ALL DONE", flush=True)