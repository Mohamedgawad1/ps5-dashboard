import time, sys
sys.path.insert(0, r'C:\Users\mylap\OneDrive\Desktop\dashboard')
import cpp_agi_dashboard as m

t0 = time.time()
print("Calling build_itr_data...", flush=True)
data = m.build_itr_data(r'C:\Users\mylap\OneDrive\Desktop\dashboard\ovTasks_TestsPlanned_1369.xlsx')
print(f"Done in {time.time()-t0:.1f}s", flush=True)
print("total_closed_project:", data.get('total_closed_project'))
print("eit_summary:", data.get('eit_summary'))