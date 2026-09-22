import os, datetime

files = {
    'ovTasks_TestsPlanned_1369.xlsx': 'OV Tasks',
    'PS-5 EIT PUNCH LIST REGISTER.xlsx': 'Punch List',
    'PS-5 INSPECTION REGISTER.xlsx': 'Inspection Register',
    'PS-5 COMPLETIONS DPR SUMMERY - 13 JUNE  2026.xlsx': 'DPR Summary',
}

print('=== File Status ===')
for fname, label in files.items():
    if os.path.exists(fname):
        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(fname))
        age = datetime.datetime.now() - mtime
        status = 'OLD!' if age.days > 1 else 'OK'
        print(f'  {label}: {mtime.strftime("%Y-%m-%d %H:%M")} ({age.days} days old) [{status}]')
    else:
        print(f'  {label}: NOT FOUND')

print('\nAfter updating files, run:')
print('  python cpp_agi_dashboard.py')
