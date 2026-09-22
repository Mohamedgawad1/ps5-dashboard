import pandas as pd
df = pd.read_excel('ovTasks_TestsPlanned_1369.xlsx', sheet_name='Exported from SC')
closed = df[df['Task State'] == 'Closed'].copy()
closed['Closing Date'] = pd.to_datetime(closed['Closing Date'])
today = '2026-08-20'
today_closed = closed[closed['Closing Date'].dt.strftime('%Y-%m-%d') == today]
print(f"Today ({today}) Closed Count: {len(today_closed)}")
print()
print("By Discipline:")
for d in today_closed['Discipline (Summary)'].value_counts().items():
    print(f"  {d[0]}: {d[1]}")
print()
print("Last 5 days:")
for i in range(7):
    from datetime import datetime, timedelta
    dt = datetime(2026, 8, 20) - timedelta(days=i)
    ds = dt.strftime('%Y-%m-%d')
    cnt = len(closed[closed['Closing Date'].dt.strftime('%Y-%m-%d') == ds])
    print(f"  {ds}: {cnt}")
