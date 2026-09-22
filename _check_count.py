import pandas as pd
df = pd.read_excel('ovTasks_TestsPlanned_1369.xlsx', sheet_name='Exported from SC')
cpp = df[df['Responsible Company (Summary)'] == 'CPP AGI']
disc = cpp['Discipline (Summary)'].astype(str).str.strip()
print('=== ALL TASKS ===')
print('Total:', len(df), '| Closed:', len(df[df['Task State']=='Closed']))
print()
print('=== CPP AGI ===')
print('Total:', len(cpp), '| Closed:', len(cpp[cpp['Task State']=='Closed']), '| Open:', len(cpp[cpp['Task State']!='Closed']))
print()
for label, pat in [('E - Electrical','Elec'), ('I - Instrument','Instr'), ('T - Telecom','Tele')]:
    sub = cpp[disc.str.contains(pat, na=False)]
    c = len(sub[sub['Task State']=='Closed'])
    o = len(sub) - c
    print(f'{label}: Total={len(sub)} Closed={c} Open={o}')
