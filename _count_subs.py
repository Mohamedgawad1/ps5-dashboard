import openpyxl
wb = openpyxl.load_workbook(r'C:\Users\mylap\OneDrive\Desktop\dashboard\PS-5 COMPLETIONS DPR SUMMERY - 19-08-26.xlsx', data_only=True)

# Count unique subsystems from RFC PROGRESS
ws = wb['RFC PROGRESS']
subs = []
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
    if row[0]:
        subs.append(row[0])
print(f'RFC PROGRESS rows: {len(subs)}')

# Count from MILESTONE_12
ws2 = wb['MILESTONE_12 AUG 2026']
subs2 = []
for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, values_only=True):
    if row[0]:
        subs2.append(row[0])
print(f'MILESTONE_12 rows: {len(subs2)}')

# Count from MILESTONE & DIS.
ws3 = wb['MILESTONE & DIS.']
subs3 = []
for row in ws3.iter_rows(min_row=8, max_row=ws3.max_row, values_only=True):
    if row[0] and row[0] not in ['Grand Total']:
        subs3.append(row[0])
print(f'MILESTONE & DIS. disciplines: {len(subs3)}')

# Unique subsystems from detailed ITR
ws4 = wb['DETAILED ITR LIST']
itr_subs = set()
for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row, values_only=True):
    if row[7]:
        itr_subs.add(row[7])
print(f'Unique subsystems in ITR list: {len(itr_subs)}')

# Unique subsystems from detailed punch
ws5 = wb['DETAILED PUNCH LIST']
punch_subs = set()
for row in ws5.iter_rows(min_row=2, max_row=ws5.max_row, values_only=True):
    if row[7]:
        punch_subs.add(row[7])
print(f'Unique subsystems in Punch list: {len(punch_subs)}')

# Unique subsystems from EIT PROGRESS
ws6 = wb['EIT PROGRESS']
eit_subs = []
for row in ws6.iter_rows(min_row=2, max_row=ws6.max_row, values_only=True):
    if row[0]:
        eit_subs.append(row[0])
print(f'EIT PROGRESS rows: {len(eit_subs)}')

# Print first 5 from RFC
print('\nFirst 5 RFC subsystems:')
for s in subs[:5]:
    print(f'  {s}')

print('\nFirst 5 ITR subsystems:')
for s in list(itr_subs)[:5]:
    print(f'  {s}')
