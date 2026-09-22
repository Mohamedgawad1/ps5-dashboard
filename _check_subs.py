import openpyxl
wb = openpyxl.load_workbook(r'C:\Users\mylap\OneDrive\Desktop\dashboard\PS-5 COMPLETIONS DPR SUMMERY - 19-08-26.xlsx', data_only=True)

# Check ITR list column headers
ws = wb['DETAILED ITR LIST']
print('ITR headers:', [cell.value for cell in ws[1]])

# Check unique values in various columns
for col_idx in range(18):
    vals = set()
    for row in ws.iter_rows(min_row=2, max_row=min(20, ws.max_row), min_col=col_idx+1, max_col=col_idx+1, values_only=True):
        if row[0]:
            vals.add(str(row[0])[:30])
    if vals and len(vals) > 1:
        print(f'Col {col_idx}: {len(vals)} unique - {list(vals)[:5]}')

# Check subsystem from column G (index 6) and H (index 7)
print('\n--- Col G (index 6) sample ---')
for row in ws.iter_rows(min_row=2, max_row=10, min_col=7, max_col=7, values_only=True):
    print(row[0])

print('\n--- Col F (index 5) sample ---')
for row in ws.iter_rows(min_row=2, max_row=10, min_col=6, max_col=6, values_only=True):
    print(row[0])

# Check unique subsystems from detailed punch list
ws2 = wb['DETAILED PUNCH LIST']
print('\nPunch headers:', [cell.value for cell in ws2[1]])
punch_subs = set()
for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, values_only=True):
    if row[7]:
        punch_subs.add(row[7])
print(f'Unique punch subsystems: {len(punch_subs)}')
print('First 10:')
for s in sorted(punch_subs)[:10]:
    print(f'  {s}')

# Extract the actual subsystem IDs from the descriptions
sub_ids = set()
for s in punch_subs:
    parts = s.split(' - ')
    if parts:
        sub_ids.add(parts[0].strip())
print(f'\nUnique subsystem IDs from punch: {len(sub_ids)}')
print('First 10:', sorted(sub_ids)[:10])
