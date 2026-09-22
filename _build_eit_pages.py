import pandas as pd
import json

EXCEL = 'PS5 EIT CPP AGI Dashboard.xlsx'

def sheet_to_rows(sheet):
    df = pd.read_excel(EXCEL, sheet_name=sheet, header=None)
    rows = []
    for _, row in df.iterrows():
        vals = []
        for v in row.tolist():
            if v is None:
                vals.append('')
            elif isinstance(v, float) and v.is_integer():
                vals.append(str(int(v)))
            elif isinstance(v, (pd.Timestamp,)):
                vals.append(str(v.date()) if hasattr(v, 'date') else str(v))
            else:
                s = str(v).strip()
                vals.append(s if s != 'nan' else '')
        while vals and vals[-1] == '':
            vals.pop()
        if any(vals):
            rows.append(vals)
    return rows

result = {}
xls = pd.ExcelFile(EXCEL)
for sheet in xls.sheet_names:
    result[sheet] = sheet_to_rows(sheet)

with open('_eit_pages.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False)

for k, v in result.items():
    print(f'{k}: {len(v)} rows, max cols {max(len(r) for r in v) if v else 0}')
