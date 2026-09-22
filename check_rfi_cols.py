import pandas as pd, os, re

df = pd.read_excel('_temp/PS-5 INSPECTION REGISTER.xlsx', 'PS-5 EIT INSPECTION REGISTER', header=None)
print(f'Shape: {df.shape}')
print()

# Check ALL columns for RFI references
rfi_pdfs = [f for f in os.listdir('WIRING - MASTER') if f.startswith('CPP-RFI')]
print(f'Total RFI PDFs: {len(rfi_pdfs)}')
print()

for col in range(df.shape[1]):
    vals = []
    for i in range(6, len(df)):
        v = str(df.iloc[i, col]).strip() if pd.notna(df.iloc[i, col]) else ''
        if 'RFI' in v.upper() or 'CPP-RFI' in v.upper() or (re.match(r'CPP|RFI', v, re.I)):
            vals.append((i, v))
    if vals:
        print(f'Col {col}: {len(vals)} RFI entries')
        for idx, v in vals[:8]:
            tag = str(df.iloc[idx, 1]).strip() if pd.notna(df.iloc[idx, 1]) else '?'
            print(f'  Row {idx}: Tag={tag}, Val={v}')
        print()
