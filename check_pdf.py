import json, os

WIRING_DIR = r'C:\Users\mylap\OneDrive\Desktop\dashboard\WIRING - MASTER'

with open(r'C:\Users\mylap\OneDrive\Desktop\dashboard\asset_tag_mapping.json','r',encoding='utf-8') as f:
    mapping = json.load(f)

with open(r'C:\Users\mylap\OneDrive\Desktop\dashboard\data.json','r',encoding='utf-8') as f:
    data = json.load(f)

tag = 'PS5-01-FCV-0004B-CL01'
print(f'Mapping: {tag} -> {mapping.get(tag, "NOT FOUND")}')

import re
base = re.sub(r'-CL\d+$|-CH\d+$|-CD\d+$', '', tag)
print(f'Base:    {base} -> {mapping.get(base, "NOT FOUND")}')

for x in data:
    if x['asset_tag'] == tag:
        print(f'Data: link={x.get("link","")}')
        break

# Check how many links point to missing files
missing = 0
total = 0
for d in data:
    link = d.get('link', '')
    if link:
        total += 1
        fname = link.replace('/pdf/', '')
        import urllib.parse
        fname_decoded = urllib.parse.unquote(fname)
        fpath = os.path.join(WIRING_DIR, fname_decoded)
        if not os.path.exists(fpath):
            missing += 1
            if missing <= 5:
                print(f'  MISSING: {fname_decoded[:100]}')

print(f'\nTotal links: {total} | Missing: {missing} | OK: {total - missing}')
