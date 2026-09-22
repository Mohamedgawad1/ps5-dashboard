import json, os, re, urllib.parse

BASE = r'C:\Users\mylap\OneDrive\Desktop\dashboard'
WIRING = os.path.join(BASE, 'WIRING - MASTER')

with open(os.path.join(BASE, 'asset_tag_mapping.json'), 'r', encoding='utf-8') as f:
    mapping = json.load(f)

tag = 'PS5-01-FCV-0004B-CL01'
base = re.sub(r'-CL\d+$|-CH\d+$|-CD\d+$', '', tag)

print('=== Search in mapping ===')
print(tag, '->', mapping.get(tag, 'NOT FOUND'))
print(base, '->', mapping.get(base, 'NOT FOUND'))

# Check PS5-01-FCV-0004 matches
matches = [t for t in mapping if 'PS5-01-FCV' in t]
print('\nPS5-01-FCV matches:', matches[:10])

# The cable from=PS5-61-IG-0043, to=PS5-01-FCV-0004B
# Check if these tags are in mapping
for t in ['PS5-61-IG-0043', 'PS5-01-FCV-0004B']:
    print(t, '->', mapping.get(t, 'NOT FOUND'))

# Now check: how many cables in data.json have WRONG links
# (link exists but cable tag is NOT in the PDF content)
with open(os.path.join(BASE, 'asset_tag_mapping.json'), 'r', encoding='utf-8') as f:
    mapping = json.load(f)

# Build reverse map: pdf -> set of tags
pdf_tags = {}
for tag_name, pdf_name in mapping.items():
    if pdf_name not in pdf_tags:
        pdf_tags[pdf_name] = set()
    pdf_tags[pdf_name].add(tag_name)

# Check: cable tags that exist in mapping
with open(os.path.join(BASE, 'data.json'), 'r', encoding='utf-8') as f:
    data = json.load(f)

cable_tags_in_mapping = 0
cable_tags_not_in_mapping = 0
for d in data:
    if d.get('link'):
        tag = d['asset_tag']
        if tag in mapping:
            cable_tags_in_mapping += 1
        else:
            cable_tags_not_in_mapping += 1

print(f'\n=== Cable tag -> PDF mapping ===')
print(f'Cables with link AND tag in mapping: {cable_tags_in_mapping}')
print(f'Cables with link but tag NOT in mapping: {cable_tags_not_in_mapping}')
print(f'Total unique tags in mapping: {len(mapping)}')
