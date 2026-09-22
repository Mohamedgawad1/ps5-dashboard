import json, re

with open(r'C:\Users\mylap\OneDrive\Desktop\dashboard\data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Analyze asset tag patterns
patterns = {}
for d in data:
    tag = d['asset_tag']
    # Find what pattern this matches
    m = re.match(r'(PS5-\d{2}-[A-Z]+-\d+)', tag)
    if m:
        base = m.group(1)
        suffix = tag[len(base):]
        patterns.setdefault(suffix, 0)
        patterns[suffix] += 1

print('=== Tag suffixes (after base PS5-XX-XXX-NNNN) ===')
for s, c in sorted(patterns.items(), key=lambda x: -x[1])[:20]:
    print(f'  {repr(s)}: {c}')

# Check current scan regex
old_pattern = re.compile(r'PS5-[\w]+-[\w]+-[\d]+')
new_pattern = re.compile(r'PS5-[\w]+-[\w]+-[\w]+-[\w]+')

# Count how many tags would match each
old_count = sum(1 for d in data if old_pattern.match(d['asset_tag']))
new_count = sum(1 for d in data if new_pattern.match(d['asset_tag']))
print(f'\nOld regex matches: {old_count}')
print(f'New regex matches: {new_count}')

# Show examples that DON'T match old regex
no_match = [d['asset_tag'] for d in data if not old_pattern.match(d['asset_tag'])]
print(f'\nTags NOT matching old regex: {len(no_match)}')
for t in no_match[:10]:
    print(f'  {t}')
