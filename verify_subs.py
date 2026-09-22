import json
d = json.load(open('data.json', 'r', encoding='utf-8'))

subs = {}
for item in d:
    s = item.get('subsystem', '') or 'No Subsystem'
    if s not in subs:
        subs[s] = {'total': 0, 'with_pdf': 0}
    subs[s]['total'] += 1
    if item.get('link'):
        subs[s]['with_pdf'] += 1

print('Top 15 subsystems:')
for s, v in sorted(subs.items(), key=lambda x: -x[1]['with_pdf'])[:15]:
    pct = v['with_pdf'] * 100 // v['total'] if v['total'] else 0
    print(f'  {s}: {v["with_pdf"]}/{v["total"]} ({pct}%)')

print(f'\nTotal subsystems: {len(subs)}')
total = sum(v['total'] for v in subs.values())
linked = sum(v['with_pdf'] for v in subs.values())
print(f'Overall: {linked}/{total} ({linked*100//total}%) with PDF')
