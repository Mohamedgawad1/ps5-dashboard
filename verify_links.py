import json, os, urllib.parse

d = json.load(open('data.json', 'r', encoding='utf-8'))
WIRING = r'C:\Users\mylap\OneDrive\Desktop\dashboard\WIRING - MASTER'

broken = 0
ok = 0
broken_tags = []
for item in d:
    if item.get('link'):
        fname = urllib.parse.unquote(item['link'].replace('/pdf/', ''))
        fpath = os.path.join(WIRING, fname.replace('/', os.sep))
        if os.path.exists(fpath):
            ok += 1
        else:
            broken += 1
            broken_tags.append(item['asset_tag'] + ' -> ' + fname)

print(f'PDF links verified: {ok} OK, {broken} BROKEN')
if broken_tags:
    print('Broken links:')
    for b in broken_tags[:20]:
        print(f'  {b}')
