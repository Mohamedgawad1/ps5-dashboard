import json

with open('index.html', encoding='utf-8') as f:
    html = f.read()

start = html.find('const ITR = ')
if start < 0:
    print('ITR not found')
else:
    start = html.index('{', start)
    depth = 0
    in_str = False
    esc = False
    for i in range(start, len(html)):
        ch = html[i]
        if esc:
            esc = False; continue
        if ch == '\\' and in_str:
            esc = True; continue
        if ch == '"':
            in_str = not in_str; continue
        if in_str: continue
        if ch == '{': depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                try:
                    itr = json.loads(html[start:i+1])
                    print('Parsed OK')
                    print('keys:', list(itr.keys()))
                    print('hourly_closed_eit:', itr.get('hourly_closed_eit'))
                    print('today_submitted_assets:', itr.get('today_submitted_assets'))
                except Exception as e:
                    print('Parse error:', e)
                break
