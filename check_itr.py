import re, json
with open('index.html', encoding='utf-8') as f:
    html = f.read()
start = html.find('const ITR = ')
if start >= 0:
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
                itr = json.loads(html[start:i+1])
                print(json.dumps(itr, indent=2))
                break
