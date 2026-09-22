import json, urllib.request

INDEX_URL = "https://mohamedgawad1.github.io/ps5-dashboard/index.html"

def extract_itr(html):
    start = html.find('const ITR = ')
    if start < 0:
        print("  'const ITR = ' not found")
        return None
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
        if in_str:
            continue
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(html[start:i+1])
                except Exception as e:
                    print(f"  json parse error: {e}")
                    return None
    print("  no closing brace found")
    return None

try:
    resp = urllib.request.urlopen(INDEX_URL, timeout=20)
    html = resp.read().decode()
    print(f"Fetched {len(html)} bytes")
    itr = extract_itr(html)
    if itr:
        print("SUCCESS:", json.dumps(itr, indent=2)[:500])
    else:
        print("FAILED")
except Exception as e:
    print(f"Error: {e}")
