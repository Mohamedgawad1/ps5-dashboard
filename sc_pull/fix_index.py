s = open("index.html", encoding="utf-8").read()
old = 'html += \'</table><script src="live_itr.js" charset="utf-8"></script>\'\n</body></html>\';'
new = 'html += \'</table>\';\n  html += \'</body></html>\';'
if old in s:
    s = s.replace(old, new, 1)
    open("index.html", "w", encoding="utf-8").write(s)
    print("REPAIRED")
else:
    print("pattern not found - trying alt")
    import re
    m = re.search(r"html\s*\+=\s*'</table><script src=\"live_itr\.js\"[^']*</body></html>';", s, re.S)
    print("alt match:", bool(m))
    if m:
        s = s.replace(m.group(0), "html += '</table>';\n  html += '</body></html>';", 1)
        open("index.html", "w", encoding="utf-8").write(s)
        print("REPAIRED alt")
    else:
        print("NO MATCH FOUND - manual inspect needed")