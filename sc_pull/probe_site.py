import re
import urllib.request

url = "https://mohamedgawad1.github.io/ps5-dashboard/?v=1790054997833"
req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Range": "bytes=0-524288"})
h = urllib.request.urlopen(req, timeout=40)
data = h.read()
print("bytes:", len(data))
s = data.decode("utf-8", errors="ignore")
m = re.search(r"<title>(.*?)</title>", s, re.I)
print("title:", m.group(1) if m else None)
for m in re.finditer(r"fetch\(\s*[\"']([^\"']+)[\"']", s):
    print("fetch:", m.group(1))
for m in re.finditer(r'src="([^"]+\.js)"', s):
    print("js:", m.group(1))
for m in re.finditer(r'href="([^"]+\.json)"', s):
    print("json:", m.group(1))
print("dashboard_data mentions:", s.count("dashboard_data"))
print("itr mentions:", s.count("itr"), "ITR:", s.count("ITR"))