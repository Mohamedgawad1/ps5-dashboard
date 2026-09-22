import re

s = open("index.html", encoding="utf-8", errors="ignore").read()
print("size MB:", round(len(s) / 1048576, 1))
for key in ["chartCombinedDaily", "combinedChart(", "PUNCH =", "RFI =", "ITR =",
            "ITR.daily", "var ITR", "const ITR", "window.ITR", "var PUNCH", "var RFI"]:
    print("%-22s -> %d" % (key, s.count(key)))

m = re.search(r"var ITR\s*=\s*\{.{0,300}", s)
print("ITR def:", m.group(0)[:200] if m else None)

m = re.search(r'"daily"\s*:\s*\[.{0,200}', s)
print("daily arr sample:", m.group(0)[:200] if m else None)

m3 = re.search(r".{60}combinedChart\(.{0,200}", s)
print("call ctx:", m3.group(0)[:260] if m3 else None)

# count daily objects
m4 = re.search(r'"daily"\s*:\s*\[(.*?)\]', s, re.S)
if m4:
    print("daily len chars:", len(m4.group(1)))

# look for PUNCH/RFI data presence
m5 = re.search(r"PUNCH\s*=\s*(\{|\[)", s)
m6 = re.search(r"RFI\s*=\s*(\{|\[)", s)
print("PUNCH assign:", m5.group(0) if m5 else None, "| RFI assign:", m6.group(0) if m6 else None)

# possible JS errors near combined
for tag in ["chartCombinedDaily", "combinedChart"]:
    idx = [i for i in range(len(s)) if s.startswith(tag, i)]
    for i in idx[:3]:
        print(tag, "@", i, "ctx:", s[max(0, i - 40):i + 90].replace("\n", " "))