import sys
sys.path.insert(0, '.')
import itr_upload as U

RULES = U._load_work_type_rules()

def guess(text):
    low = text.lower()
    import re
    desc = re.sub(r"^[^a-z0-9]+", "", low)
    code = desc.split()[0] if desc.split() else ""
    if re.match(r"^t[a-z]*\d", code):
        return "Static Test", [code]
    if re.match(r"^c[a-z]*\d", code):
        return "Conformity Check", [code]
    if re.search(r"\bvisual\s+inspection\b|\binspection\s+of\b", low):
        return "Conformity Check", ["visual inspection / inspection of"]
    if re.search(r"\btesting\b|\btest\b", low):
        return "Static Test", ["testing / test"]
    for name, keys in RULES:
        f = [k for k in keys if k in low]
        if f:
            return name, f
    return None, []

CASES = [
    ("TPX13 - Motor LV", "Static Test"),
    ("CPX13 - Cable Glanding", "Conformity Check"),
    ("Visual Inspection of Electrical Motor Testing with following tag number", "Conformity Check"),
    ("Electrical Cable testing after installation with the following Cable Tag Num", "Static Test"),
    ("Request to Witness Telecom Cable Glanding & termination with following tag number", "Static Test"),
    ("Visual inspection of Instrument cable Glanding and Termination", "Conformity Check"),
    ("Cable Glanding & termination", "Conformity Check"),
]

ok = True
for text, want in CASES:
    got, why = guess(text)
    flag = "OK " if got == want else "BAD"
    if got != want:
        ok = False
    print("%s | %-24s -> %-18s (%s)" % (flag, text[:24], got, ",".join(why[:2])))
print()
print("all rules match" if ok else "there are mismatches")
