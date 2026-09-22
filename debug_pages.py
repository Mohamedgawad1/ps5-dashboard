import fitz, re

pdf_path = r'C:\Users\mylap\Downloads\rfi\CPP-RFI-68-64-12-0094_signed by CPY.pdf'
fz = fitz.open(pdf_path)

for i in range(1, min(4, fz.page_count)):
    text = fz[i].get_text()
    if not text:
        print(f"Page {i+1}: EMPTY")
        continue
    lines = text.split('\n')
    tag = ''
    
    # Strategy 1: same line
    for idx, ln in enumerate(lines):
        m = re.search(r'CABLE\s+TAG\s+(?:NUMBER|NO\b)\s+(PS\d+[\w\-/]*\w)', ln, re.IGNORECASE)
        if m:
            tag = m.group(1)
            print(f"Page {i+1}: Strategy1 found: {tag}")
            break
    # Strategy 2: next line
    if not tag:
        for idx, ln in enumerate(lines):
            if re.search(r'CABLE\s+TAG\s+(?:NUMBER|NO\b)', ln, re.IGNORECASE):
                for j in range(idx+1, min(idx+5, len(lines))):
                    c = lines[j].strip()
                    if c:
                        tag = c
                        print(f"Page {i+1}: Strategy2 found: {tag}")
                        break
                break
    # Strategy 3: fallback
    if not tag:
        for ln in lines:
            m_ps = re.search(r'(PS\d+[\w\-/]*\w)', ln)
            if m_ps:
                tag = m_ps.group(1)
                print(f"Page {i+1}: Strategy3 found: {tag}")
                break
    
    tag_ok = bool(re.match(r'PS\d+[\w\-]*', tag)) or bool(re.match(r'[\w]+-[\w]+-[\w]+', tag))
    if not tag:
        print(f"Page {i+1}: NO TAG FOUND. First 5 lines: {lines[:5]}")
    elif not tag_ok:
        print(f"Page {i+1}: TAG '{tag}' FAILED validation")

fz.close()
