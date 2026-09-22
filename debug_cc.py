import fitz, re

pdf_path = r'C:\Users\mylap\Downloads\rfi\CPP-RFI-68-64-12-0094_signed by CPY.pdf'
fz = fitz.open(pdf_path)

for i in range(1, fz.page_count):
    text = fz[i].get_text()
    if not text or not text.strip():
        continue
    lines = text.split('\n')
    
    tag = ''
    for ln in lines:
        m = re.search(r'(PS\d+[\w\-/]*\w)', ln)
        if m and re.match(r'PS\d+[\w\-]*', m.group(1)):
            tag = m.group(1)
            break
    
    # Find Core to Core
    cc = ''
    for idx, ln in enumerate(lines):
        if 'Core to Core' in ln:
            # show context around it
            start = max(0, idx-1)
            end = min(len(lines), idx+8)
            print(f"\nPage {i+1} ({tag}): 'Core to Core' found at line {idx}")
            for k in range(start, end):
                print(f"  line[{k}]: '{lines[k]}'")
            if idx + 6 < len(lines):
                cc = lines[idx + 6].strip()
            break
    
    if not cc:
        # search for resistance values
        for idx, ln in enumerate(lines):
            if '>100' in ln or re.search(r'>\s*\d+', ln):
                print(f"Page {i+1} ({tag}): resistance candidate at line {idx}: '{ln}'")
    
    print(f"  -> Core to Core = '{cc}'")

fz.close()
