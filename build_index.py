"""Step 1: Extract all text from PDFs into a fast-searchable index."""
import os, json, re
import fitz

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WIRING_DIR = os.path.join(BASE_DIR, 'WIRING - MASTER')

TAG_RE = re.compile(r'PS5-[\w]+-[\w]+-[\w]+(?:-[\w]+)?')

# Collect all PDFs
pdfs = []
for f in os.listdir(WIRING_DIR):
    if f.lower().endswith('.pdf') and os.path.isfile(os.path.join(WIRING_DIR, f)):
        pdfs.append((os.path.join(WIRING_DIR, f), f))
for root, dirs, files in os.walk(WIRING_DIR):
    for f in files:
        if f.lower().endswith('.pdf'):
            fpath = os.path.join(root, f)
            if fpath not in [p[0] for p in pdfs]:
                pdfs.append((fpath, os.path.relpath(fpath, WIRING_DIR)))

print(f"Scanning {len(pdfs)} PDFs...")

tag_to_pdfs = {}  # tag -> list of pdf filenames
errors = 0

for i, (fpath, fname) in enumerate(pdfs):
    try:
        doc = fitz.open(fpath)
        text = ''
        for page in doc:
            text += page.get_text()
        doc.close()
        
        tags = set(TAG_RE.findall(text))
        for t in tags:
            t = t.strip().replace(' ', '')
            if t not in tag_to_pdfs:
                tag_to_pdfs[t] = []
            if fname not in tag_to_pdfs[t]:
                tag_to_pdfs[t].append(fname)
        
        if (i + 1) % 50 == 0:
            print(f"  [{i+1}/{len(pdfs)}] {len(tag_to_pdfs)} tags found, errors={errors}")
    except Exception as e:
        errors += 1

print(f"\nDONE: {len(tag_to_pdfs)} unique tags from {len(pdfs)} PDFs ({errors} errors)")

# Save
with open(os.path.join(BASE_DIR, 'tag_index.json'), 'w', encoding='utf-8') as f:
    json.dump(tag_to_pdfs, f, indent=2, ensure_ascii=False)
print("Saved tag_index.json")
