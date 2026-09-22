"""
Scan WIRING-MASTER PDFs for asset tags.
Run once or when new PDFs are added.
"""
import re, os, json, fitz

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WIRING_DIR = os.path.join(BASE_DIR, 'WIRING - MASTER')

pdfs = [f for f in os.listdir(WIRING_DIR) if f.lower().endswith('.pdf')]
tag_pattern = re.compile(r'PS5-[\w]+-[\w]+-[\w]+(?:-[\w]+)?')
all_tags = {}

print(f"Scanning {len(pdfs)} PDFs...")
for i, fname in enumerate(pdfs):
    try:
        doc = fitz.open(os.path.join(WIRING_DIR, fname))
        text = ''.join(page.get_text() for page in doc)
        doc.close()
        tags = set(tag_pattern.findall(text))
        for t in tags:
            clean = t.replace(' ', '').strip()
            if clean not in all_tags:
                all_tags[clean] = fname
        if tags:
            print(f"  [{i+1}/{len(pdfs)}] {fname}: {len(tags)} tags")
    except Exception as e:
        print(f"  [{i+1}/{len(pdfs)}] ERROR {fname}: {e}")

with open(os.path.join(BASE_DIR, 'asset_tag_mapping.json'), 'w') as f:
    json.dump(all_tags, f, indent=2)
print(f"\nDone: {len(all_tags)} unique tags from {len(pdfs)} PDFs")
print("Saved asset_tag_mapping.json")
