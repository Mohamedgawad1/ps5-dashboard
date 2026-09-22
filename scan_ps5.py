import re, os, json, fitz

BASE_DIR = r'C:\Users\mylap\OneDrive\Desktop\dashboard'
WIRING_DIR = os.path.join(BASE_DIR, 'WIRING - MASTER')
PS5_DIR = os.path.join(WIRING_DIR, 'PS5')
TAG_PATTERN = re.compile(r'PS5-[\w]+-[\w]+-[\w]+(?:-[\w]+)?')

# Load existing mapping
with open(os.path.join(BASE_DIR, 'asset_tag_mapping.json'), 'r', encoding='utf-8') as f:
    all_tags = json.load(f)

# Scan PS5 subfolder recursively
pdfs = []
for root, dirs, files in os.walk(PS5_DIR):
    for f in files:
        if f.lower().endswith('.pdf'):
            pdfs.append(os.path.join(root, f))

print(f"Scanning {len(pdfs)} PDFs from PS5 subfolder...")
new_count = 0
for i, fpath in enumerate(pdfs):
    try:
        doc = fitz.open(fpath)
        text = ''
        for page in doc:
            text += page.get_text()
        doc.close()
        tags = set(TAG_PATTERN.findall(text))
        fname = os.path.relpath(fpath, WIRING_DIR)
        for t in tags:
            clean = t.replace(' ', '').strip()
            if clean not in all_tags:
                new_count += 1
            all_tags[clean] = fname
        if tags:
            print(f"  [{i+1}/{len(pdfs)}] {fname}: {len(tags)} tags")
    except Exception as e:
        print(f"  [{i+1}/{len(pdfs)}] ERROR: {e}")

with open(os.path.join(BASE_DIR, 'asset_tag_mapping.json'), 'w', encoding='utf-8') as f:
    json.dump(all_tags, f, indent=2, ensure_ascii=False)

print(f"\nDone: {len(all_tags)} total tags (+{new_count} new) from {len(pdfs)} PS5 PDFs")
print("Saved asset_tag_mapping.json")
