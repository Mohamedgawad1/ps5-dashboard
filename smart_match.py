"""Smart tag matcher - first try base name matching, then targeted PDF scan for unmatched tags."""
import re, os, json, urllib.parse, sys
import fitz

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WIRING_DIR = os.path.join(BASE_DIR, 'WIRING - MASTER')
PS5_DIR = os.path.join(WIRING_DIR, 'PS5')

TAG_PATTERN = re.compile(r'PS5-[\w]+-[\w]+-[\w]+(?:-[\w]+)?')

# Load data
with open(os.path.join(BASE_DIR, 'asset_tag_mapping.json'), 'r', encoding='utf-8') as f:
    mapping = json.load(f)

with open(os.path.join(BASE_DIR, 'data.json'), 'r', encoding='utf-8') as f:
    data = json.load(f)

print(f"Mapping: {len(mapping)} tags")
print(f"Data: {len(data)} assets")

# Step 1: Find unmatched tags and try base-name matching
def get_base(tag):
    return re.sub(r'-(CL|CH|CD|CJ|CS|CT|CC|CB)\d+$', '', tag)

def get_suffix(tag):
    m = re.search(r'-(CL|CH|CD|CJ|CS|CT|CC|CB)(\d+)$', tag)
    return (m.group(1), m.group(2)) if m else (None, None)

unmatched = []
new_links = 0

for item in data:
    if item.get('link'):
        continue
    tag = item['asset_tag']
    # Try exact match
    if tag in mapping:
        pdf = mapping[tag]
        if os.path.exists(os.path.join(WIRING_DIR, pdf.replace('/', os.sep))):
            item['link'] = f'/pdf/{urllib.parse.quote(pdf)}'
            new_links += 1
            continue
    # Try base name (without -CLxx suffix)
    base = get_base(tag)
    if base != tag and base in mapping:
        pdf = mapping[base]
        if os.path.exists(os.path.join(WIRING_DIR, pdf.replace('/', os.sep))):
            item['link'] = f'/pdf/{urllib.parse.quote(pdf)}'
            new_links += 1
            continue
    # Try from/to tags in cable
    if item.get('from') and item['from'] in mapping:
        pdf = mapping[item['from']]
        if os.path.exists(os.path.join(WIRING_DIR, pdf.replace('/', os.sep))):
            item['link'] = f'/pdf/{urllib.parse.quote(pdf)}'
            new_links += 1
            continue
    if item.get('to') and item['to'] in mapping:
        pdf = mapping[item['to']]
        if os.path.exists(os.path.join(WIRING_DIR, pdf.replace('/', os.sep))):
            item['link'] = f'/pdf/{urllib.parse.quote(pdf)}'
            new_links += 1
            continue
    unmatched.append(tag)

print(f"\nStep 1: {new_links} links from base-name matching")
print(f"Still unmatched: {len(unmatched)}")

# Step 2: Scan remaining unmatched tags in PDFs (fast targeted scan)
print(f"\nStep 2: Scanning PDFs for {len(unmatched)} unmatched tags...")
unmatched_set = set(unmatched)

# Build a combined text index per PDF
pdf_texts = {}
all_pdfs = []
for f in os.listdir(WIRING_DIR):
    if f.lower().endswith('.pdf') and os.path.isfile(os.path.join(WIRING_DIR, f)):
        all_pdfs.append((os.path.join(WIRING_DIR, f), f))
for root, dirs, files in os.walk(PS5_DIR):
    for f in files:
        if f.lower().endswith('.pdf'):
            fpath = os.path.join(root, f)
            all_pdfs.append((fpath, os.path.relpath(fpath, WIRING_DIR)))

found_in_pdf = {}
scanned = 0
for fpath, fname in all_pdfs:
    try:
        doc = fitz.open(fpath)
        text = ''
        for page in doc:
            text += page.get_text()
        doc.close()
        # Check which unmatched tags appear in this PDF
        for tag in list(unmatched_set):
            if tag in text:
                found_in_pdf[tag] = fname
                unmatched_set.discard(tag)
        scanned += 1
        if scanned % 50 == 0:
            print(f"  [{scanned}/{len(all_pdfs)}] found={len(found_in_pdf)} remaining={len(unmatched_set)}")
        if not unmatched_set:
            print(f"  All tags found! Stopping at {scanned}/{len(all_pdfs)}")
            break
    except:
        scanned += 1
        continue

print(f"\nStep 2: Found {len(found_in_pdf)} more tags in PDFs")

# Apply new findings
for tag, fname in found_in_pdf.items():
    mapping[tag] = fname
    for item in data:
        if item['asset_tag'] == tag and not item.get('link'):
            pdf = fname.replace('\\', '/')
            item['link'] = f'/pdf/{urllib.parse.quote(pdf)}'
            break
    new_links += 1

# Save mapping
with open(os.path.join(BASE_DIR, 'asset_tag_mapping.json'), 'w', encoding='utf-8') as f:
    json.dump(mapping, f, indent=2, ensure_ascii=False)

# Save data
with open(os.path.join(BASE_DIR, 'data.json'), 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False)
os.system('copy /Y data.json mobile_app\\data.json >nul')

total_links = sum(1 for d in data if d.get('link'))
still_unmatched = len(unmatched_set)
print(f"\nFINAL: {len(data)} assets | {total_links} with PDF link | {still_unmatched} still no link")
print(f"Mapping: {len(mapping)} tags")
