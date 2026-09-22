"""Fast unified PDF scanner - scans ALL PDFs in WIRING-MASTER (root + PS5 subfolder)."""
import re, os, json, sys

try:
    import fitz
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WIRING_DIR = os.path.join(BASE_DIR, 'WIRING - MASTER')
PS5_DIR = os.path.join(WIRING_DIR, 'PS5')
MAPPING_FILE = os.path.join(BASE_DIR, 'asset_tag_mapping.json')

TAG_PATTERN = re.compile(r'PS5-[\w]+-[\w]+-[\w]+(?:-[\w]+)?')

def scan_pdf(fpath):
    tags = set()
    if HAS_FITZ:
        try:
            doc = fitz.open(fpath)
            for page in doc:
                text = page.get_text()
                found = TAG_PATTERN.findall(text)
                tags.update(t.strip().replace(' ', '') for t in found)
            doc.close()
        except:
            pass
    else:
        # Fallback: scan binary content
        try:
            with open(fpath, 'rb') as f:
                content = f.read()
            # Extract text-like content
            text = content.decode('latin-1', errors='ignore')
            found = TAG_PATTERN.findall(text)
            tags.update(t.strip().replace(' ', '') for t in found)
        except:
            pass
    return tags

def main():
    # Load existing mapping
    existing = {}
    if os.path.exists(MAPPING_FILE):
        with open(MAPPING_FILE, 'r', encoding='utf-8') as f:
            existing = json.load(f)
    print(f"Existing mapping: {len(existing)} tags")

    # Collect all PDFs
    pdfs = []
    for f in os.listdir(WIRING_DIR):
        if f.lower().endswith('.pdf') and os.path.isfile(os.path.join(WIRING_DIR, f)):
            pdfs.append((os.path.join(WIRING_DIR, f), f))

    for root, dirs, files in os.walk(PS5_DIR):
        for f in files:
            if f.lower().endswith('.pdf'):
                fpath = os.path.join(root, f)
                rel = os.path.relpath(fpath, WIRING_DIR)
                pdfs.append((fpath, rel))

    print(f"Total PDFs to scan: {len(pdfs)}")
    all_tags = dict(existing)
    scanned = 0
    new_tags = 0

    for i, (fpath, fname) in enumerate(pdfs):
        try:
            tags = scan_pdf(fpath)
            for t in tags:
                if t not in all_tags:
                    new_tags += 1
                all_tags[t] = fname
            scanned += 1
            if (i + 1) % 50 == 0 or i == len(pdfs) - 1:
                print(f"  [{i+1}/{len(pdfs)}] scanned={scanned} tags={len(all_tags)} new={new_tags}")
        except Exception as e:
            print(f"  ERROR: {fname}: {e}")

    with open(MAPPING_FILE, 'w', encoding='utf-8') as f:
        json.dump(all_tags, f, indent=2, ensure_ascii=False)

    print(f"\nDONE: {len(all_tags)} total tags ({new_tags} new) from {scanned} PDFs")
    print(f"Saved: {MAPPING_FILE}")

if __name__ == '__main__':
    main()
