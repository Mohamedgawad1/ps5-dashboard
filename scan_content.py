"""Fast PDF content scanner with per-file timeout."""
import re, os, json, sys, signal, fitz
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeout

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WIRING_DIR = os.path.join(BASE_DIR, 'WIRING - MASTER')
PS5_DIR = os.path.join(WIRING_DIR, 'PS5')

TAG_RE = re.compile(r'PS5-[\w]+-[\w]+-[\w]+(?:-[\w]+)?')

def extract_tags(fpath):
    tags = set()
    doc = fitz.open(fpath)
    for page in doc:
        text = page.get_text()
        found = TAG_RE.findall(text)
        tags.update(t.strip().replace(' ', '') for t in found)
    doc.close()
    return tags

def scan_pdf(fpath, fname):
    try:
        tags = extract_tags(fpath)
        return fname, tags, None
    except Exception as e:
        return fname, set(), str(e)

def main():
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

    print(f"Scanning {len(pdfs)} PDFs (content-based)...")
    tag_to_pdf = {}
    errors = 0
    scanned = 0

    for i, (fpath, fname) in enumerate(pdfs):
        fname, tags, err = scan_pdf(fpath, fname)
        if err:
            errors += 1
        else:
            for t in tags:
                if t not in tag_to_pdf:
                    tag_to_pdf[t] = fname
            scanned += 1
        if (i + 1) % 50 == 0:
            print(f"  [{i+1}/{len(pdfs)}] tags={len(tag_to_pdf)} scanned={scanned} errors={errors}")

    print(f"\nDONE: {len(tag_to_pdf)} tags from {scanned} PDFs ({errors} errors)")

    with open(os.path.join(BASE_DIR, 'asset_tag_content.json'), 'w', encoding='utf-8') as f:
        json.dump(tag_to_pdf, f, indent=2, ensure_ascii=False)
    print("Saved asset_tag_content.json")

if __name__ == '__main__':
    main()
