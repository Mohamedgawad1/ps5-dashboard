import fitz, pytesseract, re
from PIL import Image

pdf_path = r'C:\Users\mylap\Downloads\rfi\CPP-RFI-68-64-12-0094_signed by CPY.pdf'
doc = fitz.open(pdf_path)

# Test page 3 (index 2) - PS5-25-FIT-0001-CJ01
for pi in [2, 3]:
    page = doc[pi]
    mat = fitz.Matrix(300/72, 300/72)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
    
    # Full page OCR
    full_text = pytesseract.image_to_string(img, config='--psm 6')
    print(f"\n=== Page {pi+1} FULL OCR ===")
    print(full_text)
    
    # Search for resistance values
    lines = full_text.split('\n')
    for idx, ln in enumerate(lines):
        if 'core' in ln.lower() or '>' in ln or re.search(r'\d+\s*[gGmM]', ln):
            print(f"  MATCH line {idx}: '{ln}'")
    
    # Also try crop around the resistance table area
    pt = 300/72
    h = pix.height
    w = pix.width
    # The resistance values should be in the middle-left area of the page
    for crop_name, box in [
        ("top-left", (0, 0, w//2, h//2)),
        ("top-right", (w//2, 0, w, h//2)),
        ("mid", (0, h//4, w, 3*h//4)),
        ("bottom-half", (0, h//2, w, h)),
    ]:
        crop = img.crop(box)
        ct = pytesseract.image_to_string(crop, config='--psm 6')
        if '>' in ct or re.search(r'core.*core', ct, re.IGNORECASE):
            print(f"\n  --- {crop_name} crop has values ---")
            for ln in ct.split('\n'):
                if '>' in ln or 'core' in ln.lower() or re.search(r'\d+\s*[gGmM]', ln):
                    print(f"    '{ln}'")

doc.close()
