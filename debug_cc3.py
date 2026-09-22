import fitz, pytesseract, re
from PIL import Image, ImageEnhance, ImageOps

pdf_path = r'C:\Users\mylap\Downloads\rfi\CPP-RFI-68-64-12-0094_signed by CPY.pdf'
doc = fitz.open(pdf_path)

for pi in [2, 3]:
    page = doc[pi]
    # High DPI
    mat = fitz.Matrix(600/72, 600/72)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
    w, h = img.size
    print(f"\n=== Page {pi+1} ({w}x{h}) ===")

    # Try full page with high contrast
    gray = ImageOps.grayscale(img)
    enh = ImageEnhance.Contrast(gray).enhance(3.0)
    bw = enh.point(lambda x: 0 if x < 130 else 255, '1').convert('L')
    ot = pytesseract.image_to_string(bw, config='--psm 6')
    
    # Look for any resistance-like values
    for ln in ot.split('\n'):
        ln_s = ln.strip()
        if ln_s and (re.search(r'>?\s*\d', ln_s) or 'core' in ln_s.lower() or re.search(r'[gGmMkK]\b', ln_s)):
            print(f"  '{ln_s}'")

    # Also try raw at 600 DPI
    ot2 = pytesseract.image_to_string(img, config='--psm 6')
    print("\n  --- RAW 600DPI ---")
    for ln in ot2.split('\n'):
        ln_s = ln.strip()
        if ln_s and (re.search(r'>', ln_s) or 'core' in ln_s.lower() or re.search(r'\d+\s*[gGmMkK]', ln_s)):
            print(f"  '{ln_s}'")

    # Try to find form fields
    try:
        page_fz = doc[pi]
        widgets = page_fz.widgets()
        if widgets:
            print(f"\n  --- Form fields: ---")
            for w in widgets:
                print(f"    field: {w.field_name} = {w.field_value}")
        else:
            print("\n  No form fields found")
    except:
        print("\n  No form fields")

doc.close()
