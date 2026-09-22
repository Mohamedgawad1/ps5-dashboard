import fitz, pytesseract, re
from PIL import Image

pdf_path = r'C:\Users\mylap\Downloads\rfi\CPP-RFI-68-64-12-0094_signed by CPY.pdf'
doc = fitz.open(pdf_path)

# Page 4 (index 3) - PS5-25-PIT-1102-CJ01
page = doc[3]

# Try 400 DPI for better OCR
mat = fitz.Matrix(400/72, 400/72)
pix = page.get_pixmap(matrix=mat)
img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
w, h = img.size
print(f"Page size at 400 DPI: {w}x{h}")

# Also try rendering the page as a high-quality image and save it for inspection
img.save(r'C:\Users\mylap\OneDrive\Desktop\dashboard\page4_debug.png')
print("Saved page4_debug.png")

# Very narrow strips at the data row area (around 45-55% of page height)
# The table data row should be right below the headers
for y_pct in range(42, 58, 2):
    y1 = int(h * y_pct / 100)
    y2 = int(h * (y_pct + 3) / 100)
    # Full width strip
    crop = img.crop((0, y1, w, y2))
    ct = pytesseract.image_to_string(crop, config='--psm 6').strip()
    if ct:
        # Show only meaningful lines
        for ln in ct.split('\n'):
            ln = ln.strip()
            if ln and len(ln) > 2:
                print(f"  y={y_pct}%: '{ln}'")

# Also try the whole middle section at high DPI
print("\n--- Full middle section OCR ---")
mid = img.crop((0, int(h*0.35), w, int(h*0.65)))
full_mid = pytesseract.image_to_string(mid, config='--psm 6')
print(full_mid)

doc.close()
