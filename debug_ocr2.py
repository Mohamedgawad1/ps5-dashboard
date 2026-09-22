import fitz, pytesseract, re
from PIL import Image

pdf_path = r'C:\Users\mylap\Downloads\rfi\CPP-RFI-68-64-12-0094_signed by CPY.pdf'
doc = fitz.open(pdf_path)

# Page 4 (index 3) - PS5-25-PIT-1102-CJ01 - try different crop areas
page = doc[3]
mat = fitz.Matrix(300/72, 300/72)
pix = page.get_pixmap(matrix=mat)
img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
w, h = img.size
print(f"Page size: {w}x{h}")

# Try many horizontal strips to find where the resistance values are
for y_start_pct in range(30, 70, 5):
    for x_start_pct in [0, 10, 20]:
        y1 = int(h * y_start_pct / 100)
        y2 = int(h * (y_start_pct + 8) / 100)
        x1 = int(w * x_start_pct / 100)
        x2 = int(w * (x_start_pct + 50) / 100)
        crop = img.crop((x1, y1, x2, y2))
        ct = pytesseract.image_to_string(crop, config='--psm 7').strip()
        if ct and len(ct) > 1:
            print(f"  crop y={y_start_pct}%-{y_start_pct+8}% x={x_start_pct}%-{x_start_pct+50}%: '{ct}'")

doc.close()
