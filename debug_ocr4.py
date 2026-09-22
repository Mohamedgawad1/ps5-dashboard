import fitz, pytesseract, re
from PIL import Image, ImageFilter, ImageOps

pdf_path = r'C:\Users\mylap\Downloads\rfi\CPP-RFI-68-64-12-0094_signed by CPY.pdf'
doc = fitz.open(pdf_path)

# Try page 4 (index 3) - PS5-25-PIT-1102-CJ01
page = doc[3]

# Render at 400 DPI
mat = fitz.Matrix(400/72, 400/72)
pix = page.get_pixmap(matrix=mat)
img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
w, h = img.size

# Convert to grayscale, increase contrast, threshold
gray = ImageOps.grayscale(img)
# Increase contrast
from PIL import ImageEnhance
enhancer = ImageEnhance.Contrast(gray)
gray = enhancer.enhance(3.0)
# Threshold to black/white
gray = gray.point(lambda x: 0 if x < 140 else 255, '1')
gray = gray.convert('L')

# Save for inspection
gray.save(r'C:\Users\mylap\OneDrive\Desktop\dashboard\page4_bw.png')

# The data row should be around 47-53% of page height
# Try very precise crops for each column
# Looking at the PDF structure:
# Column order: CORE NO | COLOR | Core to Core | Core to Shield | Core to Overall Shield | Core to Armor | CONTINUITY | TEST RESULT
# The table area is roughly 42-52% from top

# First, find the exact Y position of the data row by looking for the row below headers
for y_pct in range(44, 56, 1):
    y1 = int(h * y_pct / 100)
    y2 = int(h * (y_pct + 2) / 100)
    strip = gray.crop((0, y1, w, y2))
    st = pytesseract.image_to_string(strip, config='--psm 6').strip()
    if st and len(st) > 3:
        print(f"y={y_pct}%: '{st}'")

# Now try specific column crops for the data row at y=48-50%
print("\n--- Column-by-column at y=47-52% ---")
y1 = int(h * 0.47)
y2 = int(h * 0.52)
# The Core to Core column is roughly 30-40% from left
for x_pct in range(20, 80, 5):
    x1 = int(w * x_pct / 100)
    x2 = int(w * (x_pct + 10) / 100)
    cell = gray.crop((x1, y1, x2, y2))
    ct = pytesseract.image_to_string(cell, config='--psm 7').strip()
    if ct:
        print(f"  x={x_pct}%-{x_pct+10}%: '{ct}'")

doc.close()
