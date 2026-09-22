import fitz, pytesseract, re
from PIL import Image, ImageEnhance, ImageOps

pdf_path = r'C:\Users\mylap\Downloads\rfi\CPP-RFI-68-64-12-0094_signed by CPY.pdf'
doc = fitz.open(pdf_path)

# Page 4 (index 3) - PS5-25-PIT-1102-CJ01 - expected CC = >1G
page = doc[3]
mat = fitz.Matrix(400/72, 400/72)
pix = page.get_pixmap(matrix=mat)
img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
w, h = img.size

# Crop JUST the data row area (around 47-53% of page height, full width)
y1 = int(h * 0.46)
y2 = int(h * 0.54)
strip = img.crop((0, y1, w, y2))

# Try multiple preprocessing approaches
gray = ImageOps.grayscale(strip)

# Approach 1: raw
ot1 = pytesseract.image_to_string(strip, config='--psm 6')
print("=== RAW ===")
for ln in ot1.split('\n'):
    ln = ln.strip()
    if ln:
        print(f"  '{ln}'")

# Approach 2: high contrast + threshold
enh = ImageEnhance.Contrast(gray).enhance(4.0)
bw = enh.point(lambda x: 0 if x < 130 else 255, '1').convert('L')
ot2 = pytesseract.image_to_string(bw, config='--psm 6')
print("\n=== HIGH CONTRAST BW ===")
for ln in ot2.split('\n'):
    ln = ln.strip()
    if ln:
        print(f"  '{ln}'")

# Approach 3: different thresholds
for thr in [80, 100, 150, 200]:
    bw2 = gray.point(lambda x: 0 if x < thr else 255, '1').convert('L')
    ot3 = pytesseract.image_to_string(bw2, config='--psm 6')
    matches = re.findall(r'>?\s*\d+\.?\d*\s*[gGmMkK]', ot3)
    if matches:
        print(f"\n=== threshold={thr} === matches: {matches}")
        for ln in ot3.split('\n'):
            ln = ln.strip()
            if ln and (re.search(r'\d', ln) or '>' in ln):
                print(f"  '{ln}'")

# Approach 4: try psm 7 (single line)
for thr in [100, 130, 160]:
    bw3 = gray.point(lambda x: 0 if x < thr else 255, '1').convert('L')
    ot4 = pytesseract.image_to_string(bw3, config='--psm 7')
    if ot4.strip():
        print(f"\n=== psm7 threshold={thr} === '{ot4.strip()}'")

doc.close()
