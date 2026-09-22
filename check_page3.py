import pdfplumber, fitz

pdf_path = r'C:\Users\mylap\Downloads\rfi\CPP-RFI-68-64-12-0094_signed by CPY.pdf'

# Check page 3 (index 2)
doc = fitz.open(pdf_path)
for i in [1, 2]:
    text = doc[i].get_text()
    print(f"\n=== fitz Page {i+1} (index {i}) ===")
    print(text[:1500] if text else "(empty)")
doc.close()

print("\n\n=== pdfplumber Page 3 ===")
with pdfplumber.open(pdf_path) as pdf:
    t = pdf.pages[2].extract_text()
    print(t[:1500] if t else "(empty)")
