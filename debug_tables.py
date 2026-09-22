import pdfplumber

pdf_path = r'C:\Users\mylap\Downloads\rfi\CPP-RFI-68-64-12-0094_signed by CPY.pdf'
with pdfplumber.open(pdf_path) as pdf:
    for pi in [2, 3]:
        page = pdf.pages[pi]
        tables = page.extract_tables()
        print(f"\n=== Page {pi+1} - {len(tables)} tables ===")
        for ti, t in enumerate(tables):
            print(f"\n  Table {ti}:")
            for ri, row in enumerate(t):
                print(f"    Row {ri}: {row}")
