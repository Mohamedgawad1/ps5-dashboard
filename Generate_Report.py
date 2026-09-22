import os, sys, json, re, glob
import pandas as pd
import fitz  # PyMuPDF
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import openpyxl

# ============================================================
# CONFIGURATION - Change these paths if needed
# ============================================================
FOLDER_PATH = os.path.dirname(os.path.abspath(__file__))
# FOLDER_PATH = r"C:\Users\mylap\OneDrive\Desktop\dashboard"  # or set manually

OV_TASKS_FILE = os.path.join(FOLDER_PATH, "ovTasks_TestsPlanned_1369.xlsx")
MASTER_TRACKER_FILE = os.path.join(FOLDER_PATH, "PS5 Master tracker EIT Combined.xlsx")
PDF_FOLDER = FOLDER_PATH  # PDFs location (or change to subfolder)
OUTPUT_FILE = os.path.join(FOLDER_PATH, "Asset_to_PDF_Report.xlsx")

# ============================================================
# STEP 1: Load assets from ovTasks Excel
# ============================================================
print("=" * 60)
print("ASSET TO PDF MAPPING REPORT GENERATOR")
print("=" * 60)
print(f"\n[1/5] Loading assets from: {os.path.basename(OV_TASKS_FILE)}")

df = pd.read_excel(OV_TASKS_FILE, sheet_name=0, engine='openpyxl')
# Filter: only Closed tasks (Closed Punch Tasks)
df_closed = df[df['Task State'] == 'Closed'].copy()
assets_raw = df_closed['Asset - Tag'].dropna().str.strip().unique()
assets = sorted(set(str(a) for a in assets_raw))
total_all = df['Asset - Tag'].nunique()
print(f"  -> {len(assets)} unique assets (Closed tasks only, out of {total_all} total)")

# ============================================================
# STEP 2: Scan all PDFs in the folder
# ============================================================
print(f"\n[2/5] Scanning PDFs in: {FOLDER_PATH}")

pdf_files = sorted(glob.glob(os.path.join(FOLDER_PATH, "*.pdf")))
print(f"  -> Found {len(pdf_files)} PDF files")

# Build regex patterns in batches to avoid huge regex
assets_sorted = sorted(assets, key=len, reverse=True)
escaped = [re.escape(a) for a in assets_sorted]
BATCH_SIZE = 2000
batches = [escaped[i:i+BATCH_SIZE] for i in range(0, len(escaped), BATCH_SIZE)]
print(f"  -> Split into {len(batches)} search batches")

all_results = {}
total_pages = 0

for pdf_path in pdf_files:
    pdf_name = os.path.basename(pdf_path)
    pdf_display = pdf_name.encode('ascii', 'replace').decode('ascii')
    sys.stdout.write(f"  Processing: {pdf_display}... ")
    sys.stdout.flush()
    
    try:
        doc = fitz.open(pdf_path)
        page_count = len(doc)
        total_pages += page_count
        
        # Extract text from all pages
        page_texts = {}
        for page_num in range(page_count):
            page = doc[page_num]
            page_texts[page_num + 1] = page.get_text()
        doc.close()
        
        # Search assets in batches
        found_in_this_pdf = 0
        for batch in batches:
            pattern = '|'.join(batch)
            regex = re.compile(pattern)
            
            for page_num, text in page_texts.items():
                for match in regex.finditer(text):
                    asset = match.group()
                    if asset not in all_results:
                        all_results[asset] = {}
                    if pdf_name not in all_results[asset]:
                        all_results[asset][pdf_name] = []
                    if page_num not in all_results[asset][pdf_name]:
                        all_results[asset][pdf_name].append(page_num)
                        found_in_this_pdf += 1
        
        print(f"{found_in_this_pdf} matches")
    except Exception as e:
        err = str(e).encode('ascii', 'replace').decode('ascii')
        print(f"ERROR: {err}")

print(f"\n  -> Total: {len(all_results)} assets found across {total_pages} PDF pages")

# ============================================================
# STEP 3: Load cable & punch data from Master Tracker
# ============================================================
print(f"\n[3/5] Loading cable & punch data from Master Tracker...")

cable_data = {}
punch_data = {}  # cable_tag -> {remarks, punch_desc, rfi_status}

try:
    wb_master = openpyxl.load_workbook(MASTER_TRACKER_FILE, data_only=True)
    
    # Sheet 4: Trace Heating Cable Schedule (index 3)
    ws4 = wb_master.worksheets[3]
    for row in range(3, ws4.max_row + 1):
        cable_tag = ws4.cell(row, 2).value
        if cable_tag:
            ct = str(cable_tag).strip()
            cable_data[ct] = {
                "source": str(ws4.cell(row, 11).value or "").strip(),
                "dest": str(ws4.cell(row, 18).value or "").strip(),
                "rfi_laying": str(ws4.cell(row, 31).value or "").strip(),
                "rfi_test": str(ws4.cell(row, 34).value or "").strip(),
                "rfi_cg": str(ws4.cell(row, 37).value or "").strip(),
                "rfi_lay_status": str(ws4.cell(row, 33).value or "").strip(),
                "rfi_test_status": str(ws4.cell(row, 36).value or "").strip(),
                "punch_remarks": str(ws4.cell(row, 28).value or "").strip(),
                "punch_desc": str(ws4.cell(row, 40).value or "").strip()
            }
    
    # Sheet 8: Trace Heating Groups (index 7) - has Punch A,B,C
    ws8 = wb_master.worksheets[7]
    current_asset = ""
    for row in range(3, ws8.max_row + 1):
        # Asset group header
        sr = ws8.cell(row, 1).value
        if sr and isinstance(sr, str) and sr.startswith("Asset:"):
            current_asset = sr.replace("Asset:", "").strip()
        
        cable_tag = ws8.cell(row, 2).value
        if cable_tag and str(cable_tag).strip():
            ct = str(cable_tag).strip()
            remarks = str(ws8.cell(row, 28).value or "").strip()
            punch_desc = str(ws8.cell(row, 40).value or "").strip()
            rfi_lay = str(ws8.cell(row, 31).value or "").strip()
            rfi_test = str(ws8.cell(row, 34).value or "").strip()
            rfi_lay_status = str(ws8.cell(row, 33).value or "").strip()
            rfi_test_status = str(ws8.cell(row, 36).value or "").strip()
            
            punch_data[ct] = {
                "asset_group": current_asset,
                "remarks": remarks,
                "punch_desc": punch_desc,
                "rfi_laying": rfi_lay,
                "rfi_test": rfi_test,
                "rfi_lay_status": rfi_lay_status,
                "rfi_test_status": rfi_test_status
            }
            
            # Parse A=0 B=2 C=0 format
            punch_counts = {"A": "0", "B": "0", "C": "0"}
            m = re.search(r'A=(\d+)\s+B=(\d+)\s+C=(\d+)', remarks)
            if m:
                punch_counts = {"A": m.group(1), "B": m.group(2), "C": m.group(3)}
            punch_data[ct]["punch_a"] = punch_counts["A"]
            punch_data[ct]["punch_b"] = punch_counts["B"]
            punch_data[ct]["punch_c"] = punch_counts["C"]
    
    wb_master.close()
    print(f"  -> {len(cable_data)} cables loaded from Sheet 4")
    print(f"  -> {len(punch_data)} cables with punch data from Sheet 8")
except Exception as e:
    print(f"  -> Warning: Could not load cable/punch data: {e}")

# ============================================================
# STEP 4: Find which assets are heat-tracing related in isometric PDF
# ============================================================
print(f"\n[4/5] Cross-referencing with cable schedule...")

# For each found asset, try to find related cable info
asset_cable_info = {}
for asset in all_results:
    if asset in cable_data:
        asset_cable_info[asset] = cable_data[asset]
    else:
        for ct, ci in cable_data.items():
            if asset in ci["source"] or asset in ci["dest"]:
                if asset not in asset_cable_info:
                    asset_cable_info[asset] = {"source": "", "dest": "", "rfi_laying": "", "rfi_test": "", "rfi_cg": "", "rfi_lay_status": "", "rfi_test_status": "", "punch_remarks": "", "punch_desc": ""}
                for k in ci:
                    if ci[k]:
                        asset_cable_info[asset][k] = ci[k]
    
    # Also merge punch data from Sheet 8
    if asset not in asset_cable_info:
        asset_cable_info[asset] = {}
    if asset in punch_data:
        asset_cable_info[asset].update(punch_data[asset])

print(f"  -> Cable/punch info found for {len(asset_cable_info)} assets")

# ============================================================
# STEP 5: Create formatted Excel report
# ============================================================
print(f"\n[5/5] Creating Excel report...")

wb = openpyxl.Workbook()

# Styles
header_font = Font(bold=True, color="FFFFFF", size=11)
header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
title_font = Font(bold=True, size=14, color="2F5496")
subtitle_font = Font(italic=True, size=10, color="666666")
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
wrap_align = Alignment(wrap_text=True, vertical='center')
center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
alt_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")

# ----- SHEET 1: Overview -----
ws = wb.active
ws.title = "Overview"

ws.merge_cells('A1:D1')
ws['A1'] = "Asset-to-PDF Mapping Report"
ws['A1'].font = title_font
ws['A1'].alignment = Alignment(horizontal='center')

ws.merge_cells('A2:D2')
ws['A2'] = f"Generated on: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')} | Source: {os.path.basename(OV_TASKS_FILE)}"
ws['A2'].font = subtitle_font
ws['A2'].alignment = Alignment(horizontal='center')

# Summary
row = 4
ws.cell(row, 1, "SUMMARY").font = Font(bold=True, size=12, color="2F5496")
row = 5
for c, h in enumerate(["Metric", "Value"], 1):
    cell = ws.cell(row, c, h)
    cell.font = header_font; cell.fill = header_fill; cell.border = thin_border; cell.alignment = center_align

not_found_count = len(assets) - len(all_results)
for i, (label, val) in enumerate([
    ("Total Unique Assets in ovTasks", len(assets)),
    ("Assets Found in Any PDF", len(all_results)),
    ("Assets NOT Found in Any PDF", not_found_count),
    ("PDF Documents Searched", len(pdf_files)),
    ("Total PDF Pages Scanned", total_pages),
]):
    r = row + 1 + i
    ws.cell(r, 1, label).border = thin_border
    c = ws.cell(r, 2, val)
    c.border = thin_border; c.alignment = center_align
    if i == 2:
        c.font = Font(color="FF0000", bold=True) if val > 0 else Font(color="008000")

# Per-PDF breakdown
row = row + len([
    ("Total Unique Assets in ovTasks", len(assets)),
    ("Assets Found in Any PDF", len(all_results)),
    ("Assets NOT Found in Any PDF", not_found_count),
    ("PDF Documents Searched", len(pdf_files)),
    ("Total PDF Pages Scanned", total_pages),
]) + 3
ws.cell(row, 1, "ASSETS FOUND PER PDF").font = Font(bold=True, size=12, color="2F5496")
row += 1
for c, h in enumerate(["PDF File", "Assets Found", "Total Pages"], 1):
    cell = ws.cell(row, c, h)
    cell.font = header_font; cell.fill = header_fill; cell.border = thin_border; cell.alignment = center_align

pdf_asset_count = {}
for asset, pdfs_dict in all_results.items():
    for pdf_name in pdfs_dict:
        pdf_asset_count[pdf_name] = pdf_asset_count.get(pdf_name, 0) + 1

row += 1
for pdf_name, count in sorted(pdf_asset_count.items(), key=lambda x: -x[1]):
    pdf_display = pdf_name.encode('ascii', 'replace').decode('ascii')
    ws.cell(row, 1, pdf_display).border = thin_border
    ws.cell(row, 2, count).border = thin_border
    ws.cell(row, 2).alignment = center_align
    # Get page count
    for pdf_path in pdf_files:
        if os.path.basename(pdf_path) == pdf_name:
            try:
                doc = fitz.open(pdf_path)
                ws.cell(row, 3, len(doc)).border = thin_border
                ws.cell(row, 3).alignment = center_align
                doc.close()
            except:
                ws.cell(row, 3, "?").border = thin_border
            break
    row += 1

ws.column_dimensions['A'].width = 55
ws.column_dimensions['B'].width = 18
ws.column_dimensions['C'].width = 15

# ----- SHEET 2: Detailed Mapping -----
ws2 = wb.create_sheet("Detailed Mapping")

headers = ["Asset Tag", "PDF Document", "Page Numbers", "Source Tag", "Dest Tag", 
           "Punch A", "Punch B", "Punch C", "Remarks / Punch Desc", 
           "RFI Cable Laying", "RFI Lay Status", "RFI Test", "RFI Test Status"]
for c, h in enumerate(headers, 1):
    cell = ws2.cell(1, c, h)
    cell.font = header_font; cell.fill = header_fill; cell.border = thin_border; cell.alignment = center_align

detail_row = 2
for asset in sorted(all_results.keys()):
    ci = asset_cable_info.get(asset, {})
    for pdf_name, pages in sorted(all_results[asset].items()):
        pages_str = ", ".join(str(p) for p in sorted(set(pages)))
        pdf_short = pdf_name.replace("_signed by CPY", "").replace(".pdf", "")
        pdf_short = pdf_short.encode('ascii', 'replace').decode('ascii')
        
        ws2.cell(detail_row, 1, asset).border = thin_border
        ws2.cell(detail_row, 2, pdf_short).border = thin_border
        ws2.cell(detail_row, 3, pages_str).border = thin_border
        ws2.cell(detail_row, 3).alignment = wrap_align
        ws2.cell(detail_row, 4, ci.get("source", "")).border = thin_border
        ws2.cell(detail_row, 5, ci.get("dest", "")).border = thin_border
        ws2.cell(detail_row, 6, ci.get("punch_a", "")).border = thin_border
        ws2.cell(detail_row, 6).alignment = center_align
        ws2.cell(detail_row, 7, ci.get("punch_b", "")).border = thin_border
        ws2.cell(detail_row, 7).alignment = center_align
        ws2.cell(detail_row, 8, ci.get("punch_c", "")).border = thin_border
        ws2.cell(detail_row, 8).alignment = center_align
        # Combine remarks + punch desc
        punch_text = ci.get("remarks", "")
        if ci.get("punch_desc"):
            if punch_text: punch_text += " | "
            punch_text += ci["punch_desc"]
        ws2.cell(detail_row, 9, punch_text[:200]).border = thin_border
        ws2.cell(detail_row, 9).alignment = wrap_align
        ws2.cell(detail_row, 10, ci.get("rfi_laying", "")).border = thin_border
        ws2.cell(detail_row, 11, ci.get("rfi_lay_status", "")).border = thin_border
        ws2.cell(detail_row, 11).alignment = center_align
        ws2.cell(detail_row, 12, ci.get("rfi_test", "")).border = thin_border
        ws2.cell(detail_row, 13, ci.get("rfi_test_status", "")).border = thin_border
        ws2.cell(detail_row, 13).alignment = center_align
        
        # Color code punch columns
        for pc, col_idx in [("A", 6), ("B", 7), ("C", 8)]:
            val = ci.get(f"punch_{pc.lower()}", "")
            if val and val != "0":
                ws2.cell(detail_row, col_idx).font = Font(bold=True, color="FF0000")
        
        # Color code RFI status
        for status, col_idx in [("OPEN", 11), ("OPEN", 13)]:
            val = ci.get("rfi_lay_status" if col_idx == 11 else "rfi_test_status", "")
            if val and "OPEN" in val.upper():
                ws2.cell(detail_row, col_idx).font = Font(bold=True, color="FF0000")
        
        if detail_row % 2 == 0:
            for col in range(1, 14):
                ws2.cell(detail_row, col).fill = alt_fill
        
        detail_row += 1

ws2.column_dimensions['A'].width = 40
ws2.column_dimensions['B'].width = 35
ws2.column_dimensions['C'].width = 18
ws2.column_dimensions['D'].width = 28
ws2.column_dimensions['E'].width = 28
ws2.column_dimensions['F'].width = 9
ws2.column_dimensions['G'].width = 9
ws2.column_dimensions['H'].width = 9
ws2.column_dimensions['I'].width = 45
ws2.column_dimensions['J'].width = 28
ws2.column_dimensions['K'].width = 14
ws2.column_dimensions['L'].width = 28
ws2.column_dimensions['M'].width = 14
ws2.freeze_panes = 'A2'
ws2.auto_filter.ref = f"A1:M{detail_row-1}"

# ----- SHEET 3: Not Found Assets -----
ws3 = wb.create_sheet("Not Found")
ws3.cell(1, 1, "Asset Tag (not found in any PDF)").font = header_font
ws3.cell(1, 1).fill = header_fill
ws3.cell(1, 1).border = thin_border
ws3.column_dimensions['A'].width = 65

not_found_list = sorted(set(assets) - set(all_results.keys()))
for i, asset in enumerate(not_found_list):
    r = i + 2
    ws3.cell(r, 1, asset).border = thin_border
    if r % 2 == 0:
        ws3.cell(r, 1).fill = alt_fill

ws3.freeze_panes = 'A2'

# Save
wb.save(OUTPUT_FILE)
print(f"\n{'=' * 60}")
print(f"REPORT GENERATED SUCCESSFULLY!")
print(f"{'=' * 60}")
print(f"Output: {OUTPUT_FILE}")
print(f"Sheets: Overview | Detailed Mapping | Not Found")
print(f"Found: {len(all_results)} | Not Found: {len(not_found_list)}")
print(f"\nTo regenerate, simply run this script again:")
print(f"  python \"{os.path.abspath(__file__)}\"")