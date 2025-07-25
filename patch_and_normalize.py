from openpyxl import load_workbook
from copy import copy

# Constants
MONTH_LABELS = ["JAN", "FEB", "MARCH", "APRIL", "MAY", "JUNE", "JUL", "AUG",
                "SEP", "OCT", "NOV", "DEC"]
months_to_add = ["SEP", "OCT", "NOV", "DEC"]
TEMPLATE_ROW = 5
REF = "test_invoice_costing.xlsx"  # Input file
OUT = "test_invoice_costing_ALL_FIXED.xlsx"  # Output file

# Load workbook and worksheet
wb = load_workbook(REF)
ws = wb["Sorting by part number"]

# Detect banner + header heights
header_h, normal_h = 15, 15
banner_h = {}
for r in range(1, ws.max_row + 1):
    if ws.cell(r, 2).value == "Data" and ws.cell(r, 3).value == "Part":
        header_h = ws.row_dimensions[r].height or 15
        normal_h = ws.row_dimensions[r + 1].height or 15
    elif ws.cell(r, 2).value == "PART NUMBER":
        banner_h[r] = ws.row_dimensions[r].height or 15

# Copy style helper
def copy_row_format(src, tgt):
    for c in range(1, ws.max_column + 1):
        s, t = ws.cell(src, c), ws.cell(tgt, c)
        if s.has_style:
            t.font = copy(s.font)
            t.border = copy(s.border)
            t.fill = copy(s.fill)
            t.number_format = copy(s.number_format)
            t.protection = copy(s.protection)
            t.alignment = copy(s.alignment)
    ws.row_dimensions[tgt].height = ws.row_dimensions[src].height or 15

# Append missing months with formulas
r = 1
while r <= ws.max_row:
    if ws.cell(r, 2).value == "PART NUMBER":
        start = r + 3

        # Find AUG row
        aug = None
        scan = start
        while scan <= ws.max_row:
            if ws.cell(scan, 2).value == "PART NUMBER":
                break
            if ws.cell(scan, 1).value == "AUG":
                aug = scan
            scan += 1
        if not aug:
            r += 1
            continue

        # Find block end
        end = aug
        scan = aug + 1
        while scan <= ws.max_row:
            if ws.cell(scan, 2).value == "PART NUMBER":
                break
            a, b = ws.cell(scan, 1).value, ws.cell(scan, 2).value
            if a or b:
                end = scan
            else:
                break
            scan += 1

        # Detect existing months
        existing = {
            ws.cell(i, 1).value.strip().upper()
            for i in range(start, end + 1)
            if isinstance(ws.cell(i, 1).value, str)
            and ws.cell(i, 1).value.strip().upper() in MONTH_LABELS
        }

        # Insert missing months with formulas
        ins = end + 1
        for m in months_to_add:
            if m not in existing:
                ws.insert_rows(ins)
                copy_row_format(TEMPLATE_ROW, ins)
                ws.cell(ins, 1).value = m
                ws.cell(ins, 6).value = f"=D{ins}*E{ins}"  # Amount formula
                ws.cell(ins, 7).value = f"=IF(E{ins}=0, 0, F{ins}/E{ins})"  # Unit Price / Lot
                ins += 1

        r = ins + 1
    else:
        r += 1

# Normalize heights
def is_data_row(idx):
    return ws.cell(idx, 2).value == "Data" and ws.cell(idx, 3).value == "Part"

for idx in range(1, ws.max_row + 1):
    if idx in banner_h:
        ws.row_dimensions[idx].height = banner_h[idx]
    elif is_data_row(idx):
        ws.row_dimensions[idx].height = header_h
    else:
        ws.row_dimensions[idx].height = normal_h

# Save
wb.save(OUT)
print(f"✅ Saved to {OUT}")
