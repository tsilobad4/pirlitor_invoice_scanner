from openpyxl import load_workbook
from copy import copy

# Constants
MONTH_LABELS = ["JAN","FEB","MARCH","APRIL","MAY","JUNE","JUL","AUG","SEP","OCT","NOV","DEC"]
months_to_add = ["SEP","OCT","NOV","DEC"]
REF = "test_invoice_costing_1.xlsx"
OUT = "test_invoice_costing_FINAL_INSERT_ONLY.xlsx"

# Load workbook and worksheet
wb = load_workbook(REF)
ws = wb["Sorting by part number"]

# Get formatting reference heights
header_h, normal_h = 15, 15
for r in range(1, ws.max_row+1):
    if ws.cell(r,2).value == "Data" and ws.cell(r,3).value == "Part":
        header_h = ws.row_dimensions[r].height or 15
        normal_h = ws.row_dimensions[r+1].height or 15
        break

# Save banner heights
banner_h = {}
for r in range(1, ws.max_row+1):
    if ws.cell(r,2).value == "PART NUMBER":
        banner_h[r] = ws.row_dimensions[r].height or 15

# Insert missing months (NO math, NO formulas)
tpl_row = 5
blocks, row = 0, 1

def copy_style(src, tgt):
    for c in range(1, ws.max_column + 1):
        s = ws.cell(src, c)
        t = ws.cell(tgt, c)
        t.font = copy(s.font)
        t.border = copy(s.border)
        t.fill = copy(s.fill)
        t.number_format = copy(s.number_format)
        t.protection = copy(s.protection)
        t.alignment = copy(s.alignment)
    ws.row_dimensions[tgt].height = ws.row_dimensions[src].height or 15

while row <= ws.max_row and blocks < 50:
    if ws.cell(row,2).value == "PART NUMBER":
        blocks += 1
        start = row + 3

        # find AUG
        aug = None
        r = start
        while r <= ws.max_row:
            if ws.cell(r,2).value == "PART NUMBER": break
            if ws.cell(r,1).value == "AUG": aug = r
            r += 1
        if not aug:
            row += 1
            continue

        # find end of block
        end = aug
        r = aug + 1
        while r <= ws.max_row:
            if ws.cell(r,2).value == "PART NUMBER": break
            a = ws.cell(r,1).value
            b = ws.cell(r,2).value
            if a or b:
                end = r
            else:
                break
            r += 1

        # detect existing months
        existing = {
            ws.cell(i,1).value.strip().upper()
            for i in range(start, end+1)
            if isinstance(ws.cell(i,1).value, str)
            and ws.cell(i,1).value.strip().upper() in MONTH_LABELS
        }

        # insert missing months with formatting only
        ins = end + 1
        for m in months_to_add:
            if m not in existing:
                ws.insert_rows(ins)
                copy_style(tpl_row, ins)
                ws.cell(ins,1).value = m
                # Leave all other columns blank
                ins += 1

        row = ins + 1
    else:
        row += 1

# Normalize row heights
def is_data(r): return ws.cell(r,2).value == "Data" and ws.cell(r,3).value == "Part"

for r in range(1, ws.max_row+1):
    if r in banner_h:
        ws.row_dimensions[r].height = banner_h[r]
    elif is_data(r):
        ws.row_dimensions[r].height = header_h
    else:
        ws.row_dimensions[r].height = normal_h

# Save
wb.save(OUT)
print(f"✅ Months inserted (NO formulas, NO math). Saved to {OUT}")
