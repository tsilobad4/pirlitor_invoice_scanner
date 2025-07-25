from openpyxl import load_workbook, Workbook

# Constants
SOURCE_FILE = "test_invoice_costing.xlsx"
OUTPUT_FILE = "copied_values_only.xlsx"
SHEET_NAME = "Sorting by part number"
MAX_TABLES = 30

# Load source workbook twice — once to get static values
wb_formulas = load_workbook(SOURCE_FILE)
wb_values = load_workbook(SOURCE_FILE, data_only=True)

ws_f = wb_formulas[SHEET_NAME]
ws_v = wb_values[SHEET_NAME]

# Create new workbook for the output
wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = "Copied Tables"

out_row = 1
table_count = 0
row = 1

while row <= ws_v.max_row and table_count < MAX_TABLES:
    if ws_v.cell(row, 2).value == "PART NUMBER":
        table_count += 1

        # Copy 1 header row + 1 spacer + 1 column header + up to 12 month rows
        for i in range(row, row + 16):
            for col in range(1, ws_v.max_column + 1):
                val = ws_v.cell(i, col).value
                ws_out.cell(out_row, col).value = val
            out_row += 1

        out_row += 1  # Leave a gap between tables
        row += 16     # Move to next part table
    else:
        row += 1

# Save the clean version
wb_out.save(OUTPUT_FILE)
print(f"✅ Copied first {table_count} tables to '{OUTPUT_FILE}' (no formulas)")
