import xlwings as xw

def strip_formulas_from_price_columns(filepath, output_file, max_tables=30):
    wb = xw.Book(filepath)
    ws = wb.sheets["Sorting by part number"]
    data = ws.used_range.value

    MONTHS = ["JAN", "FEB", "MARCH", "APRIL", "MAY", "JUNE", "JUL", "AUG",
              "SEP", "OCT", "NOV", "DEC"]

    row = 0
    tables_processed = 0

    while row < len(data) and tables_processed < max_tables:
        row_data = data[row]
        if isinstance(row_data, list) and "PART NUMBER" in row_data:
            tables_processed += 1
            start_row = row
            # Scan for month rows
            for r in range(start_row + 3, start_row + 20):  # rough window
                if r >= len(data):
                    break
                month_val = data[r][0]
                if isinstance(month_val, str) and month_val.strip().upper() in MONTHS:
                    # Get current values from Unit Price (col 5) and Amount (col 6)
                    for col in [5, 6]:
                        cell_val = data[r][col - 1]
                        if cell_val is not None:
                            ws.cells(r + 1, col).value = cell_val  # xlwings is 1-indexed
            row += 15  # skip ahead assuming ~15 row blocks
        else:
            row += 1

    wb.save(output_file)
    wb.close()
    print(f"✅ Formulas stripped from first {max_tables} tables. Saved as '{output_file}'.")

# Example usage:
strip_formulas_from_price_columns("test_invoice_costing.xlsx", "test_invoice_costing_STRIPPED30.xlsx", max_tables=30)
