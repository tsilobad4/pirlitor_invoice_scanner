import xlwings as xw

def copy_tables_with_formatting(source_path, output_path, max_tables=30):
    app = xw.App(visible=False)
    wb = app.books.open(source_path)
    ws = wb.sheets["Sorting by part number"]

    wb_out = xw.Book()  # New blank workbook
    ws_out = wb_out.sheets[0]
    ws_out.name = "Copied Tables"

    row = 0
    out_row = 1
    table_count = 0

    data = ws.range("A1").expand().value  # Get all values
    while row < len(data) and table_count < max_tables:
        row_data = data[row]
        if isinstance(row_data, list) and "PART NUMBER" in row_data:
            table_count += 1
            # Copy 16 rows for this table
            for i in range(row, row + 16):
                src_range = ws.range(f"A{i+1}:H{i+1}")
                tgt_range = ws_out.range(f"A{out_row}:H{out_row}")
                src_range.copy(tgt_range)
                out_row += 1
            out_row += 2  # extra space between tables
            row += 16
        else:
            row += 1

    wb_out.save(output_path)
    wb_out.close()
    wb.close()
    app.quit()
    print(f"✅ Finished copying {table_count} tables to '{output_path}'")

# Call the function
copy_tables_with_formatting("test_invoice_costing.xlsx", "formatted_tables_only.xlsx", max_tables=30)
