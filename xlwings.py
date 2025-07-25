import xlwings as xw

# Constants
MONTHS = ["JAN", "FEB", "MARCH", "APRIL", "MAY", "JUNE", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

def append_missing_months(filepath):
    wb = xw.Book(filepath)
    ws = wb.sheets["Sorting by part number"]

    used_range = ws.used_range
    values = used_range.value

    row = 0
    while row < len(values):
        row_data = values[row]
        if isinstance(row_data, list) and "PART NUMBER" in row_data:
            part_number = row_data[row_data.index("PART NUMBER") + 1]
            start_row = row
            month_rows = {}

            # Scan forward to collect month rows
            for r in range(start_row + 3, start_row + 20):  # Safe window to capture months
                if r >= len(values):
                    break
                cell_val = values[r][0]
                if isinstance(cell_val, str):
                    clean_val = cell_val.strip().upper()
                    if clean_val in MONTHS:
                        month_rows[clean_val] = r + 1  # Excel is 1-indexed

            missing_months = [m for m in MONTHS if m not in month_rows]

            if missing_months:
                print(f"Inserting missing months for part: {part_number}")
                # Insert below the last existing month
                reference_month = max(month_rows.values()) if month_rows else start_row + 3
                insert_index = reference_month + 1

                # Copy format from AUG if exists, else skip formatting
                template_row = month_rows.get("AUG", None)

                for m in missing_months:
                    ws.api.Rows(insert_index).Insert()
                    ws.range(f"A{insert_index}").value = m

                    if template_row:
                        for col in range(1, ws.used_range.last_cell.column + 1):
                            src = ws.cells(template_row, col)
                            tgt = ws.cells(insert_index, col)
                            tgt.api.Font.Bold = src.api.Font.Bold
                            tgt.api.Font.Name = src.api.Font.Name
                            tgt.api.Font.Size = src.api.Font.Size
                            tgt.color = src.color
                            tgt.number_format = src.number_format

                    insert_index += 1

                row = insert_index
            else:
                row += 15  # Skip to next block if nothing missing
        else:
            row += 1

    wb.save("test_invoice_costing_UPDATED.xlsx")
    wb.close()
