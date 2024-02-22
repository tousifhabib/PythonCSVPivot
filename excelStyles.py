from openpyxl.styles import PatternFill, Font, Alignment

from openpyxl.styles import PatternFill, Font


def apply_excel_colors(worksheet, config):
    # Apply header styles
    header_fill = PatternFill(start_color=config['colors']['header']['background'],
                              end_color=config['colors']['header']['background'],
                              fill_type="solid")
    header_font = Font(color=config['colors']['header']['text'], bold=True)

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font

    # Define fills and fonts for default, subtotal, and grand total rows
    default_fill = PatternFill(start_color=config['colors']['default']['background'],
                               end_color=config['colors']['default']['background'],
                               fill_type="solid")
    default_font = Font(color=config['colors']['default']['text'])

    subtotal_fill = PatternFill(start_color=config['colors']['subtotal']['background'],
                                end_color=config['colors']['subtotal']['background'],
                                fill_type="solid")
    subtotal_font = Font(color=config['colors']['subtotal']['text'])

    for row_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
        row_values = [cell.value for cell in row]
        print(row_values)
        is_subtotal_row = any("Subtotal" in str(value) for value in row_values)
        is_grand_total_row = any("Grand Total" in str(value) for value in row_values)

        for cell in row:
            if is_subtotal_row or is_grand_total_row:
                cell.fill = subtotal_fill
                cell.font = subtotal_font
            else:
                cell.fill = default_fill
                cell.font = default_font
