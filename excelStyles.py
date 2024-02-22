from openpyxl.styles import PatternFill, Font
from tableDataProcessing import assign_subtotal_levels


def apply_excel_colors(worksheet, config, df, dynamic_cols):
    header_fill = PatternFill(start_color=config['colors']['header']['background'],
                              end_color=config['colors']['header']['background'],
                              fill_type="solid")
    header_font = Font(color=config['colors']['header']['text'], bold=True)
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font

    default_fill = PatternFill(start_color=config['colors']['default']['background'],
                               end_color=config['colors']['default']['background'],
                               fill_type="solid")
    default_font = Font(color=config['colors']['default']['text'])

    subtotal_fill = PatternFill(start_color=config['colors']['subtotal']['background'],
                                end_color=config['colors']['subtotal']['background'],
                                fill_type="solid")
    subtotal_font = Font(color=config['colors']['subtotal']['text'])

    grand_total_fill = PatternFill(start_color=config['colors']['grand_total']['background'],
                                   end_color=config['colors']['grand_total']['background'],
                                   fill_type="solid")
    grand_total_font = Font(color=config['colors']['grand_total']['text'])

    processed_df = df.copy()
    assign_subtotal_levels(processed_df, dynamic_cols)

    for row_idx, (row, level) in enumerate(zip(worksheet.iter_rows(min_row=2), processed_df['Subtotal_Level']),
                                           start=2):
        row_values = [cell.value for cell in row]
        is_grand_total_row = "Grand Total" in row_values
        is_subtotal_row = level > 1 and not is_grand_total_row

        for cell in row:
            if is_grand_total_row:
                cell.fill = grand_total_fill
                cell.font = grand_total_font
            elif is_subtotal_row:
                cell.fill = subtotal_fill
                cell.font = subtotal_font
            else:
                cell.fill = default_fill
                cell.font = default_font
