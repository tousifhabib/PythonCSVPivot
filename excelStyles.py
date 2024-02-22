from typing import Dict, Any, List
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.worksheet import Worksheet
from pandas import DataFrame


def create_fill(color: str) -> PatternFill:
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def create_font(color: str, bold: bool = False) -> Font:
    return Font(color=color, bold=bold)


def style_cell(cell, fill: PatternFill, font: Font):
    cell.fill = fill
    cell.font = font


def is_special_row_excel(row_values: List, dynamic_cols: list) -> bool:
    # Check for explicit "Subtotal" or "Grand Total" markers
    if "Subtotal" in row_values or "Grand Total" in row_values:
        return True

    for i, col_name in enumerate(dynamic_cols):
        if i < len(row_values) - 1:
            if row_values[i] and not row_values[i + 1]:
                return True

    return False


def apply_row_styles(worksheet: Worksheet, config: Dict[str, Any], dynamic_cols: list):
    header_fill = create_fill(config['colors']['header']['background'])
    header_font = create_font(config['colors']['header']['text'], bold=True)

    for cell in worksheet[1]:
        style_cell(cell, header_fill, header_font)

    default_fill = create_fill(config['colors']['default']['background'])
    default_font = create_font(config['colors']['default']['text'])
    subtotal_fill = create_fill(config['colors']['subtotal']['background'])
    subtotal_font = create_font(config['colors']['subtotal']['text'])
    grand_total_fill = create_fill(config['colors']['grand_total']['background'])
    grand_total_font = create_font(config['colors']['grand_total']['text'])

    for row_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
        row_values = [cell.value for cell in row]
        is_grand_total_row = "Grand Total" in row_values
        is_subtotal_row = is_special_row_excel(row_values, dynamic_cols) and not is_grand_total_row

        for cell in row:
            if is_grand_total_row:
                fill, font = grand_total_fill, grand_total_font
            elif is_subtotal_row:
                fill, font = subtotal_fill, subtotal_font
            else:
                fill, font = default_fill, default_font
            style_cell(cell, fill, font)


def apply_excel_colors(worksheet: Worksheet, config: Dict[str, Any], df: DataFrame, dynamic_cols: list):
    apply_row_styles(worksheet, config, dynamic_cols)
