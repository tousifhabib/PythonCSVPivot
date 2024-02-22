from typing import Dict, Any
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.worksheet import Worksheet
from pandas import DataFrame
from tableDataProcessing import assign_subtotal_levels


def create_fill(color: str) -> PatternFill:
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def create_font(color: str, bold: bool = False) -> Font:
    return Font(color=color, bold=bold)


def style_cell(cell, fill: PatternFill, font: Font):
    cell.fill = fill
    cell.font = font


def apply_row_styles(worksheet: Worksheet, df: DataFrame, config: Dict[str, Any], dynamic_cols: list):
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

    processed_df = df.copy()
    assign_subtotal_levels(processed_df, dynamic_cols)

    for row_idx, (row, level) in enumerate(zip(worksheet.iter_rows(min_row=2), processed_df['Subtotal_Level']),
                                           start=2):
        row_values = [cell.value for cell in row]
        is_grand_total_row = "Grand Total" in row_values
        is_subtotal_row = level > 1 and not is_grand_total_row

        for cell in row:
            fill, font = (grand_total_fill, grand_total_font) if is_grand_total_row else \
                (subtotal_fill, subtotal_font) if is_subtotal_row else \
                    (default_fill, default_font)
            style_cell(cell, fill, font)


def apply_excel_colors(worksheet: Worksheet, config: Dict[str, Any], df: DataFrame, dynamic_cols: list):
    apply_row_styles(worksheet, df, config, dynamic_cols)
