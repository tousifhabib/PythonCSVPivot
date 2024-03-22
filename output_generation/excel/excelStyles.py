import logging
from typing import Dict, Any, List, Tuple, Set
import pandas as pd
from openpyxl.styles import PatternFill, Font
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet
from reportlab.lib import units
from reportlab.lib.pagesizes import letter

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def calculate_dynamic_page_size(table_width, table_height, config, paginate, num_rows):
    default_cell_height = 0.25 * units.inch
    header_footer_height = 0.55 * units.inch

    margins = config.get("margins", {"top": 1, "bottom": 1, "left": 1, "right": 1})

    estimated_content_height = num_rows * default_cell_height + header_footer_height

    total_width = table_width + (margins["left"] + margins["right"]) * units.inch
    total_height = estimated_content_height + (margins["top"] + margins["bottom"]) * units.inch

    min_page_size = config.get("min_page_size", letter)
    max_page_size = config.get("max_page_size", (letter[0] * 2, letter[1] * 2))

    if paginate:
        return (
            max(min(total_width, max_page_size[0]), min_page_size[0]),
            max(min(total_height, max_page_size[1]), min_page_size[1])
        )
    else:
        new_width = max(min(total_width, max_page_size[0]), min_page_size[0])
        new_height = max(total_height, min_page_size[1])
        return new_width, new_height


def create_fill(color: str) -> PatternFill:
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def create_font(color: str, bold: bool = False) -> Font:
    return Font(color=color, bold=bold)


def style_cell(cell: Cell, fill: PatternFill, font: Font) -> None:
    cell.fill = fill
    cell.font = font


def get_style_mappings(config: Dict[str, Any]) -> Dict[str, Tuple[PatternFill, Font]]:
    return {key: (create_fill(value['background']), create_font(value['text'], key == 'header'))
            for key, value in config['colors'].items()}


def determine_row_style(row_values: List[Any], style_mappings: Dict[str, Tuple[PatternFill, Font]]) -> Tuple[
    Tuple[PatternFill, Font], int]:
    if "Grand Total" in row_values:
        return style_mappings['grand_total'], 1
    if any(keyword in row_values for keyword in ["Grand Total"]):
        return style_mappings['grand_total'], 1
    level = next((i for i, value in enumerate(row_values[:-2], 1) if value), 0)
    if level:
        return style_mappings.get(f'subtotal_{level}', style_mappings['default']), level
    return style_mappings['default'], 1


def is_special_row(row_values: List[Any]) -> bool:
    if any(keyword in row_values for keyword in ["Grand Total"]):
        return True
    return any(cell and (i == 1 or not row_values[i - 1]) for i, cell in enumerate(row_values[1:], start=1))


def apply_row_styles(worksheet: Worksheet, config: Dict[str, Any]) -> None:
    style_mappings = get_style_mappings(config)
    for row_index, row in enumerate(worksheet.iter_rows(), start=1):
        row_values = [cell.value for cell in row]
        if row_index == 1:
            fill, font = style_mappings['header']
            start_col = 1
        else:
            (fill, font), start_col = determine_row_style(row_values, style_mappings)
        for cell in row[start_col - 1:]:
            style_cell(cell, fill, font)


def apply_excel_colors(worksheet: Worksheet, config: Dict[str, Any]) -> None:
    apply_row_styles(worksheet, config)


def merge_empty_cells(worksheet, df: pd.DataFrame) -> None:
    print(df)
    for col_idx, col in enumerate(df.columns, start=1):
        merge_column(worksheet, df, col_idx)


def merge_column(worksheet, df: pd.DataFrame, col_idx: int) -> None:
    start_row = None
    prev_row_was_special = False  # Add a flag to track if the previous row was a special row

    for row_idx, row_data in enumerate(df.values, start=2):
        current_row_is_special = is_special_row1(df, row_idx)
        next_row_is_special = is_special_row1(df, row_idx + 1) if row_idx < len(df) else False

        if start_row is not None and (current_row_is_special or next_row_is_special):
            merge_cells(worksheet, col_idx, start_row, row_idx - 1)
            start_row = None

        if current_row_is_special:
            prev_row_was_special = True
            continue
        else:
            prev_row_was_special = False

        if should_span(row_data[col_idx - 1]):
            if start_row is None and not prev_row_was_special:
                start_row = row_idx
        else:
            if start_row is not None:
                merge_cells(worksheet, col_idx, start_row, row_idx - 1)
                start_row = None

    if start_row is not None and not prev_row_was_special:
        merge_cells(worksheet, col_idx, start_row, len(df) + 1)


def is_special_row1(df: pd.DataFrame, row_idx: int) -> bool:
    if row_idx - 1 < len(df):
        row_data = df.iloc[row_idx - 2]
        is_total = "Grand Total" in str(row_data.iloc[0])
        if not is_total and not pd.isna(row_data.iloc[0]) and all(pd.isna(val) for val in row_data.iloc[1:]):
            return True
        return is_total
    return False


def should_span(cell_value) -> bool:
    return pd.isna(cell_value) or cell_value == ""


def merge_cells(worksheet, col_idx: int, start_row: int, end_row: int) -> None:
    if start_row is not None and end_row - start_row > 1:
        worksheet.merge_cells(start_row=start_row, start_column=col_idx, end_row=end_row, end_column=col_idx)
