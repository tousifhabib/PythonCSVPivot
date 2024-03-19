from typing import Dict, Any, List, Tuple, Set
import pandas as pd
from openpyxl.styles import PatternFill, Font
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet
from reportlab.lib import units
from reportlab.lib.pagesizes import letter


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
    subtotal_rows, subtotal_levels = identify_subtotal_rows(df)
    merge_first_column(worksheet, subtotal_rows, subtotal_levels)
    merge_remaining_columns(worksheet, df, subtotal_rows, subtotal_levels)


def identify_subtotal_rows(df: pd.DataFrame) -> Tuple[Set[int], Dict[int, int]]:
    subtotal_rows = set()
    subtotal_levels = {}
    for col_idx, col in enumerate(df.columns[:-2], start=1):
        blank_seen = False
        level = 0
        for row_idx, value in enumerate(df[col], start=2):
            if pd.isna(value) or value == '':
                blank_seen = True
            else:
                if blank_seen or "Grand Total" in str(value):
                    subtotal_rows.add(row_idx)
                    if col_idx == 1:
                        subtotal_levels[row_idx] = level
                    blank_seen = False
                    level += 1
    return subtotal_rows, subtotal_levels


def merge_first_column(worksheet, subtotal_rows: Set[int], subtotal_levels: Dict[int, int]) -> None:
    row_idx_list = sorted(subtotal_rows)
    start_row = 3
    prev_level = -1
    for row_idx in row_idx_list:
        level = subtotal_levels.get(row_idx, -1)
        if level > prev_level and row_idx > start_row:
            merge_cells(worksheet, start_row, 1, row_idx - 1)
            start_row = row_idx + 1
        prev_level = level
    if start_row <= row_idx_list[-1]:
        merge_cells(worksheet, start_row, 1, row_idx_list[-1])


def merge_remaining_columns(worksheet, df: pd.DataFrame, subtotal_rows: Set[int],
                            subtotal_levels: Dict[int, int]) -> None:
    for col_idx, col in enumerate(df.columns[:-2], start=1):
        start_row = None
        prev_subtotal_row = None
        prev_subtotal_level = -1
        data_row_found = False

        for row_idx, value in enumerate(df[col], start=2):
            is_subtotal_row = row_idx in subtotal_rows
            current_subtotal_level = subtotal_levels.get(row_idx, -1)

            if is_subtotal_row:
                if start_row is not None and row_idx - 1 != start_row and current_subtotal_level > prev_subtotal_level:
                    merge_cells(worksheet, start_row, col_idx, row_idx - 1)
                elif prev_subtotal_row is not None and current_subtotal_level == prev_subtotal_level and not data_row_found:
                    merge_cells(worksheet, prev_subtotal_row + 1, col_idx, row_idx - 1)
                start_row = None
                prev_subtotal_row = row_idx
                prev_subtotal_level = current_subtotal_level
                data_row_found = False
            elif not pd.isna(value) and value != '':
                if start_row is not None:
                    merge_cells(worksheet, start_row, col_idx, row_idx - 1)
                start_row = None
                data_row_found = True
            elif start_row is None:
                start_row = row_idx

        if start_row is not None and start_row <= len(df):
            merge_cells(worksheet, start_row, col_idx, len(df) + 1)


def merge_cells(worksheet, start_row: int, col_idx: int, end_row: int) -> None:
    if end_row > start_row:
        worksheet.merge_cells(start_row=start_row, start_column=col_idx, end_row=end_row, end_column=col_idx)
