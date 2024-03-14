from typing import Dict, Any, List, Tuple

import openpyxl
import pandas as pd
from openpyxl.styles import PatternFill, Font
from openpyxl.cell import Cell
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.worksheet import Worksheet
from pandas import DataFrame


def create_fill(color: str) -> PatternFill:
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def create_font(color: str, bold: bool = False) -> Font:
    return Font(color=color, bold=bold)


def style_cell(cell: Cell, fill: PatternFill, font: Font) -> None:
    cell.fill = fill
    cell.font = font


def apply_row_styles(worksheet: Worksheet, config: Dict[str, Any], dynamic_cols: List[str]) -> None:
    style_mappings = {
        'header': (
            create_fill(config['colors']['header']['background']),
            create_font(config['colors']['header']['text'], True)),
        'default': (
            create_fill(config['colors']['default']['background']), create_font(config['colors']['default']['text'])),
        'subtotal': (
            create_fill(config['colors']['subtotal']['background']), create_font(config['colors']['subtotal']['text'])),
        'grand_total': (create_fill(config['colors']['grand_total']['background']),
                        create_font(config['colors']['grand_total']['text']))
    }

    def determine_style(row_values: List[Any], idx: int) -> Tuple[PatternFill, Font]:
        if "Grand Total" in row_values:
            return style_mappings['grand_total']
        elif is_special_row_excel(row_values, dynamic_cols) and idx != 1:
            return style_mappings['subtotal']
        return style_mappings['default']

    for idx, row in enumerate(worksheet.iter_rows(), start=1):
        row_values = [cell.value for cell in row]
        fill, font = determine_style(row_values, idx) if idx > 1 else style_mappings['header']
        for cell in row:
            style_cell(cell, fill, font)


def is_special_row_excel(row_values: List[Any], dynamic_cols: List[str]) -> bool:
    return any([
        "Subtotal" in row_values,
        "Grand Total" in row_values,
        any(row_values[i] and not row_values[i + 1] for i, _ in enumerate(dynamic_cols) if i < len(row_values) - 1)
    ])


def apply_excel_colors(worksheet: Worksheet, config: Dict[str, Any], df: DataFrame, dynamic_cols: List[str]) -> None:
    apply_row_styles(worksheet, config, dynamic_cols)


def merge_empty_cells(worksheet, df):
    for col_idx, col in enumerate(df.columns, start=1):
        start_row = None
        for row_idx, value in enumerate(df[col], start=2):
            current_row = df.iloc[row_idx - 2]
            previous_row = df.iloc[row_idx - 3] if row_idx > 2 else None
            next_row = df.iloc[row_idx - 1] if row_idx < len(df) else None

            if is_special_row(current_row) or is_special_row(next_row):
                if start_row is not None and not is_special_row(previous_row):
                    merge_cells(worksheet, start_row, col_idx, row_idx - 1)
                    start_row = None
                continue

            if should_span(value, current_row, previous_row, next_row):
                start_row = start_row or row_idx
            else:
                if start_row is not None and row_idx - start_row > 1 and not is_special_row(previous_row):
                    merge_cells(worksheet, start_row, col_idx, row_idx - 1)
                    start_row = None

        if start_row is not None and row_idx - start_row > 1 and not is_special_row(df.iloc[row_idx - 1]):
            merge_cells(worksheet, start_row, col_idx, row_idx)


def should_span(cell_value, current_row, previous_row, next_row):
    return (
            pd.isna(cell_value) or cell_value == ''
    ) and not any(is_special_row(row) for row in (current_row, previous_row, next_row) if row is not None)


def is_special_row(row):
    return is_subtotal_row(row) or is_grand_total_row(row)


def is_subtotal_row(row):
    return any("Subtotal" in str(cell) for cell in row)


def is_grand_total_row(row):
    return any("Grand Total" in str(cell) for cell in row)


def merge_cells(worksheet, start_row, col_idx, end_row):
    worksheet.merge_cells(start_row=start_row, start_column=col_idx, end_row=end_row, end_column=col_idx)
    print(f"Merged from {start_row} to {end_row} in column: {col_idx}")
