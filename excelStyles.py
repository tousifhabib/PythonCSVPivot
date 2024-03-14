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


def merge_empty_cells(worksheet: Worksheet, df: pd.DataFrame) -> None:
    print("DATAFRAME", df)

    subtotal_rows = {}
    for col_idx, col in enumerate(df.columns, start=1):
        for row_idx, value in enumerate(df[col], start=2):
            if "Grand Total" in str(value):
                subtotal_rows[row_idx] = col_idx

    for col_idx, col in enumerate(df.columns, start=1):
        start_row = None
        for row_idx, value in enumerate(df[col], start=2):
            if row_idx in subtotal_rows and subtotal_rows[row_idx] >= col_idx:
                if start_row is not None:
                    merge_cells(worksheet, start_row, col_idx, row_idx - 1)
                start_row = None
                continue

            if pd.isna(value) or value == '':
                start_row = start_row or row_idx
            else:
                if start_row is not None and row_idx - start_row > 1:
                    merge_cells(worksheet, start_row, col_idx, row_idx - 1)
                start_row = None

        if start_row is not None and row_idx - start_row > 1:
            merge_cells(worksheet, start_row, col_idx, row_idx)


def merge_cells(worksheet: Worksheet, start_row: int, col_idx: int, end_row: int) -> None:
    worksheet.merge_cells(start_row=start_row, start_column=col_idx, end_row=end_row, end_column=col_idx)
    print(f"Merged from {start_row} to {end_row} in column: {col_idx}")
