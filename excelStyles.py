from typing import Dict, Any, List, Tuple
from openpyxl.styles import PatternFill, Font
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet
from pandas import DataFrame


def create_fill(color: str) -> PatternFill:
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def create_font(color: str, bold: bool = False) -> Font:
    return Font(color=color, bold=bold)


def style_cell(cell: Cell, fill: PatternFill, font: Font) -> None:
    cell.fill = fill
    cell.font = font


def is_special_row_excel(row_values: List[Any], dynamic_cols: List[str]) -> bool:
    return any([
        "Subtotal" in row_values,
        "Grand Total" in row_values,
        any(row_values[i] and not row_values[i + 1] for i, _ in enumerate(dynamic_cols) if i < len(row_values) - 1)
    ])


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


def apply_excel_colors(worksheet: Worksheet, config: Dict[str, Any], df: DataFrame, dynamic_cols: List[str]) -> None:
    apply_row_styles(worksheet, config, dynamic_cols)
