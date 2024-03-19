from reportlab.lib import colors
from reportlab.platypus import TableStyle


def apply_base_styles(table, config):
    """Apply base styles to the table."""
    global_align = config['alignment']['global']
    header_align = config['alignment']['header']

    style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, 0), header_align),
        ("ALIGN", (0, 1), (-1, -1), global_align),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
    ])

    table.setStyle(style)
    return style


def apply_special_row_styles(table, table_data, style, dynamic_cols, config):
    """Apply styles for subtotal, grand total, and other special rows."""
    colors_list = [colors.lightgrey, colors.lightblue]
    alignment = extract_alignment_settings(config)

    for row_num, row_data in enumerate(table_data[1:], 1):
        if "Grand Total" in row_data:
            apply_grand_total_style(style, row_num, alignment['numbers'])
        else:
            apply_row_style(style, row_data, row_num, dynamic_cols, colors_list, alignment)

    table.setStyle(style)


def extract_alignment_settings(config):
    """Extract and return alignment settings from the config."""
    return {
        'subtotal': config['alignment']['content']['subtotal'],
        'numbers': config['alignment']['content']['numbers'],
        'text': config['alignment']['content']['text']
    }


def apply_grand_total_style(style, row_num, number_align):
    """Apply style specific to grand total rows."""
    style.add("BACKGROUND", (0, row_num), (-1, row_num), colors.yellow)
    style.add("ALIGN", (-1, row_num), (-1, row_num), number_align)


def apply_row_style(style, row_data, row_num, dynamic_cols, colors_list, alignment):
    """Apply styles for a standard row."""
    apply_background_color(style, row_data, row_num, dynamic_cols, colors_list)
    apply_text_alignment(style, row_data, row_num, alignment)


def apply_background_color(style, row_data, row_num, dynamic_cols, colors_list):
    """Apply background color based on the level of the row."""
    level = determine_level(row_data, dynamic_cols)
    if level is not None and level < len(colors_list):
        start_col = next((i for i, x in enumerate(row_data) if x), len(row_data))
        style.add("BACKGROUND", (start_col, row_num), (-1, row_num), colors_list[level])


def determine_level(row_data, dynamic_cols):
    """Determine the level of the row for styling."""
    for i, col in enumerate(dynamic_cols):
        if row_data[i] != "" and (i == len(dynamic_cols) - 1 or row_data[i + 1] == ""):
            return i
    return None


def apply_text_alignment(style, row_data, row_num, alignment):
    """Apply text alignment based on the content of the cells."""
    for col_num, cell in enumerate(row_data):
        if isinstance(cell, (int, float)):
            style.add("ALIGN", (col_num, row_num), (col_num, row_num), alignment['numbers'])
        elif "Subtotal" in cell or "Grand Total" in cell:
            style.add("ALIGN", (col_num, row_num), (col_num, row_num), alignment['subtotal'])
        else:
            style.add("ALIGN", (col_num, row_num), (col_num, row_num), alignment['text'])


def apply_span_styles(table, table_data, dynamic_cols):
    num_columns = len(table_data[0])
    for col_index in range(num_columns):
        apply_column_span(table, table_data, col_index, dynamic_cols)


def apply_column_span(table, table_data, col_index, dynamic_cols):
    start_row = None
    for row_num, row_data in enumerate(table_data):
        if row_num == 0 or is_special_row(row_data, dynamic_cols, row_num, table_data):
            span_cells(table, col_index, start_row, row_num)
            start_row = None
        elif should_span(row_data[col_index], row_data):
            start_row = start_row if start_row is not None else row_num
        else:
            span_cells(table, col_index, start_row, row_num)
            start_row = None
    span_cells(table, col_index, start_row, len(table_data))


def is_special_row(row_data, dynamic_cols, row_num, table_data):
    is_total = "Grand Total" in row_data
    if row_num > 0 and not is_total:
        if row_data[0] != "" and all(row_data[i] == "" for i in range(1, len(dynamic_cols) - 1)):
            prev_row_data = table_data[row_num - 1]
            if prev_row_data[0] != row_data[0] and not all(
                    prev_row_data[i] == "" for i in range(1, len(dynamic_cols) - 1)):
                return True
    return is_total


def should_span(cell_value, row_data):
    return cell_value == "" and not ("Grand Total" in row_data)


def span_cells(table, col_index, start_row, end_row):
    if start_row is not None and start_row < end_row - 1:
        table.setStyle(TableStyle([("SPAN", (col_index, start_row), (col_index, end_row - 1))]))
