from reportlab.lib import colors
from reportlab.platypus import TableStyle


def create_table_data(df, dynamic_cols):
    df = df.copy()
    df.fillna("", inplace=True)

    subtotal_and_total_rows = df[dynamic_cols[-1]].isin(["", "Grand Total"])
    df.loc[subtotal_and_total_rows, dynamic_cols[:-1]] = ""

    for idx, col in enumerate(dynamic_cols[:-1]):
        df[col] = df[col].mask(
            df[col].eq(df[col].shift()) &
            df[dynamic_cols[:idx]]
            .eq(df[dynamic_cols[:idx]].shift(1))
            .all(axis=1), "")

    table_data = [df.columns.tolist()] + df.values.tolist()

    return table_data


def apply_base_styles(table):
    """Apply base styles to the table."""
    style = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ]
    )
    table.setStyle(style)
    return style


def apply_table_styles(table, table_data, dynamic_cols):
    """Apply styles to the table including span styles and special row styles."""
    style = apply_base_styles(table)

    # Call apply_span_styles with the correct number of arguments
    apply_span_styles(table, table_data, dynamic_cols)

    apply_special_row_styles(table_data, style, dynamic_cols)

    table.setStyle(style)


def apply_special_row_styles(table_data, style, dynamic_cols):
    """Apply styles for subtotal, grand total, and other special rows."""
    for row_num, row_data in enumerate(table_data):
        if row_num == 0:
            continue

        if "Grand Total" in row_data:
            style.add("BACKGROUND", (0, row_num), (-1, row_num), colors.yellow)
        elif row_data[dynamic_cols.index(dynamic_cols[-1])] == "":
            style.add("BACKGROUND", (0, row_num), (-1, row_num), colors.lightgrey)


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
    is_subtotal_or_total = "Subtotal" in row_data or "Grand Total" in row_data
    if row_num > 0 and not is_subtotal_or_total:
        if row_data[0] != "" and all(row_data[i] == "" for i in range(1, len(dynamic_cols) - 1)):
            prev_row_data = table_data[row_num - 1]
            if prev_row_data[0] != row_data[0] and not all(
                    prev_row_data[i] == "" for i in range(1, len(dynamic_cols) - 1)):
                return True
    return is_subtotal_or_total


def should_span(cell_value, row_data):
    return cell_value == "" and not ("Subtotal" in row_data or "Grand Total" in row_data)


def span_cells(table, col_index, start_row, end_row):
    if start_row is not None and start_row < end_row - 1:
        table.setStyle(TableStyle([("SPAN", (col_index, start_row), (col_index, end_row - 1))]))
