from data_processing.tableDataProcessing import preprocess_dataframe, assign_subtotal_levels, clean_up_columns, format_table_data
from output_generation.pdf.tableStyleEnhancements import apply_span_styles, apply_special_row_styles, apply_base_styles


def create_table_data(df, dynamic_cols):
    df = preprocess_dataframe(df)
    assign_subtotal_levels(df, dynamic_cols)
    clean_up_columns(df, dynamic_cols)
    return format_table_data(df)


def apply_table_styles(table, table_data, dynamic_cols, config):
    """Apply styles to the table including span styles and special row styles."""
    style = apply_base_styles(table, config)
    apply_span_styles(table, table_data, dynamic_cols)
    apply_special_row_styles(table, table_data, style, dynamic_cols, config)
    table.setStyle(style)
