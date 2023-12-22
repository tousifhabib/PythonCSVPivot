from reportlab.platypus import SimpleDocTemplate, Table
import logging
from tableStyling import create_table_data, apply_table_styles


def save_pdf(data, file_path, page_size, dynamic_cols):
    if data.empty:
        print("No data to display in the PDF.")
        return

    doc = SimpleDocTemplate(file_path, pagesize=page_size)
    table_data = create_table_data(data, dynamic_cols)
    table = Table(table_data)

    apply_table_styles(table, table_data, dynamic_cols)

    try:
        doc.build([table])
    except Exception as e:
        logging.error(f"Error building PDF: {e}")
    logging.info("PDF generation completed.")


def dynamic_columns_for_pdf(group_cols, agg_col):
    return group_cols + [agg_col]
