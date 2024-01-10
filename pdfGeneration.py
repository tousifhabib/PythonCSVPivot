import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table
import logging
from tableStyling import create_table_data, apply_table_styles

logging.basicConfig(level=logging.INFO)


def save_pdf(data, file_path, page_size, dynamic_cols, config):
    """Generate a PDF with the given data and specifications."""
    if data.empty:
        logging.warning("No data to display in the PDF.")
        return

    try:
        doc = SimpleDocTemplate(file_path, pagesize=page_size)
        table_data = create_table_data(data, dynamic_cols)
        table = Table(table_data)

        print(table_data)

        apply_table_styles(table, table_data, dynamic_cols, config)
        doc.build([table])
        logging.info("PDF generation completed successfully.")

    except Exception as e:
        logging.error(f"Error building PDF: {e}")


def save_excel(data, file_path, dynamic_cols, config):
    """Create the Excel report from the given DataFrame."""
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        data.to_excel(writer, sheet_name='Report', index=False)

    print(f"Saved Excel report to {file_path}")

def dynamic_columns_for_pdf(group_cols, agg_col):
    """Determine dynamic columns for the PDF based on grouping and aggregate columns."""
    return group_cols + [agg_col]
