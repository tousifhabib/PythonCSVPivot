import logging
from reportlab.platypus import SimpleDocTemplate, Table
from output_generation.excel.excelStyles import calculate_dynamic_page_size
from output_generation.pdf.tableStyling import create_table_data, apply_table_styles

logging.basicConfig(level=logging.DEBUG)


def save_pdf(data, file_path, dynamic_cols, config):
    if data.empty:
        logging.warning("No data to display in the PDF.")
        return

    paginate = config.get("paginate", True)

    try:
        table_data = create_table_data(data, dynamic_cols)
        num_rows = len(data)
        table = Table(table_data)
        table_width, table_height = table.wrap(0, 0)

        dynamic_page_size = calculate_dynamic_page_size(table_width, table_height, config, paginate, num_rows)

        apply_table_styles(table, table_data, dynamic_cols, config)
        build_pdf(file_path, table, dynamic_page_size)
        logging.info("PDF generation completed successfully.")
    except Exception as e:
        logging.error(f"Error building PDF: {e}")


def build_pdf(file_path, table, page_size):
    doc = SimpleDocTemplate(file_path, pagesize=page_size)
    doc.build([table])


def dynamic_columns_for_pdf(group_cols, agg_col):
    return group_cols + [agg_col]
