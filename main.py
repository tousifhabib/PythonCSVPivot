import json
import os
import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from dataProcessing import load_data, preprocess_data, calculate_totals
from pdfGeneration import save_pdf, dynamic_columns_for_pdf


def load_config():
    with open('config.json', 'r') as file:
        return json.load(file)


def create_pdf_file_path(csv_file_path):
    """Create a file path for the PDF based on the CSV file path and current timestamp."""
    base_dir = os.path.dirname(csv_file_path)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(base_dir, f"{timestamp}_PivotTable.pdf")


def main(csv_file_path, filters, group_cols, agg_func, agg_col, subtotal_col=None, page_size=letter):
    """Main function to process data and generate PDF."""
    config = load_config()
    try:
        df = load_data(csv_file_path)
        if df.empty:
            print("No data found in the CSV file.")
            return

        processed_data = preprocess_data(df, filters, group_cols, agg_func)
        final_data = calculate_totals(processed_data, subtotal_col, agg_col)
        dynamic_columns = dynamic_columns_for_pdf(group_cols, agg_col)

        pdf_file_path = create_pdf_file_path(csv_file_path)
        save_pdf(final_data, pdf_file_path, page_size, dynamic_columns, config)

    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    csv_file_path = "/Users/ts-tousif.habib/Downloads/TEST/full_report_231220_1926 1.csv"
    filters = {"result": "failed"}
    group_cols = ["errorType", "comment", "feature"]
    agg_func = "count"
    agg_col = "Count"
    subtotal_col = ["errorType", "comment"]

    custom_page_size = (10 * inch, 20 * inch)

    main(csv_file_path, filters, group_cols, agg_func, agg_col, subtotal_col, custom_page_size)
