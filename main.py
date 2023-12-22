import os
import datetime
from reportlab.lib.pagesizes import letter
from dataProcessing import load_data, preprocess_data, calculate_totals
from pdfGeneration import save_pdf, dynamic_columns_for_pdf


def create_pdf_file_path(csv_file_path):
    return os.path.join(
        os.path.dirname(csv_file_path),
        datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + "_PivotTable.pdf"
    )


def main(csv_file_path, filters, group_cols, agg_func, agg_col, subtotal_col=None):
    df = load_data(csv_file_path)
    if df.empty:
        print("No data found in the CSV file.")
        return

    processed_data = preprocess_data(df, filters, group_cols, agg_func)
    final_data = calculate_totals(processed_data, subtotal_col, agg_col)
    dynamic_columns = dynamic_columns_for_pdf(group_cols, agg_col)

    pdf_file_path = create_pdf_file_path(csv_file_path)
    save_pdf(final_data, pdf_file_path, letter, dynamic_columns)


if __name__ == "__main__":
    csv_file_path = "/Users/tousifhabib/Code/pivotTable/Data/full_report_231220_1926 1.csv"
    filters = {"result": "failed"}
    group_cols = ["errorType", "comment", "feature"]
    agg_func = "count"
    agg_col = "Count"
    subtotal_col = ["errorType", "comment"]

    main(csv_file_path, filters, group_cols, agg_func, agg_col, subtotal_col)
