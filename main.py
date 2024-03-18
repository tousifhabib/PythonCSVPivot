import json
import os
import datetime
from dataProcessing import load_data, preprocess_data, calculate_totals
from excelGeneration import save_excel
from pdfGeneration import save_pdf, dynamic_columns_for_pdf


def load_config():
    with open('config.json', 'r') as file:
        return json.load(file)


def create_pdf_file_path(csv_file_path):
    base_dir = os.path.dirname(csv_file_path)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(base_dir, f"{timestamp}_PivotTable.pdf")


def create_excel_file_path(csv_file_path):
    base_dir = os.path.dirname(csv_file_path)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(base_dir, f"{timestamp}_PivotTable.xlsx")


def main():
    config = load_config()
    try:
        csv_file_path = config["data"]["csv_file_path"]
        filters = config["data"]["filters"]
        group_cols = config["data"]["group_cols"]
        agg_func = config["data"]["agg_func"]
        agg_col = config["data"]["agg_col"]
        subtotal_col = config["data"]["subtotal_col"]

        df = load_data(csv_file_path)
        if df.empty:
            print("No data found in the CSV file.")
            return

        processed_data = preprocess_data(df, filters, group_cols, agg_func)
        final_data = calculate_totals(processed_data, subtotal_col, agg_col)
        dynamic_columns = dynamic_columns_for_pdf(group_cols, agg_col)

        pdf_file_path = create_pdf_file_path(csv_file_path)
        save_pdf(final_data, pdf_file_path, dynamic_columns, config["styles"])

        excel_file_path = create_excel_file_path(csv_file_path)
        save_excel(final_data, excel_file_path, dynamic_columns, config["styles"])

    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    main()
