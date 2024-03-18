import argparse
import json
import logging
import os
import datetime
from typing import Dict
from dataProcessing import load_data, preprocess_data, calculate_totals
from excelGeneration import save_excel
from pdfGeneration import dynamic_columns_for_pdf, save_pdf


logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(module)s:%(message)s')
logger = logging.getLogger(__name__)

DEFAULT_CONFIG_FILE = 'config.json'


def read_config(config_file: str) -> Dict:
    """Read the configuration file."""
    with open(config_file, 'r') as file:
        config = json.load(file)
    return config


def get_interactive_config() -> Dict:
    """Get the data configuration from user input."""
    data_config = {}
    data_config['csv_file_path'] = input("Enter the CSV file path: ")
    data_config['filters'] = {'result': input("Enter the result filter (e.g., 'failed'): ")}
    data_config['group_cols'] = input("Enter the group columns (comma-separated): ").split(',')
    data_config['agg_func'] = input("Enter the aggregation function (e.g., 'count'): ")
    data_config['agg_col'] = input("Enter the aggregation column name: ")
    data_config['subtotal_col'] = input("Enter the subtotal columns (comma-separated): ").split(',')
    return data_config


def validate_csv_path(csv_file_path: str) -> None:
    """Validate the CSV file path."""
    if not os.path.isfile(csv_file_path):
        raise FileNotFoundError(f"CSV file not found: {csv_file_path}")


def create_pdf_file_path(csv_file_path: str) -> str:
    """Create the PDF file path."""
    base_dir = os.path.dirname(csv_file_path)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(base_dir, f"{timestamp}_PivotTable.pdf")


def create_excel_file_path(csv_file_path: str) -> str:
    """Create the Excel file path."""
    base_dir = os.path.dirname(csv_file_path)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(base_dir, f"{timestamp}_PivotTable.xlsx")


def build_pdf(config: Dict) -> None:
    """Build the PDF file."""
    try:
        csv_file_path = config["data"]["csv_file_path"]
        filters = config["data"]["filters"]
        group_cols = config["data"]["group_cols"]
        agg_func = config["data"]["agg_func"]
        agg_col = config["data"]["agg_col"]
        subtotal_col = config["data"]["subtotal_col"]

        df = load_data(csv_file_path)
        if df.empty:
            logger.warning("No data found in the CSV file.")
            return

        processed_data = preprocess_data(df, filters, group_cols, agg_func)
        final_data = calculate_totals(processed_data, subtotal_col, agg_col)
        dynamic_columns = dynamic_columns_for_pdf(group_cols, agg_col)
        pdf_file_path = create_pdf_file_path(csv_file_path)
        save_pdf(final_data, pdf_file_path, dynamic_columns, config["styles"])
        logger.info("PDF built successfully.")
    except KeyError as e:
        logger.error(f"Error building PDF: {str(e)}")


def build_excel(config: Dict) -> None:
    """Build the Excel file."""
    try:
        csv_file_path = config["data"]["csv_file_path"]
        filters = config["data"]["filters"]
        group_cols = config["data"]["group_cols"]
        agg_func = config["data"]["agg_func"]
        agg_col = config["data"]["agg_col"]
        subtotal_col = config["data"]["subtotal_col"]

        df = load_data(csv_file_path)
        if df.empty:
            logger.warning("No data found in the CSV file.")
            return

        processed_data = preprocess_data(df, filters, group_cols, agg_func)
        final_data = calculate_totals(processed_data, subtotal_col, agg_col)
        dynamic_columns = dynamic_columns_for_pdf(group_cols, agg_col)
        excel_file_path = create_excel_file_path(csv_file_path)
        save_excel(final_data, excel_file_path, dynamic_columns, config["styles"])
        logger.info("Excel file built successfully.")
    except KeyError as e:
        logger.error(f"Error building Excel: {str(e)}")


def main(args: argparse.Namespace) -> None:
    """Main function."""
    if args.interactive:
        data_config = get_interactive_config()
        config = {'data': data_config}
        config_file = args.config_file or DEFAULT_CONFIG_FILE
        default_config = read_config(config_file)
        config['styles'] = default_config.get('styles', {})
    else:
        config_file = args.config_file or DEFAULT_CONFIG_FILE
        config = read_config(config_file)

    validate_csv_path(config['data']['csv_file_path'])

    build_pdf(config)
    build_excel(config)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Build PDF and Excel files.")
    parser.add_argument('--config-file', help="Path to the configuration file")
    parser.add_argument('--interactive', action='store_true', help="Run in interactive mode")
    args = parser.parse_args()
    main(args)