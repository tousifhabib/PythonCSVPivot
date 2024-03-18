import argparse
import json
import logging
import os
import datetime
from typing import Dict
from dataProcessing import load_data, preprocess_data, calculate_totals
from excelGeneration import save_excel
from pdfGeneration import dynamic_columns_for_pdf, save_pdf

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(module)s:%(message)s')
logger = logging.getLogger(__name__)

DEFAULT_CONFIG_FILE = 'config.json'


def read_config(config_file: str) -> Dict:
    """Read the configuration file."""
    with open(config_file, 'r') as file:
        return json.load(file)


def get_interactive_config() -> Dict:
    """Get data configuration from user input."""
    return {
        'csv_file_path': input("Enter the CSV file path: "),
        'filters': {'result': input("Enter the result filter (e.g., 'failed'): ")},
        'group_cols': input("Enter the group columns (comma-separated): ").split(','),
        'agg_func': input("Enter the aggregation function (e.g., 'count'): "),
        'agg_col': input("Enter the aggregation column name: "),
        'subtotal_col': input("Enter the subtotal columns (comma-separated): ").split(',')
    }


def validate_csv_path(csv_file_path: str) -> None:
    """Validate the CSV file path."""
    if not os.path.isfile(csv_file_path):
        raise FileNotFoundError(f"CSV file not found: {csv_file_path}")


def create_file_path(csv_file_path: str, file_extension: str) -> str:
    """Create the file path for PDF or Excel."""
    base_dir, timestamp = os.path.dirname(csv_file_path), datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(base_dir, f"{timestamp}_PivotTable.{file_extension}")


def process_data_and_generate_files(config: Dict, generate_files: bool = True) -> None:
    """Process data and generate PDF and Excel files."""
    try:
        csv_path = config["data"]["csv_file_path"]
        validate_csv_path(csv_path)
        df, file_data = load_data(csv_path), {"pdf": "pdf", "excel": "xlsx"}
        if df.empty:
            logger.warning("No data found in the CSV file.")
            return

        processed_data = preprocess_data(df, **{k: config["data"][k] for k in ["filters", "group_cols", "agg_func"]})
        final_data = calculate_totals(processed_data, config["data"]["subtotal_col"], config["data"]["agg_col"])
        dynamic_columns = dynamic_columns_for_pdf(config["data"]["group_cols"], config["data"]["agg_col"])

        for file_type, extension in file_data.items():
            file_path = create_file_path(csv_path, extension)
            if file_type == "pdf":
                save_pdf(final_data, file_path, dynamic_columns, config["styles"])
            else:
                save_excel(final_data, file_path, dynamic_columns, config["styles"])
            logger.info(f"{file_type.upper()} file built successfully.")
    except KeyError as e:
        logger.error(f"Error processing data: {str(e)}")


def main(args: argparse.Namespace) -> None:
    """Main function."""
    config_file = args.config_file or DEFAULT_CONFIG_FILE
    config = read_config(config_file) if not args.interactive else {'data': get_interactive_config(),
                                                                    'styles': read_config(config_file).get('styles',
                                                                                                           {})}
    process_data_and_generate_files(config)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Build PDF and Excel files.")
    parser.add_argument('--config-file', help="Path to the configuration file")
    parser.add_argument('--interactive', action='store_true', help="Run in interactive mode")
    main(parser.parse_args())