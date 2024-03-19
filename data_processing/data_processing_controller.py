import json
from typing import Dict, Tuple, Union

import pandas as pd

from data_processing.dataProcessing import load_data, preprocess_data, calculate_totals
from data_processing.data_validation import validate_csv_path
from utils.file_utilities import create_file_path
from output_generation.excel.excelGeneration import save_excel
from output_generation.pdf.pdfGeneration import dynamic_columns_for_pdf, save_pdf
import logging

logger = logging.getLogger(__name__)


def process_data_and_generate_files(config: Dict) -> None:
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


def process_json_data_and_generate_files(config: Dict, json_data: Dict) -> Union[tuple[None, None], tuple[str, str]]:
    try:
        if 'data' not in json_data:
            logger.warning("No 'data' key found in the JSON.")
            return None, None

        df = pd.DataFrame(json_data['data'])

        print(df)

        if df.empty:
            logger.warning("No data found in the JSON.")
            return None, None

        processed_data = preprocess_data(df, **{k: config["data"][k] for k in ["filters", "group_cols", "agg_func"]})
        final_data = calculate_totals(processed_data, config["data"]["subtotal_col"], config["data"]["agg_col"])
        dynamic_columns = dynamic_columns_for_pdf(config["data"]["group_cols"], config["data"]["agg_col"])

        pdf_file = create_file_path("data", "pdf")
        excel_file = create_file_path("data", "xlsx")

        save_pdf(final_data, pdf_file, dynamic_columns, config["styles"])
        save_excel(final_data, excel_file, dynamic_columns, config["styles"])

        logger.info("PDF and Excel files built successfully.")
        return pdf_file, excel_file

    except KeyError as e:
        logger.error(f"Error processing JSON data: {str(e)}")
        return None, None