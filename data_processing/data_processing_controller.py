import json
import logging
from typing import Dict, Union, Tuple, Optional, List

import pandas as pd

from data_processing.dataProcessing import load_data, preprocess_data, calculate_totals
from data_processing.data_validation import validate_csv_path
from utils.file_utilities import create_file_path
from output_generation.excel.excelGeneration import save_excel
from output_generation.pdf.pdfGeneration import dynamic_columns_for_pdf, save_pdf

logger = logging.getLogger(__name__)


def process_data_and_generate_files(config: Dict) -> None:
    try:
        csv_path = config["data"]["csv_file_path"]
        validate_csv_path(csv_path)
        df, file_data = load_data(csv_path), {"pdf": "pdf", "excel": "xlsx"}

        if df.empty:
            logger.warning("No data found in the CSV file.")
            return

        processed_data = preprocess_data(
            df,
            filters=config["data"]["filters"],
            group_cols=config["data"]["group_cols"],
            agg_func=config["data"]["agg_func"]
        )
        final_data = calculate_totals(processed_data, config["data"]["subtotal_col"], config["data"]["agg_col"])
        dynamic_columns = dynamic_columns_for_pdf(config["data"]["group_cols"], config["data"]["agg_col"])

        save_files(csv_path, file_data, final_data, dynamic_columns, config["styles"])

    except KeyError as e:
        logger.error(f"Error processing data: {str(e)}")


def save_files(csv_path: str, file_data: Dict[str, str], final_data: pd.DataFrame,
               dynamic_columns: List[str], styles: Dict) -> None:
    for file_type, extension in file_data.items():
        file_path = create_file_path(csv_path, extension)
        save_file = save_pdf if file_type == "pdf" else save_excel
        save_file(final_data, file_path, dynamic_columns, styles)
        logger.info(f"{file_type.upper()} file built successfully.")


def process_json_data_and_generate_files(config: Dict, json_data: Union[str, Dict]) -> Tuple[
    Optional[str], Optional[str]]:
    try:
        df = extract_dataframe_from_json(json_data)

        if df.empty:
            logger.warning("No data found in the JSON.")
            return None, None

        processed_data = preprocess_data(
            df,
            filters=config["data"]["filters"],
            group_cols=config["data"]["group_cols"],
            agg_func=config["data"]["agg_func"]
        )
        final_data = calculate_totals(processed_data, config["data"]["subtotal_col"], config["data"]["agg_col"])
        dynamic_columns = dynamic_columns_for_pdf(config["data"]["group_cols"], config["data"]["agg_col"])

        pdf_file, excel_file = create_file_paths("data")
        save_pdf(final_data, pdf_file, dynamic_columns, config["styles"])
        save_excel(final_data, excel_file, dynamic_columns, config["styles"])

        logger.info("PDF and Excel files built successfully.")
        return pdf_file, excel_file

    except (KeyError, json.JSONDecodeError) as e:
        logger.error(f"Error processing JSON data: {str(e)}")
        return None, None


def extract_dataframe_from_json(json_data: Union[str, Dict]) -> pd.DataFrame:
    if isinstance(json_data, str):
        json_data = json.loads(json_data)

    if 'data' not in json_data:
        logger.warning("No 'data' key found in the JSON.")
        return pd.DataFrame()

    return pd.DataFrame(json_data['data'])


def create_file_paths(base_name: str) -> Tuple[str, str]:
    return create_file_path(base_name, "pdf"), create_file_path(base_name, "xlsx")
