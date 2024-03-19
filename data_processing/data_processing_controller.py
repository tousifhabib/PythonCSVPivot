import json
import logging
from typing import Dict, List, Tuple, Optional, Union

import pandas as pd

from data_processing.dataProcessing import load_data, preprocess_data, calculate_totals
from data_processing.data_validation import validate_csv_path
from utils.file_utilities import create_file_path
from output_generation.excel.excelGeneration import save_excel
from output_generation.pdf.pdfGeneration import dynamic_columns_for_pdf, save_pdf

logger = logging.getLogger(__name__)

FileType = Union[str, Dict]


def process_data_and_generate_files(config: Dict) -> None:
    try:
        csv_path, processed_data = load_and_process_data(config["data"], source_type='csv')
        if processed_data is None:
            return

        final_data, dynamic_columns = prepare_data_for_output(processed_data, config["data"])
        save_output_files(csv_path, final_data, dynamic_columns, config["styles"], file_types=["pdf", "xlsx"])
    except Exception as e:
        logger.error(f"Failed to process and generate files: {e}", exc_info=True)


def load_and_process_data(data_config: Dict, source_type: str = 'csv') -> Tuple[str, Optional[pd.DataFrame]]:
    if source_type == 'csv':
        csv_path = data_config["csv_file_path"]
        validate_csv_path(csv_path)
        df = load_data(csv_path)
    else:  # Add logic for other types if necessary
        raise ValueError(f"Unsupported source type: {source_type}")

    if df.empty:
        logger.warning("No data found in the CSV file.")
        return csv_path, None

    return csv_path, preprocess_and_calculate(df, data_config)


def preprocess_and_calculate(df: pd.DataFrame, data_config: Dict) -> pd.DataFrame:
    processed_data = preprocess_data(
        df,
        filters=data_config["filters"],
        group_cols=data_config["group_cols"],
        agg_func=data_config["agg_func"]
    )
    return calculate_totals(processed_data, data_config["subtotal_col"], data_config["agg_col"])


def prepare_data_for_output(df: pd.DataFrame, data_config: Dict) -> Tuple[pd.DataFrame, List[str]]:
    dynamic_columns = dynamic_columns_for_pdf(data_config["group_cols"], data_config["agg_col"])
    return df, dynamic_columns


def save_output_files(base_path: str, final_data: pd.DataFrame, dynamic_columns: List[str], styles: Dict,
                      file_types: List[str]) -> Tuple[Optional[str], Optional[str]]:
    pdf_file = None
    excel_file = None

    for file_type in file_types:
        file_path = create_file_path(base_path, "pdf" if file_type == "pdf" else "xlsx")
        save_file = save_pdf if file_type == "pdf" else save_excel
        save_file(final_data, file_path, dynamic_columns, styles)
        logger.info(f"{file_type.upper()} file generated successfully at {file_path}.")

        if file_type == "pdf":
            pdf_file = file_path
        else:
            excel_file = file_path

    return pdf_file, excel_file


def process_json_data_and_generate_files(config: Dict, json_data: FileType) -> Tuple[Optional[str], Optional[str]]:
    try:
        df = extract_dataframe_from_json(json_data)
        if df.empty:
            logger.warning("No data found in the JSON.")
            return None, None

        processed_data = preprocess_and_calculate(df, config["data"])  # Directly use the DataFrame from JSON.
        if processed_data is None:
            return None, None

        final_data, dynamic_columns = prepare_data_for_output(processed_data, config["data"])
        pdf_file, excel_file = create_file_paths("data")
        save_output_files("data", final_data, dynamic_columns, config["styles"], file_types=["pdf", "xlsx"])

        return pdf_file, excel_file
    except Exception as e:
        logger.error(f"Failed to process JSON data: {e}", exc_info=True)
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
