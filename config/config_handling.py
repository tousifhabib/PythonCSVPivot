import json
from pathlib import Path
from typing import Dict


def read_config(config_file: Path) -> Dict:
    with open(config_file, 'r') as file:
        return json.load(file)


def get_interactive_config() -> Dict:
    return {
        'csv_file_path': input("Enter the CSV file path: "),
        'filters': {'result': input("Enter the result filter (e.g., 'failed'): ")},
        'group_cols': input("Enter the group columns (comma-separated): ").split(','),
        'agg_func': input("Enter the aggregation function (e.g., 'count'): "),
        'agg_col': input("Enter the aggregation column name: "),
        'subtotal_col': input("Enter the subtotal columns (comma-separated): ").split(',')
    }
