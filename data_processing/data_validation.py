import os


def validate_csv_path(csv_file_path: str) -> None:
    """Validate the CSV file path."""
    if not os.path.isfile(csv_file_path):
        raise FileNotFoundError(f"CSV file not found: {csv_file_path}")
