import os
import datetime


def create_file_path(csv_file_path: str, file_extension: str) -> str:
    base_dir, timestamp = os.path.dirname(csv_file_path), datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(base_dir, f"{timestamp}_PivotTable.{file_extension}")
