import os
from datetime import datetime


def ensure_output_folder(folder_path):
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def timestamp_string():
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


def build_log_path(output_folder):
    return os.path.join(output_folder, f"email_log_{timestamp_string()}.csv")


def safe_get(value):
    if value is None:
        return ""
    return str(value).strip()