import os
import pandas as pd
from src.file_checker import get_valid_excel_files 


def read_excel_file(file_path):
    """
    Read and process the Excel file.
    """
    print(f"Reading file: {file_path}")
    try:
        df = pd.read_excel(file_path, header=1)  # Adjust as per your specific format
        return df
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return None

# Example usage
if __name__ == "__main__":
    folder_path = "data"
    try:
        excel_files = get_valid_excel_files(folder_path)  # Get valid Excel files
        for file in excel_files:
            file_path = os.path.join(folder_path, file)  # Full file path
            read_excel_file(file_path)  # Read the file
    except (FileNotFoundError, ValueError) as e:
        print("Error:", e)
