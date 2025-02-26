import os
import re

def get_valid_excel_files(data_folder: str) -> list:
    """
    Checks if the given folder exists and retrieves a list of valid .xlsx files using regex.

    Args:
        data_folder (str): The folder path containing the Excel files.

    Returns:
        list: A list of valid .xlsx file paths.

    Raises:
        FileNotFoundError: If the data folder does not exist.
        ValueError: If no valid .xlsx files are found.
    """
    if not os.path.exists(data_folder):
        raise FileNotFoundError(f"Data folder not found: {data_folder}")

    # Regular expression to match only .xlsx files (not .xls or other formats)
    excel_pattern = re.compile(r".+\.xlsx$", re.IGNORECASE)

    valid_files = [
        os.path.join(data_folder, file)  # Include full path to the file
        # os.path.basename(file)     # Extract only the filename
        for file in os.listdir(data_folder)
            if excel_pattern.match(file)   # Match filenames using regex
    ]

    if not valid_files:
        raise ValueError(f"No valid .xlsx files found in {data_folder}")

    return valid_files

# Example usage
if __name__ == "__main__":
    folder_path = f"{os.getcwd()}/data"

    try:
        excel_files = get_valid_excel_files(folder_path)
        print("Valid Excel files:", excel_files)
    except (FileNotFoundError, ValueError) as e:
        print("Error:", e)
