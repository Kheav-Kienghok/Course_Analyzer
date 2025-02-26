from src.file_checker import get_valid_excel_files
from src.data_processing import process_excel_files, save_to_excel

import os

def main():

    # Get the folder path where the files are located
    folder_path = input("Enter the file directory: ")

    # Ensure the folder path is correct and exists
    if not os.path.exists(folder_path):
        print(f"Error: The folder '{folder_path}' does not exist.")
        return

    try:
        # Get the list of valid Excel files from the folder
        excel_files = get_valid_excel_files(folder_path)

        if not excel_files:
            print("No valid Excel files found in the specified directory.")
            return
        
        # Process the Excel files
        final_result = process_excel_files(excel_files)
        
        # Do something with final_result, like saving to a new file
        print("Processing completed.")

        # Save the final result to an Excel file
        save_to_excel(final_result)

    
    except (FileNotFoundError, ValueError) as e:
        print("Error:", e)

if __name__ == "__main__":
    main()
