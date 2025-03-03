import pandas as pd
import os
import openpyxl

from src.file_reader import read_excel_file


def separate_courses(final_result):
    """
    Separates courses into two categories: General Education (GE) and Major & Elective,
    ensuring 'ENGL 101 (ENR)-Enrichment' stays in GE + Other but is excluded from Major & Elective.

    Args:
        final_result (pd.DataFrame): The processed DataFrame containing all courses.

    Returns:
        tuple: Two DataFrames:
            - df_ge: General Education courses
            - df_major_elective: Major & Elective courses
    """
    
    # Select rows for General Education (GE) courses
    df_ge = final_result[final_result["Courses"].str.contains(r"-GE", na=False)]

    # The remaining courses go under Major & Elective, excluding 'ENGL 101 (ENR)-Enrichment'
    df_major_elective = final_result[
        ~final_result["Courses"].str.contains(r"-GE", na=False) & 
        ~final_result["Courses"].str.contains(r"ENGL 101 \(ENR\)-Enrichment", na=False)
    ]

    return df_ge, df_major_elective



def sum_only_numbers(column):
    """
    Sums only numeric values (int or float) in a given column.
    
    Args:
        column (pd.Series): The column of data to sum.

    Returns:
        float: The sum of numeric values in the column.
    """
    return sum(value for value in column if isinstance(value, (int, float)))




def process_excel_files(excel_files):
    """
    Processes multiple Excel files to generate a combined summary of courses.

    This function reads each provided Excel file, removes unnecessary columns, 
    handles missing values, processes numerical data, and merges results into 
    a final summarized DataFrame.

    Parameters:
    -----------
    excel_files : list of str
        A list of file paths to the Excel files to be processed.

    Returns:
    --------
    pd.DataFrame
        A DataFrame where:
        - Rows represent courses (with " (Required)" removed from names).
        - Columns represent the sum of values for each course from each file.
        - A "Total" column aggregates all course values across files.

    Processing Steps:
    -----------------
    1. Read each Excel file into a DataFrame.
    2. Drop the columns "Srl No" and "Name" if present.
    3. Fill missing values with 1.
    4. Sum numeric values for each course.
    5. Merge results from all files based on course names.
    6. Remove "(Required)" from course names and aggregate values.
    7. Compute a "Total" column summing all values across files.
    8. Replace zero values with empty strings for better Excel readability.

    """

    result_dfs = []

    for file_name in excel_files:
        
        # Read each Excel file into a DataFrame
        df = read_excel_file(file_name)

        if df is not None:

            # Remove unnecessary columns if present
            columns_to_remove = ["Srl No", "Name"]
            df = df.drop(columns=[col for col in df.columns if col in columns_to_remove], errors='ignore')

            # Fill missing values with 1
            df.fillna(1, inplace=True) 
            
            # Process columns (sum, etc.) based on your need
            course_sums = df.apply(sum_only_numbers)  # Implement the sum_only_numbers function
            course_sums_df = course_sums.reset_index()

            # Use the base file name as a column header
            file_base_name = os.path.splitext(os.path.basename(file_name))[0]
            course_sums_df.columns = ["Courses", file_base_name]

            result_dfs.append(course_sums_df)

    # Merge all results into one DataFrame
    final_result = result_dfs[0]
    for df in result_dfs[1:]:
        final_result = pd.merge(final_result, df, on="Courses", how="outer")

    # Clean up course names by removing "(Required)"
    final_result['Courses'] = final_result['Courses'].str.replace(r" \(Required\)", "", regex=True)

    # Sum values for each course across all files
    final_result = final_result.groupby("Courses", as_index=False).sum()

    # Compute "Total" column
    final_result["Total"] = final_result.iloc[:, 1:].sum(axis=1)

    # Replace zero values with empty strings for Excel readability
    final_result.replace({0: ""}, inplace=True)     
    
    return final_result


def save_to_excel(final_result, selected_path=None):
    """
    Saves the processed data to an Excel file with three sheets:
    - "GE + Other": Combined General Education + Major & Elective courses
    - "General Education": GE courses only
    - "Major & Elective": Major and elective courses only

    Parameters:
    -----------
    final_result : pd.DataFrame
        The processed DataFrame containing course data to be saved.
    selected_path : str, optional
        The desired output file path for the Excel file. If not provided, defaults to "output/result.xlsx".

    Process:
    --------
    1. Separates the courses into General Education (GE) and Major & Elective categories.
    2. Saves the full dataset and separated datasets into respective sheets.
    3. Adjusts column widths for better readability.
    4. Handles exceptions and prints error messages if any issues arise during the saving process.
    """

    try:

        # Default output directory if no path is provided
        if selected_path is None:
            output_dir = "output"
            output_file = os.path.join(output_dir, "result.xlsx")
        elif selected_path.endswith(".xlsx"):
            output_file = selected_path
            output_dir = os.path.dirname(output_file)
        else:
            output_dir = selected_path
            output_file = os.path.join(output_dir, "result.xlsx")

        # Ensure the directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        # Separate GE and Major/Elective courses
        df_ge, df_major_elective = separate_courses(final_result)

        # Create the Excel writer object and write data to sheets
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

            # Save all courses together
            final_result.to_excel(writer, index=False, sheet_name='GE + Other')

            # Save General Education courses
            df_ge.to_excel(writer, index=False, sheet_name='General Education')

            # Save Major & Elective courses
            df_major_elective.to_excel(writer, index=False, sheet_name='Major & Elective')

            workbook = writer.book
        
            # Adjust column width for better readability
            for sheet_name in ['GE + Other', 'General Education', 'Major & Elective']:
                worksheet = workbook[sheet_name]
                for col in worksheet.columns:
                    max_length = 0
                    column = [cell for cell in col]
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 4)
                    worksheet.column_dimensions[openpyxl.utils.get_column_letter(column[0].column)].width = adjusted_width

        print(f"Data saved to {output_file} with expanded column widths.")

    except Exception as e:
        print(f"Error saving file: {e}")
