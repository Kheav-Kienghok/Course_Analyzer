import pandas as pd
import os
import openpyxl
from src.file_reader import read_excel_file



def separate_courses(final_result):
    """
    Separates courses into GE (General Education) and Major & Elective,
    ensuring that 'ENGL 101 (ENR)-Enrichment' remains in GE + Other 
    but is excluded from Major & Elective.

    Args:
        final_result (pd.DataFrame): The processed DataFrame containing all courses.

    Returns:
        tuple: DataFrames for GE courses and Major & Elective courses.
    """
    # Identify GE courses
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
    fills missing values, processes numerical data, and merges results into 
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
        
        # Reading each Excel file
        df = read_excel_file(file_name)

        if df is not None:

            # Assuming 'df' contains your data and needs processing
            columns_to_remove = ["Srl No", "Name"]
            df = df.drop(columns=[col for col in df.columns if col in columns_to_remove], errors='ignore')
            df.fillna(1, inplace=True)  # Handle missing data
            
            # Process columns (sum, etc.) based on your need
            course_sums = df.apply(sum_only_numbers)  # Implement sum_only_numbers logic
            course_sums_df = course_sums.reset_index()
            file_base_name = os.path.splitext(os.path.basename(file_name))[0]
            course_sums_df.columns = ["Courses", file_base_name]
            result_dfs.append(course_sums_df)

    # Merge all results into one
    final_result = result_dfs[0]
    for df in result_dfs[1:]:
        final_result = pd.merge(final_result, df, on="Courses", how="outer")

    final_result['Courses'] = final_result['Courses'].str.replace(r" \(Required\)", "", regex=True)
    final_result = final_result.groupby("Courses", as_index=False).sum()

    # Compute total column
    final_result["Total"] = final_result.iloc[:, 1:].sum(axis=1)
    final_result.replace({0: ""}, inplace=True)  # Replace 0 with empty string for Excel
    
    return final_result


def save_to_excel(final_result, output_file="output/result.xlsx"):
    """
    Saves the processed data to an Excel file with three sheets for better organization.

    This function writes the final processed DataFrame into an Excel file, separating 
    the data into three sheets for clarity:
    - "GE + Other" (General Education + Major & Elective courses combined)
    - "General Education" (GE courses only)
    - "Major & Elective" (Major, Core/Major and elective courses only)

    Parameters:
    -----------
    final_result : pd.DataFrame
        The processed DataFrame containing course data to be saved.
    output_file : str, optional
        The file path for the output Excel file (default is "output/result.xlsx").

    Processing Steps:
    -----------------
    1. Separate General Education (GE) and Major & Elective courses using `separate_courses`.
    2. Write the full dataset and separated datasets into three Excel sheets.
    3. Automatically adjust column widths for better readability.
    4. Handle any exceptions and print errors if the saving process fails.

    Exceptions:
    -----------
    - Catches general exceptions and prints an error message if any issues occur.
    """
    try:
        
        output_dir = os.path.dirname(output_file)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Separate GE and Major/Elective courses
        df_ge, df_major_elective = separate_courses(final_result)

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
