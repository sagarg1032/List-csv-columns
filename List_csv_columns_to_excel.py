import os
import pandas as pd

# Initialize Function
def list_csv_columns_to_excel(directory, output_file):
    """
    Reads all CSV files in a specified directory, extracts their column names,
    and saves the results to an Excel file.

    Parameters:
    - directory: str, path to the folder containing CSV files.
    - output_file: str, path to save the output Excel file.
    """
    
    # Check if directory exists
    if not os.path.isdir(directory):
        print(f"The specified directory does not exist: {directory}")
        return
    
    # List to hold DataFrame rows
    rows = []

    # Loop through all files in the directory
    for filename in os.listdir(directory):
        if filename.endswith(".csv"):
            filepath = os.path.join(directory, filename)
            try:
                # Read the CSV file
                df = pd.read_csv(filepath)
                rows.append({'Filename': filename, 'Columns': df.columns.tolist()})
            except Exception as e:
                print(f"Error reading file {filename}: {e}")
                continue

    # Check if there are no CSV files or all failed
    if not rows:
        print("No valid CSV files found in the directory.")
        return

    # Expand columns into separate fields and prepare data for a DataFrame
    data = []
    max_columns = max(len(row['Columns']) for row in rows)
    for row in rows:
        aligned_columns = row['Columns'] + [''] * (max_columns - len(row['Columns']))
        data.append([row['Filename']] + aligned_columns)

    # Define the column names
    column_names = ['Filename'] + [f'Column_{i+1}' for i in range(max_columns)]

    # Create a DataFrame and use pd.concat for flexibility
    result_df = pd.concat([pd.DataFrame([row], columns=column_names) for row in data], ignore_index=True)

    # Save the result to an Excel file
    try:
        result_df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"Excel file saved successfully: {output_file}")
    except Exception as e:
        print(f"Error saving Excel file: {e}")

# Specify the directory containing the CSV files
directory = r"C:\Users\YourUsername\Desktop\CSV_Files"

# Specify the output Excel file
output_file = r"C:\Users\YourUsername\Desktop\output.xlsx"

# List CSV filenames and columns and save to Excel
list_csv_columns_to_excel(directory, output_file)

# Generate user notification once file is saved
print('Your output file is saved.')