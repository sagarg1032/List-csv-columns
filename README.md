# List CSV Columns to Excel

This Python script reads all `.csv` files in a specified directory, extracts their column names, and saves the results to an Excel file. This can be useful for validating or documenting column names across multiple CSV files.

## Features
- Automatically detects and processes all `.csv` files in a given folder.
- Extracts column names from each file and aligns them in a single Excel sheet.
- Pads rows with fewer columns to match the maximum number of columns.
- Saves the results as an Excel file with filename and column names clearly labeled.

## Requirements
- Python 3.x
- Required libraries:
  - `pandas`
  - `openpyxl`

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/list-csv-columns.git
   ```
2. Navigate to the project directory:
   ```bash
   cd list-csv-columns
   ```
3. Install the required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Update the `directory` and `output_file` variables in the script to match your directory and desired output file path.

   Example:
   ```python
   directory = r"C:\Users\YourUsername\Desktop\CSV_Files"
   output_file = r"C:\Users\YourUsername\Desktop\output.xlsx"
   ```

2. Run the script:
   ```bash
   python list_csv_columns_to_excel.py
   ```

3. Check the specified Excel file for the output.

## Output Format
The Excel file contains:
- **Filename**: Name of each processed CSV file.
- **Column_1, Column_2, ... Column_N**: Column names from each CSV file.

## Example
### Input:
Folder contains the following CSV files:
- `file1.csv` with columns `['A', 'B', 'C']`
- `file2.csv` with columns `['X', 'Y']`

### Output:
| Filename | Column_1 | Column_2 | Column_3 |
|----------|----------|----------|----------|
| file1.csv| A        | B        | C        |
| file2.csv| X        | Y        |          |


## Contributing
Contributions are welcome! Feel free to open issues or submit pull requests.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.