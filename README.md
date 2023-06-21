# ExcelFill
Filling two excel sheets with missing information from one to another

Key Features
Reads an Excel file and parses each sheet into a separate pandas DataFrame.
Identifies missing information in the target sheets based on specific columns ('Invoice Number', 'Cost', 'Purchase Date').
Checks for matches in the source sheet based on the 'Serial Number' column.
Fills the missing information in the target sheets with corresponding data from the source sheet.
Writes the updated DataFrames back to new sheets in the Excel file, overwriting the original sheets.

Installation
To run this script, you will need Python installed along with the following Python packages:
pandas
openpyxl

You can install these packages using pip:
pip install pandas openpyxl

Usage
Please ensure to replace '/path/to/your/excel/file.xlsx' and '/path/to/your/output/file.xlsx' with the paths to your actual Excel files.

Refer to the provided example in the repository for more detailed usage instructions.

Known Issues
The current release of Excel Fill v1.0.0 has identified the following issues:

Date Format Inconsistency: The date format in the output Excel file is inconsistent across sheets. The first sheet formats dates as MM/DD/YYYY, but the second sheet formats dates as YYYY/DD/MM HH:MM:SS.

Price Format Inconsistency: The format of the price is not consistent across sheets. In the first sheet, prices are formatted correctly (e.g., 1700.00), but in the second sheet, the prices lack the decimal and subsequent two digits (e.g., 1700 instead of 1700.00).

Please note that these issues can be manually corrected in the output .xlsx file after the script has been run.

If you encounter any additional issues, please open an issue in the GitHub repository. Contributions for fixing these known issues are also welcomed.

Contributing
Contributions are welcomed. Please fork this repository and open a pull request to add more features or fix issues.

Conclusion
Excel Fill is ideal for those who work with large Excel files and need an automated way to ensure data consistency across sheets. It is flexible and can be adapted to work with different source/target sheets and different column names.
