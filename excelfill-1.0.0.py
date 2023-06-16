import pandas as pd

# Christian Diaz
# Filling information from one Excel sheet to another
# Version 1.0.0
# 6/15/23

# Load the excel file using the Pandas ExcelFile function. Make sure to use the correct path where the file is located.
xl = pd.ExcelFile("C:\\Users\\CDiaz\\OneDrive - HITT Contracting Inc\\Documents\\Python\\InvoiceMerge\\the-file-of-dreams.xlsx")

# Read the 'SSRS-SL-IND002XLS Sales by Cust' sheet from the Excel file into a DataFrame.
sales_df = xl.parse('SSRS-SL-IND002XLS Sales by Cust')

# Set the second row as the header.
sales_df.columns = sales_df.iloc[1]

# Discard the first two rows and keep the rest.
sales_df = sales_df.iloc[2:]

# Read the 'Already in AP' and 'Add to AP' sheets from the Excel file into separate DataFrames.
already_in_ap_df = xl.parse('Already in AP')
add_to_ap_df = xl.parse('Add to AP')

# Define a function named fill_missing_info that takes a source DataFrame (source_df) and a target DataFrame (target_df) as parameters.
# The function iterates over each row in the target DataFrame, checks if the 'Serial Number' from the source DataFrame matches any serial numbers in the target DataFrame,
# and if a match is found, fills in the 'Invoice Number', 'Cost', and 'Purchase Date' columns in the target DataFrame using the corresponding values from the source DataFrame.
def fill_missing_info(source_df, target_df):
    # For each row in the target DataFrame
    for index, row in target_df.iterrows():
        # Split the 'Serial Number' column on commas (in case there are multiple serial numbers in a cell)
        serials = str(row['Serial Number']).split(',')
        # For each serial number in the split serial numbers
        for serial in serials:
            # Find rows in the source DataFrame where the 'Serial Number' column contains the current serial number
            match = source_df[source_df['Serial Number'].str.contains(serial, na=False)]
            # If a match is found
            if not match.empty:
                # Fill in the 'Invoice Number', 'Cost', and 'Purchase Date' columns in the current row of the target DataFrame 
                # using the corresponding values from the first row of the matched rows in the source DataFrame.
                target_df.loc[index, 'Invoice Number'] = match['Invoice Nbr'].values[0]
                target_df.loc[index, 'Cost'] = match['Unit Price'].values[0]
                target_df.loc[index, 'Purchase Date'] = match['Invoice Date'].values[0]
    # Return the modified target DataFrame
    return target_df

# Apply the fill_missing_info function to both the 'Already in AP' and 'Add to AP' sheets, using the 'SSRS-SL-IND002XLS Sales by Cust' sheet as the source.
already_in_ap_df = fill_missing_info(sales_df, already_in_ap_df)
add_to_ap_df = fill_missing_info(sales_df, add_to_ap_df)

# Write the modified DataFrames back to new sheets in the Excel file, overwriting the original sheets.
with pd.ExcelWriter('C:\\Users\\CDiaz\\OneDrive - HITT Contracting Inc\\Documents\\Python\\InvoiceMerge\\output.xlsx') as writer:
    already_in_ap_df.to_excel(writer, sheet_name='Already in AP', index=False)
    add_to_ap_df.to_excel(writer, sheet_name='Add to AP', index=False)
