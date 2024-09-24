import pandas as pd

from openpyxl import load_workbook

def delete_row_from_excel(file_path, row_number):
    # Load the existing workbook
    workbook = load_workbook(file_path)
    
    # Select the active worksheet
    sheet = workbook.active

    # Delete the specified row
    sheet.delete_rows(row_number)

    # Save the workbook after deletion
    workbook.save(file_path)

# Specify the path to your Excel file
file_path = 'product_list_with_image.xlsx'

# Delete row 2 from the Excel file
delete_row_from_excel(file_path, 2)

print("Row 2 has been deleted from the Excel file.")

def merge_product_and_rate(product_excel_path, rate_excel_path, output_excel_path):
    # Read the product list and rate list Excel files
    product_df = pd.read_excel(product_excel_path, sheet_name='Product List')  # Adjust the sheet name if needed
    rate_df = pd.read_excel(rate_excel_path, sheet_name='Rate List')  # Adjust the sheet name if needed
    
    # Print original data for debugging
    print("Product Data:\n", product_df.head())
    print("Rate Data:\n", rate_df.head())
    
    
    # # Merge the product and rate DataFrames on 'Product Name'
    merged_df = pd.merge(product_df, rate_df, how='left', on='Product Name')

    # # Print merged data for debugging
    print("Merged Data:\n", merged_df.head())
    
     # Replace NaN values with 0 in the 'Quantity' and 'Rate' columns
    merged_df['Quantity'] = merged_df['Quantity'].fillna(0)
    merged_df['Rate'] = merged_df['Rate'].fillna(0)
    
    # # Calculate Total Bill
    merged_df['Total'] = merged_df['Quantity'] * merged_df['Rate']  # Assuming the column names in merged_df are 'Quantity' and 'Rate'
    
    # # Print final data for debugging
    print("Final Data with Total:\n", merged_df.head())

    # # Save the final DataFrame to a new Excel file
    merged_df.to_excel(output_excel_path, index=False)
    print(f"Final bill data saved to {output_excel_path}")

# Example usage
product_excel_path = 'product_list_with_image.xlsx'  # Path to the product list Excel file
rate_excel_path = 'rate_list.xlsx'  # Path to the rate Excel file
output_excel_path = 'final_bill.xlsx'  # Output Excel file for the final bill

merge_product_and_rate(product_excel_path, rate_excel_path, output_excel_path)
