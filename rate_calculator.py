import pandas as pd

# Function to create the bill by multiplying quantities with rates
def create_bill(product_excel_path, rate_excel_path, bill_excel_path):
    # Read the product list
    product_df = pd.read_excel(product_excel_path, sheet_name='Product List')
    print("Product Data:\n", product_df)  # Debugging output

    # Read the rates
    rate_df = pd.read_excel(rate_excel_path)
    print("Rate Data:\n", rate_df)  # Debugging output

    # Merge the product data with the rates on the product name
    bill_df = pd.merge(product_df, rate_df, how='left', left_on='Product Name', right_on='Product Name')

    # Calculate total cost
    bill_df['Total Cost'] = bill_df['Quantity'] * bill_df['Rate']

    # Select relevant columns
    bill_df = bill_df[['Product Name', 'Quantity', 'Rate', 'Total Cost']]

    # Save the bill to an Excel file
    bill_df.to_excel(bill_excel_path, index=False)
    print(f"Bill saved to {bill_excel_path}")

# Example usage
product_excel_path = 'product_list_with_image.xlsx'  # The Excel file containing product data
rate_excel_path = 'product_rates.xlsx'  # The Excel file containing product rates
bill_excel_path = 'bill_output.xlsx'  # The output Excel file for the bill

create_bill(product_excel_path, rate_excel_path, bill_excel_path)
