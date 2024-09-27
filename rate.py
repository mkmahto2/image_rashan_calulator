import pandas as pd

def calculate_total_bill(product_excel_path, rate_excel_path):
    # Read the product list (product name and quantity) from the Product List Excel
    product_df = pd.read_excel(product_excel_path, sheet_name='Product List')

    # Read the rate list (product name and rate) from the Rate Excel
    rate_df = pd.read_excel(rate_excel_path, sheet_name='Rate List')

    # Ensure column names are correctly interpreted (change the column names if different)
    # Product List should have columns 'Product Name' and 'Quantity'
    # Rate List should have columns 'Product Name' and 'Rate'
    
    # Merge the two DataFrames on 'Product Name'
    merged_df = pd.merge(product_df, rate_df, on='Product Name', how='inner')

    # Ensure 'Quantity' column is numeric
    merged_df['Quantity'] = pd.to_numeric(merged_df['Quantity'], errors='coerce')
    
    # Calculate the final cost for each product (Rate * Quantity)
    merged_df['Final Cost'] = merged_df['Rate'] * merged_df['Quantity']

    # Print merged data with product name, quantity, rate, and final cost
    print("\nMerged Product Details (with Rates):")
    print(merged_df[['Product Name', 'Quantity', 'Rate', 'Final Cost']])

    # Calculate total bill
    total_bill = merged_df['Final Cost'].sum()

    # Print total bill
    print(f"\nTotal Bill: {total_bill:.2f}")

    return total_bill

# Example usage
product_excel_path = 'product_list_with_image.xlsx'  # Path to the product list Excel file
rate_excel_path = 'rate_list.xlsx'        # Path to the rate list Excel file

calculate_total_bill(product_excel_path, rate_excel_path)
