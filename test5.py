import pandas as pd

def clean_all_cells(excel_path, cleaned_excel_path, characters_to_remove):
    # Read the Excel file
    df = pd.read_excel(excel_path, sheet_name='Product List')  # Adjust the sheet name as needed
    
    # Print original data for debugging
    print("Original Data:\n", df.head())  
    
    # Function to clean a single cell
    def clean_cell(cell):
        try:
            # Ensure the cell is a string before processing
            if isinstance(cell, str):
                return ''.join(char for char in cell if char not in characters_to_remove)
            else:
                return cell  # Return original value if it's not a string
        except Exception as e:
            print(f"Error cleaning cell '{cell}': {e}")
            return cell  # Return original value in case of error

    # Apply the cleaning function to each cell in the DataFrame
    cleaned_df = df.applymap(clean_cell)

    # Print cleaned data for debugging
    print("Cleaned Data:\n", cleaned_df.head())  
    
    # Save the cleaned data to a new Excel file
    cleaned_df.to_excel(cleaned_excel_path, index=False)
    print(f"Cleaned data saved to {cleaned_excel_path}")

# Example usage
excel_path = 'product_list_with_image.xlsx'  # Original Excel file
cleaned_excel_path = 'cleaned_product_list.xlsx'  # New Excel file for cleaned data
characters_to_remove = set(r'!@#$%^&*(),.?":{}|<>')  # Specify the characters to remove (set)

clean_all_cells(excel_path, cleaned_excel_path, characters_to_remove)
