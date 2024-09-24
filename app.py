import pytesseract
from PIL import Image
import pandas as pd
import os

# Function to extract product details from an image
def extract_items_from_image(image_path):
    # Open the image using PIL
    img = Image.open(image_path)
    
    # Use pytesseract to extract text from the image
    text = pytesseract.image_to_string(img)
    
    # Process the extracted text to find products and quantities
    product_items = {}
    for line in text.split('\n'):
        line = line.strip().lower()
        if 'kg' in line:
            parts = line.split()
            if len(parts) >= 3:  # Example: ['rice', '1', 'kg']
                product_name = parts[0]
                quantity = float(parts[1])
                product_items[product_name] = quantity
    
    return product_items

# Function to calculate the total amount based on product prices from the Excel file
def calculate_total(product_items, excel_file):
    # Read the Excel file
    product_rates = pd.read_excel(excel_file)
    
    # Create a dictionary to store product rates from the Excel file
    product_rates_dict = {}
    for index, row in product_rates.iterrows():
        product_rates_dict[row['Product']] = row['Price_per_kg']
    
    # Calculate total amount for each product
    bill_data = []
    
    for product, qty in product_items.items():
        price_per_kg = product_rates_dict.get(product)
        if price_per_kg:
            amount = qty * price_per_kg
            bill_data.append([product.capitalize(), qty, price_per_kg, amount])
        else:
            bill_data.append([product.capitalize(), qty, 'N/A', 'Price not found'])
    
    return bill_data

# Function to save the bill data to a new Excel file
def save_bill_to_excel(bill_data, output_file):
    # Create a DataFrame to hold the bill data
    df = pd.DataFrame(bill_data, columns=['Product', 'Quantity (kg)', 'Price per kg (USD)', 'Total Amount (USD)'])
    
    # Save the DataFrame to an Excel file
    df.to_excel(output_file, index=False)
    print(f"Bill saved to {output_file}")

# Main function to execute the workflow
def main(image_path, excel_file, output_file):
    # Step 1: Extract products and quantities from the image
    product_items = extract_items_from_image(image_path)
    
    # Step 2: Calculate the total amount based on the Excel file
    bill_data = calculate_total(product_items, excel_file)
    
    # Step 3: Save the bill to a new Excel file
    save_bill_to_excel(bill_data, output_file)

# Run the program
if __name__ == "__main__":
    # Example: Paths to the image, Excel file, and output Excel file
    image_path = 'products_image.jpg'  # The image containing product details
    excel_file = 'product_rate_list.xlsx'  # The Excel file containing product rates
    output_file = 'product_bill.xlsx'  # The Excel file where the bill will be saved

    # Make sure the files exist before proceeding
    if os.path.exists(image_path) and os.path.exists(excel_file):
        main(image_path, excel_file, output_file)
    else:
        print("Please check the file paths. The image or Excel file is missing.")
