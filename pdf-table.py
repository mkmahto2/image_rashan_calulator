import os
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

# Function to calculate the total bill and generate a PDF
def calculate_total_bill_and_generate_pdf(product_excel_path, rate_excel_path, pdf_output_path):
    # Read the product and rate list Excel files
    product_df = pd.read_excel(product_excel_path, sheet_name='Product List')
    rate_df = pd.read_excel(rate_excel_path, sheet_name='Rate List')

    # Merge product and rate data
    merged_df = pd.merge(product_df, rate_df, on='Product Name', how='inner')
    merged_df['Quantity'] = pd.to_numeric(merged_df['Quantity'], errors='coerce')
    merged_df['Final Cost'] = merged_df['Rate'] * merged_df['Quantity']

    # Create table data for the PDF
    table_data = [['Product Name', 'Quantity', 'Rate', 'Final Cost']]
    for index, row in merged_df.iterrows():
        table_data.append([row['Product Name'], row['Quantity'], row['Rate'], row['Final Cost']])
    
    total_bill = merged_df['Final Cost'].sum()
    table_data.append(['', '', 'Total Bill', total_bill])

    # Generate PDF
    pdf = SimpleDocTemplate(pdf_output_path, pagesize=letter)
    table = Table(table_data)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ])
    table.setStyle(style)

    pdf.build([table])
    print(f"PDF created successfully at: {pdf_output_path}")

# Example usage to generate the PDF bill
product_excel_path = 'product_list_with_image.xlsx'  # Path to your product list Excel file
rate_excel_path = 'rate_list.xlsx'        # Path to your rate list Excel file
pdf_output_path = 'product_bill.pdf'      # Path to save the PDF file

# Generate the PDF
calculate_total_bill_and_generate_pdf(product_excel_path, rate_excel_path, pdf_output_path)
