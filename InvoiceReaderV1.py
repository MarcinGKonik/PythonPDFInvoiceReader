import pdfplumber
import pandas as pd
import os

def extract_column_content_with_header_inheritance(pdf_path, column_name):
    """Extract content from the specified column across all pages in the PDF, even when headers are missing on subsequent pages."""
    column_content = []
    headers = None  # Store the headers from the first occurrence
    col_index = None  # Store the index of the desired column
    
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        print(f"Total number of pages in the PDF: {total_pages}")
        
        for page_num, page in enumerate(pdf.pages, start=1):
            print(f"\nProcessing page {page_num} of {total_pages}")
            tables = page.extract_tables()

            if not tables:
                print(f"No tables found on page {page_num}")
                continue

            for table_num, table in enumerate(tables, start=1):
                # If headers are not yet found, try to extract them
                if headers is None:
                    headers = table[0]
                    print(f"Found headers on page {page_num}: {headers}")

                    if column_name in headers:
                        col_index = headers.index(column_name)
                        print(f"Column '{column_name}' found at index {col_index}")
                    else:
                        print(f"Column '{column_name}' not found in headers on page {page_num}")
                        continue

                for row in table[1:]:
                    if len(row) > col_index:
                        column_content.append(row[col_index])
                    else:
                        print(f"Row {row} does not contain enough columns on page {page_num}")

            # If headers are missing on this page, continue using the previous headers
            if headers and col_index is not None:
                print(f"Continuing to use headers from previous pages on page {page_num}")

    return column_content

# File and column definitions
pdf_path = "xyz.pdf"
column_name_nazwa = "columnHeaderX"
column_name_ilosc = "columnHeaderY"
column_name_cena = "columnNameZ"

# Extract data from the PDF
content_nazwa = extract_column_content_with_header_inheritance(pdf_path, column_name_x)
content_ilosc = extract_column_content_with_header_inheritance(pdf_path, column_name_y)
content_cena = extract_column_content_with_header_inheritance(pdf_path, column_name_z)

# Create a DataFrame with the extracted data
df = pd.DataFrame({
    column_name_nazwa: content_x,
    column_name_ilosc: content_y,
    column_name_cena: content_z
})

# Check if the Excel file exists and append new data if it does
excel_path = "output.xlsx"
if os.path.exists(excel_path):
    existing_data = pd.read_excel(excel_path)
    updated_data = pd.concat([existing_data, df], ignore_index=True)
else:
    updated_data = df

# Export to Excel
updated_data.to_excel(excel_path, index=False)
print(f"\nData successfully exported to {excel_path}")

