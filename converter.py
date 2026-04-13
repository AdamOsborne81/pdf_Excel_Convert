import camelot
import pandas as pd
from openpyxl import Workbook

# Advanced PDF to Excel conversion

def convert_pdf_to_excel(pdf_path, excel_path):
    # Use camelot to read tables from the PDF while preserving layout and formatting
    tables = camelot.read_pdf(pdf_path, pages="all")

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for i, table in enumerate(tables):
            # Convert the table to a DataFrame and write it to a specific sheet
            df = table.df
            df.to_excel(writer, sheet_name=f'Sheet{i + 1}', index=False)

            # Apply formatting to each sheet (example: set column width)
            worksheet = writer.sheets[f'Sheet{i + 1}']
            for column in df.columns:
                max_length = max(df[column].astype(str).map(len).max(), len(column)) + 2
                adjusted_width = max_length
                worksheet.column_dimensions[column].width = adjusted_width

# Example usage of the function
# convert_pdf_to_excel('input.pdf', 'output.xlsx')
