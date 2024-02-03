import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment
import pandas as pd
import os

file_pdf_name = "FileNameHere"
pdf_input_file_path = file_pdf_name + ".pdf"  
pdf_output_file_path_xlsx = file_pdf_name + ".xlsx"

# Open the PDF file using pdfplumber
with pdfplumber.open(pdf_input_file_path) as pdf:
    # Open the Excel file using openpyxl
    wb = Workbook()
    ws = wb.active

    # Set sheet name to file_pdf_name
    ws.title = file_pdf_name

    # Loop through all pages
    for page_num in range(len(pdf.pages)):
        current_page = pdf.pages[page_num]
        table = current_page.extract_table()

        # Convert the table to a pandas DataFrame
        df = pd.DataFrame(table[1:], columns=table[0])

        # Convert the DataFrame to a list of lists
        data = [df.columns.tolist()] + df.values.tolist()

        # Write the data to the Excel file
        for row in data:
            ws.append(row)

        # Auto-fit columns for the current page
        for column in ws.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column[0].column_letter].width = adjusted_width

        # Auto-fit rows for the current page
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True)
                if cell.value is not None and ws.row_dimensions[row[0].row].height is not None:
                    ws.row_dimensions[row[0].row].height = 1.2 * ws.row_dimensions[row[0].row].height

        # Add an empty row between pages
        ws.append([])

    # Save the modified Excel file
    wb.save(pdf_output_file_path_xlsx)
