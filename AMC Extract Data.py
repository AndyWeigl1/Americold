import PyPDF2
import win32com.client
import uuid
import re
from PyPDF2 import PdfReader
import pyautogui
import time
import xlwings as xw
import win32gui
import os
import shutil
from win10toast import ToastNotifier


def extract_invoices_from_emails(output_folder):
    # Connect to Outlook
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook_app.GetNamespace("MAPI")

    # Get the selected emails in Outlook
    selected_items = outlook_app.ActiveExplorer().Selection
    num_files_created = 0

    # Iterate over each selected email
    for email in selected_items:
        # Check if the email has attachments
        if email.Attachments.Count == 0:
            print(f"The email '{email.Subject}' does not have any attachments.")
            continue

        # Iterate over each attachment in the email
        for attachment in email.Attachments:
            if attachment.FileName.lower().endswith('.pdf'):
                # Save the attachment to a temporary file
                temp_file_path = os.path.join(output_folder, 'temp.pdf')
                attachment.SaveAsFile(temp_file_path)

                # Extract invoices from the PDF
                num_files_created += extract_invoices_from_pdf(temp_file_path, output_folder)

                # Delete the temporary file
                os.remove(temp_file_path)


def extract_invoices_from_pdf(input_pdf_path, output_folder):
    pdf_file = open(input_pdf_path, 'rb')
    pdf_reader = PyPDF2.PdfReader(pdf_file)

    num_files_created = 0

    # Iterate over each page in the PDF
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text = page.extract_text()

        # Check if the page contains the invoice number or the specific phrase
        if re.search(r'S\d{7}', text):
            # Create a new PDF writer and add the current page
            pdf_writer = PyPDF2.PdfWriter()
            pdf_writer.add_page(page)

            # Generate a unique filename for the extracted invoice
            unique_identifier = str(uuid.uuid4())[:8]
            output_file_name = f'invoice_{unique_identifier}.pdf'
            output_file_path = os.path.join(output_folder, output_file_name)

            # Save the extracted invoice as a new PDF with the unique filename
            with open(output_file_path, 'wb') as output_file:
                pdf_writer.write(output_file)

            num_files_created += 1

    pdf_file.close()

    return num_files_created


# Example usage
output_folder = r"C:\Users\Andy Weigl\OneDrive - Kodiak Cakes\Americold\Invoices"

extract_invoices_from_emails(output_folder)


def extract_invoice_number(text):
    pattern = r'S\d{7}'
    match = re.search(pattern, text)
    if match:
        return match.group()
    else:
        return None


def extract_date(text):
    pattern = r'\b\d{1,2}/\d{1,2}/\d{2}\b'
    match = re.search(pattern, text)
    if match:
        date = match.group()
        modified_date = date.replace('/', '.')  # Replace "/" with "."
        return modified_date
    else:
        return None


def extract_location(text):
    locations = {
        'ALLENTOWN': 'Allentown',
        'CLEARFIELD': 'Clearfield',
        'ATLANTA': 'Atlanta',
        'CARTHAGE': 'Carthage'
    }
    for location, formatted_location in locations.items():
        if location in text:
            return formatted_location
    return None


# Function to extract the data from the text
def extract_data(text):
    invoice_number = extract_invoice_number(text)
    date = extract_date(text)
    location = extract_location(text)

    return invoice_number, date, location


folder_path = "C:\\Users\\Andy Weigl\\OneDrive - Kodiak Cakes\\Americold\\Invoices"

for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):
        pdf_file_path = os.path.join(folder_path, filename)

        # Extract the data from the PDF
        with open(pdf_file_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            extracted_text = ''
            for page in pdf_reader.pages:
                extracted_text += page.extract_text()

        invoice_number, date, location = extract_data(extracted_text)

        invoice_data = extract_data(extracted_text)
        print(invoice_data)

        # Generate the new filename
        new_filename = f'{invoice_number} {location} {date} .pdf'
        print(new_filename)

        # Rename the file
        new_file_path = os.path.join(folder_path, new_filename)
        shutil.move(pdf_file_path, new_file_path)


# create an object to ToastNotifier class
toaster = ToastNotifier()

# Issue the notification with your custom parameters
toaster.show_toast("AMC Extract Data", "All invoices have been moved and renamed!", duration=1)


def extract_invoice_number2(text):
    pattern = r'S\d{7}'
    match = re.search(pattern, text)
    if match:
        return match.group()
    else:
        return None


def extract_po_number2(text):
    lines = text.split('\n')
    kodiak_cakes_indices = [i for i, line in enumerate(lines) if 'KODIAK CAKES' in line]
    if len(kodiak_cakes_indices) >= 2:
        po_line = lines[kodiak_cakes_indices[1] - 1]
        po_number = po_line.strip()
        if 'PO' not in po_number:
            po_number = 'PO' + po_number
        return po_number
    else:
        return None


def extract_date2(text):
    pattern = r'\b\d{1,2}/\d{1,2}/\d{2}\b'
    match = re.search(pattern, text)
    if match:
        return match.group()
    else:
        return None


def extract_amount2(text):
    pattern = r'TOT HND:\s+([\d.]+)'
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    else:
        return None


def extract_location2(text):
    locations = {
        'ALLENTOWN': 'Allentown',
        'CLEARFIELD': 'Clearfield',
        'ATLANTA': 'Atlanta',
        'CARTHAGE': 'Carthage'
    }
    for location, formatted_location in locations.items():
        if location in text:
            return formatted_location
    return None


# Function to extract the data from the text
def extract_data2(text):
    invoice_number = extract_invoice_number2(text)
    po_number = extract_po_number2(text)
    date = extract_date2(text)
    amount = extract_amount2(text)
    location = extract_location2(text)

    return invoice_number, po_number, date, amount, location


def clear_excel_cells(worksheet, cell_positions):
    for cell_position in cell_positions.values():
        try:
            worksheet.Range(cell_position).Value = None
        except Exception as e:
            print(f"Failed to clear cell {cell_position}: {e}")
            raise  # re-throw the exception to halt the program

# Folder path containing the PDF files
folder_path = "C:\\Users\\Andy Weigl\\OneDrive - Kodiak Cakes\\Americold\\Invoices"

# Excel file path
excel_file_path = r"C:\Users\Andy Weigl\OneDrive - Kodiak Cakes\Excel Macro Workbooks\Americold SA Receipt Bills.xlsm"

# Excel worksheet name
worksheet_name = "SA (Receipt)"

# Cell positions in Excel
cell_positions = {
    "invoice_number": "C3",
    "po_number": "C7",
    "date": "C6",
    "amount": "C5",
    "location": "C4"
}

# Create an instance of Excel
Excel = win32com.client.Dispatch('Excel.Application')

# To run Excel in the background, set Visible to False
Excel.Visible = False

# Open Your Workbook
wb = Excel.Workbooks.Open(r"C:\Users\Andy Weigl\OneDrive - Kodiak Cakes\Excel Macro Workbooks\Americold SA Receipt Bills.xlsm")
worksheet = wb.Sheets["SA (Receipt)"]

# Loop over PDF files
for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):
        pdf_file_path = os.path.join(folder_path, filename)

        # Extract the data from the PDF
        with open(pdf_file_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            extracted_text = ''
            for page in pdf_reader.pages:
                extracted_text += page.extract_text()

        invoice_number, po_number, date, amount, location = extract_data2(extracted_text)

        invoice_data = extract_data2(extracted_text)
        print(invoice_data)

        invoice_number, po_number, date, amount, location = invoice_data

        # Clear the previous cell values
        clear_excel_cells(worksheet, cell_positions)

        # Write the data to the Excel cells
        values = [invoice_number, po_number, date, amount, location]
        cell_names = ['C3', 'C7', 'C6', 'C5', 'C4']

        for value, cell_name in zip(values, cell_names):
            worksheet.Range(cell_name).Value = value

        # Run Macro
        Excel.Application.Run("CopyAndPasteValues")

# Save your workbook
wb.Save()

# Quit Excel
Excel.Application.Quit()

excel_file_path = r"C:\Users\Andy Weigl\OneDrive - Kodiak Cakes\Excel Macro Workbooks\Americold SA Receipt Bills.xlsm"

os.startfile(excel_file_path)