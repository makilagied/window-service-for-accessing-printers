import os
import win32print
import tempfile
import win32ui
from openpyxl import load_workbook
from fpdf import FPDF
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from flask import Flask, request, jsonify
from pathlib import Path
from flask_cors import CORS

# Function to get the Downloads directory path
def get_downloads_folder():
    try:
        if os.name == 'nt':  # For Windows
            path = os.path.join(os.environ['USERPROFILE'], 'Downloads')
        else:
            path = str(Path.home() / 'Downloads')
        print(f"Downloads folder path: {path}")
        return path
    except Exception as e:
        print(f"Error getting downloads folder: {e}")
        raise

# Function to process Excel template
def update_excel_with_data(data):
    try:
        print("Starting Excel file update process.")
        template_path = "invoice.xlsx"  # Path to the Excel template
        if not os.path.exists(template_path):
            raise FileNotFoundError("Invoice template not found.")
        print(f"Using template: {template_path}")

        workbook = load_workbook(template_path)
        sheet = workbook.active

        # Update specified cells with data
        sheet["F8"] = data.get("trade_date", "N/A")
        sheet["F9"] = data.get("order_number", "N/A")
        sheet["B18"] = data.get("client", "N/A")
        sheet["E25"] = data.get("subtotal", 0)
        sheet["F27"] = data.get("brokerage", 0)
        sheet["F28"] = data.get("dse", 0)
        sheet["F29"] = data.get("cmsa", 0)
        sheet["F30"] = data.get("cds", 0)
        sheet["F31"] = data.get("fidelity", 0)
        sheet["F34"] = data.get("subtotal", 0)
        sheet["F35"] = data.get("vat", 0)
        sheet["F37"] = data.get("total_fees", 0)

        # Save the updated file
        downloads_folder = get_downloads_folder()
        updated_excel_path = os.path.join(downloads_folder, "updated_invoice.xlsx")
        workbook.save(updated_excel_path)
        workbook.close()
        print(f"Updated Excel file saved at: {updated_excel_path}")
        return updated_excel_path
    except Exception as e:
        print(f"Error updating Excel file: {e}")
        raise

# Function to print the updated Excel file
def print_excel_file(file_path, printer_name=None):
    try:
        print(f"Attempting to print file: {file_path}")

        # Check if the file exists
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found for printing at path: {file_path}")

        # Select printer
        if not printer_name:
            print("No printer specified. Fetching default printer.")
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            if printers:
                printer_name = printers[0][2]  # Select the first available printer
                print(f"Default printer selected: {printer_name}")
            else:
                raise Exception("No printers available.")
        else:
            print(f"Using specified printer: {printer_name}")

        # Verify printer is available
        printer_info = win32print.GetPrinter(win32print.OpenPrinter(printer_name))
        if not printer_info:
            raise Exception(f"Printer '{printer_name}' is not available or offline.")

        # Use win32print to send a file to the printer (Excel method is replaced)
        hPrinter = win32print.OpenPrinter(printer_name)
        printer_job = win32print.StartDocPrinter(hPrinter, 1, ("Excel Print Job", None, "RAW"))
        win32print.StartPagePrinter(hPrinter)
        win32print.WritePrinter(hPrinter, open(file_path, "rb").read())
        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)
        win32print.ClosePrinter(hPrinter)

        print(f"File '{file_path}' sent to printer: {printer_name}")

    except Exception as e:
        print(f"Error printing Excel file: {e}")
        raise



# def print_text_file(content, printer_name=None):
#     try:
#         # Create a temporary TXT file
#         temp_dir = tempfile.gettempdir()
#         temp_file_path = os.path.join(temp_dir, "temp_print_file.txt")
        
#         with open(temp_file_path, "w", encoding="utf-8") as temp_file:
#             temp_file.write(content)
        
#         print(f"Text file created at: {temp_file_path}")

#         # Set the default printer if not specified
#         if not printer_name:
#             printer_name = win32print.GetDefaultPrinter()
#             print(f"Using default printer: {printer_name}")
#         else:
#             print(f"Using specified printer: {printer_name}")

#         # Open the printer
#         hPrinter = win32print.OpenPrinter(printer_name)
#         try:
#             # Start a print job
#             job_name = "Text File Print Job"
#             hJob = win32print.StartDocPrinter(hPrinter, 1, (job_name, None, "RAW"))
#             win32print.StartPagePrinter(hPrinter)

#             # Read the TXT file and send it to the printer
#             with open(temp_file_path, "rb") as temp_file:
#                 data = temp_file.read()
#                 win32print.WritePrinter(hPrinter, data)

#             win32print.EndPagePrinter(hPrinter)
#             win32print.EndDocPrinter(hPrinter)
#             print(f"Text file printed successfully on {printer_name}.")
#         finally:
#             win32print.ClosePrinter(hPrinter)

#         # Cleanup temporary file
#         # os.remove(temp_file_path)
#         print("Temporary text file deleted.")
#     except Exception as e:
#         print(f"Error printing text file: {e}")
#         raise


def list_printers():
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    for printer in printers:
        print(printer[2])  # Printer name


# Flask setup
app = Flask(__name__)
CORS(app)

@app.route('/run-receipt-script', methods=['POST'])
def run_invoice_script():
    try:
        print("Received request for invoice processing.")
        data = request.json
        if not data:
            print("No data received in the request.")
            return jsonify({"error": "No data provided"}), 400
        

        list_printers()

        print(f"Request data: {data}")
        # Process Excel file
        updated_excel_path = update_excel_with_data(data)

        # Print the Excel file
        print_excel_file(updated_excel_path, data.get("printer_name", None))

        # Send email with the Excel file
        # send_email_with_attachment(data, updated_excel_path)

        print("Invoice processing completed successfully.")
        return jsonify({"message": "Invoice processed successfully."})
    except Exception as e:
        print(f"Error processing invoice: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    print("Starting Flask server...")
    app.run(debug=True, host='0.0.0.0', port=7860)