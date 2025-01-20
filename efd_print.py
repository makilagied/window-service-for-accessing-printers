import os
import sys
import win32print
import win32api
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
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
from datetime import datetime
import logging


# Configure logging
logging.basicConfig(filename='app.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')




def resource_path(relative_path):
    """Get the absolute path to a resource, whether running in development or as a packaged executable."""
    try:
        # PyInstaller creates a temp folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Example usage
logo_path = resource_path("logo.png")


# Function to get the Downloads directory path
def get_downloads_folder():
    try:
        if os.name == 'nt':  # For Windows
            path = os.path.join(os.environ['USERPROFILE'], 'Downloads')
        else:
            path = str(Path.home() / 'Downloads')
        logging.info(f"Downloads folder path: {path}")
        return path
    except Exception as e:
        logging.debug(f"Error getting downloads folder: {e}")
        raise


def format_number(value):
    try:
        # Convert the value to a float (if possible), then format it
        value = float(value)
        return f"{value:,.2f}"  # Format number with thousand separators
    except ValueError:
        # If conversion fails, return 0 formatted
        return "0.00"


def update_excel_with_data(data):
    try:
        logging.info("Starting Excel file update process.")
        # template_path = "invoice.xlsx"
        template_path = resource_path("invoice.xlsx")
        if not os.path.exists(template_path):
            raise FileNotFoundError("Invoice template not found.")
        logging.info(f"Using template: {template_path}")

        # Load the workbook and keep formatting
        workbook = load_workbook(template_path)
        sheet = workbook.active

        # Insert image into cells B2 to B6
        # logo_path = "logo.png"  # Path to the logo image
        # if os.path.exists(logo_path):
        #     # Create an Image object
        #     img = Image(logo_path)
        #     # Set the image width and height (optional, adjust as necessary)
        #     img.width = 180
        #     img.height = 100
        #     # Position the image at cell B2 (adjust the cell if needed)
        #     sheet.add_image(img, "B2")
        # else:
        #     print(f"Logo image not found at {logo_path}")

        # Update specified cells with data
        sheet["F8"] = data.get("trade_date", "N/A")
        sheet["F9"] = data.get("order_number", "N/A")
        sheet["B18"] = data.get("client", "")
        sheet["B19"] = data.get("district", "")
        sheet["B20"] = data.get("region", "Dar es Salaam")
        sheet["B21"] = data.get("tin", "")
        sheet["B22"] = data.get("vrn", "")
        sheet["E25"] = format_number(data.get("consideration", 0))
        sheet["F27"] = format_number(data.get("brokerage", 0))
        sheet["F28"] = format_number(data.get("dse", 0))
        sheet["F29"] = format_number(data.get("cmsa", 0))
        sheet["F30"] = format_number(data.get("cds", 0))
        sheet["F31"] = format_number(data.get("fidelity", 0))
        sheet["F34"] = format_number(data.get("subtotal", 0))
        sheet["F35"] = format_number(data.get("vat", 0))
        sheet["F37"] = format_number(data.get("total_fees", 0))

        # Set alignment to right for the above cells
        for cell in ["E25", "F27", "F28", "F29", "F30", "F31", "F34", "F35", "F37"]:
            sheet[cell].alignment = Alignment(horizontal="right")

        # Save the updated file
        workbook.save(template_path)
        workbook.close()
        logging.info(f"Excel file updated successfully at: {template_path}")
        return template_path
    except Exception as e:
        logging.info(f"Error updating Excel file: {e}")
        raise


def print_excel_file(file_path, printer_name=None):
    try:
        logging.info(f"Printing the Excel file: {file_path}")

        # Select printer
        if not printer_name:
            logging.info("No printer specified. Fetching default printer.")
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            if printers:
                printer_name = printers[0][2]  # Select the first available printer
                logging.info(f"Default printer selected: {printer_name}")
            else:
                raise Exception("No printers available.")
        else:
            logging.info(f"Using specified printer: {printer_name}")

        # Print the Excel file with all formatting using the default application
        win32api.ShellExecute(
            0,
            "printto",
            file_path,
            f'"{printer_name}"',
            ".",
            0
        )

        logging.info(f"Excel file '{file_path}' sent to printer: {printer_name}")
    except Exception as e:
        logging.debug(f"Error printing Excel file: {e}")
        raise


def list_printers():
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    logging.info("Available Printers") 
    for printer in printers:
        logging.info(printer[2])  # Printer name


# Flask setup
app = Flask(__name__)
CORS(app)

@app.route('/run-receipt-script', methods=['POST'])
def run_invoice_script():
    list_printers()
    try:
        logging.info("Received request for invoice processing.")
        data = request.json
        if not data:
            logging.debug("No data received in the request.")
            return jsonify({"error": "No data provided"}), 400
        

        
        logging.info(f"Request data: {data}")
        # Process Excel file
        updated_excel_path = update_excel_with_data(data)

        # Print the Excel file
        print_excel_file(updated_excel_path, data.get("printer_name", None))
        # print_excel_file(updated_excel_path, "Incotex")

        

        # Send email with the Excel file
        # send_email_with_attachment(data, updated_excel_path)


        logging.info("Invoice processing completed successfully.")
        return jsonify({"message": "Invoice processed successfully."})
    except Exception as e:
        logging.error(f"Error processing invoice: {e}")
        return jsonify({"error": str(e)}), 500


        # print("Invoice processing completed successfully.")
        # return jsonify({"message": "Invoice processed successfully."})
        # except Exception as e:
        # print(f"Error processing invoice: {e}")
        # return jsonify({"error": str(e)}), 500

# if __name__ == '__main__':
#     print("Starting Flask server...")
#     app.run(debug=True, host='0.0.0.0', port=7860)

if __name__ == '__main__':
    logging.info("Starting Flask server...")
    app.run(debug=True, host='0.0.0.0', port=7860)
