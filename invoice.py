import os
import win32print
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
    if os.name == 'nt':  # For Windows
        return os.path.join(os.environ['USERPROFILE'], 'Downloads')
    else:
        return str(Path.home() / 'Downloads')


# Function to process Excel template
def update_excel_with_data(data):
    try:
        template_path = "invoice.xlsx"  # Path to the Excel template
        if not os.path.exists(template_path):
            raise FileNotFoundError("Invoice template not found.")
        
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

        return updated_excel_path
    except Exception as e:
        raise Exception(f"Error updating Excel file: {e}")


# Function to print the updated Excel file
def print_excel_file(file_path, printer_name=None):
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError("Excel file not found for printing.")

        if not printer_name:
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            if printers:
                printer_name = printers[0][2]  # Select the first available printer
            else:
                raise Exception("No printers available.")
        
        # Use the printer to print the file (command for demonstration; adapt for specific needs)
        os.system(f"start /min excel.exe /p /m {file_path}")
    except Exception as e:
        raise Exception(f"Error printing Excel file: {e}")


# Function to send an email with the Excel file
def send_email_with_attachment(data, attachment_path):
    try:
        sender_email = "your_email@example.com"
        receiver_email = data.get("email", "")
        if not receiver_email:
            raise Exception("No recipient email provided.")

        subject = "Your Updated Invoice"
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = subject

        body = "Please find your updated invoice attached."
        message.attach(MIMEText(body, 'plain'))

        with open(attachment_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(attachment_path)}")
            message.attach(part)

        with smtplib.SMTP('smtp.example.com', 587) as server:
            server.starttls()
            server.login(sender_email, "your_password")
            server.sendmail(sender_email, receiver_email, message.as_string())
    except Exception as e:
        raise Exception(f"Error sending email: {e}")


# Flask setup
app = Flask(__name__)
CORS(app)


@app.route('/run-receipt-script', methods=['POST'])
def run_invoice_script():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No data provided"}), 400

        # Process Excel file
        updated_excel_path = update_excel_with_data(data)

        # Print the Excel file
        print_excel_file(updated_excel_path, data.get("printer_name", None))

        # Send email with the Excel file
        # send_email_with_attachment(data, updated_excel_path)

        return jsonify({"message": "Invoice processed successfully."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=7860)
