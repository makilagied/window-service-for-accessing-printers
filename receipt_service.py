import os
import win32print
import win32ui
import win32serviceutil
import win32service
import win32event
import socket
from fpdf import FPDF
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from flask import Flask, request, jsonify
import threading
from pathlib import Path


# Function to get the Downloads directory path
def get_downloads_folder():
    if os.name == 'nt':  # For Windows
        # Using the environment variable to get the user's Downloads directory
        downloads_folder = os.path.join(os.environ['USERPROFILE'], 'Downloads')
    else:
        # For Unix-based systems (Linux/macOS), use the XDG user directories
        downloads_folder = str(Path.home() / 'Downloads')
    return downloads_folder

# Function to print receipt, generate PDF, and send email
def PrintReceipt(data):
    try:
        # Create a Text File for printing
        downloads_folder = get_downloads_folder()  # Get the Downloads folder path
        
        receipt_text_file = os.path.join(downloads_folder, "receipt.txt")
        with open(receipt_text_file, "w") as file:
            file.write("Receipt Details:\n")
            for key, value in data.items():
                file.write(f"{key}: {value}\n")
                
        print(f"Receipt saved to: {receipt_text_file}")  # Print path for confirmation


        # Log the data being sent to the printer
        print("Data being sent to printer:")
        for key, value in data.items():
            print(f"{key}: {value}")

        # Retrieve Available Printers
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        printer_name = data.get('printer_name', None)
        
        if not printer_name:
            if printers:
                printer_name = printers[0][2]  # Select the first available printer if none specified
            else:
                return jsonify({"error": "No printers available."}), 404

        # Print the text file using the selected printer
        if os.name == 'nt':  # Windows
            # Initialize the printer
            printer = win32print.OpenPrinter(printer_name)
            print_info = win32print.GetPrinter(printer, 2)
            # Create a printer device context
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)
            hdc.StartDoc("Receipt Print")
            hdc.StartPage()
            hdc.TextOut(100, 100, "Receipt Details")
            y_offset = 120
            for key, value in data.items():
                hdc.TextOut(100, y_offset, f"{key}: {value}")
                y_offset += 20
            hdc.EndPage()
            hdc.EndDoc()
            hdc.DeleteDC()
        else:  # Linux/macOS
            os.system(f"lp {receipt_text_file}")

        # Generate PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="Receipt Details", ln=True, align='C')
        for key, value in data.items():
            pdf.cell(200, 10, txt=f"{key}: {value}", ln=True)
        pdf_output_file = "receipt.pdf"
        pdf.output(pdf_output_file)

        # Send Email with PDF
        sender_email = "your_email@example.com"
        receiver_email = data.get("email", "")
        if receiver_email:
            subject = "Your Receipt"
            message = MIMEMultipart()
            message['From'] = sender_email
            message['To'] = receiver_email
            message['Subject'] = subject

            body = "Please find your receipt attached."
            message.attach(MIMEText(body, 'plain'))

            # Attach PDF
            with open(pdf_output_file, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f"attachment; filename={pdf_output_file}")
                message.attach(part)

            # Send email
            try:
                with smtplib.SMTP('smtp.example.com', 587) as server:
                    server.starttls()
                    server.login(sender_email, "your_password")
                    server.sendmail(sender_email, receiver_email, message.as_string())
            except Exception as e:
                print(f"Failed to send email: {e}")

        # Clean up temporary files
        os.remove(receipt_text_file)
        os.remove(pdf_output_file)

    except Exception as e:
        print(f"Error in PrintReceipt: {e}")
        raise


# Flask Setup
app = Flask(__name__)

@app.route('/')
def root():
    return "Service is up and running"

# Define the endpoint for the Flask app
@app.route('/run-receipt-script', methods=['POST'])
def run_receipt_script():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No data provided"}), 400
        
        # Process the receipt (print, PDF generation, and email)
        PrintReceipt(data)
        
        return jsonify({"message": "Receipt processed successfully"})
    
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# Create the Service class to keep the Flask service running
# class FlaskService(win32serviceutil.ServiceFramework):
#     _svc_name_ = "iTrustEfdService"
#     _svc_display_name_ = "iTrust Service for Printing receipts"
#     _svc_description_ = "Service to run Flask web application for receipt printing"

#     def __init__(self, args):
#         win32serviceutil.ServiceFramework.__init__(self, args)
#         self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)

#     def SvcStop(self):
#         self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
#         win32event.SetEvent(self.hWaitStop)

#     def SvcDoRun(self):
#         self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        
#         # Start Flask in a separate thread to run continuously
#         threading.Thread(target=self.run_flask).start()

#         # Wait for the stop event
#         win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)

#     def run_flask(self):
#         app.run(debug=False, host='0.0.0.0', port=7860)

# if __name__ == '__main__':
#     win32serviceutil.HandleCommandLine(FlaskService)


# Main block for testing as a standalone app
if __name__ == '__main__':
    # Run Flask app for testing
    print("Running as a standalone app...")
    app.run(debug=True, host='0.0.0.0', port=7860)