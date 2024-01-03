from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import io
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from os.path import basename



def send_email_with_attachment(subject, message, from_email, to_email, password, file_path):
    # Create a MIMEMultipart object
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # Attach the message to the MIMEMultipart object
    msg.attach(MIMEText(message, 'plain'))

    # Attach the file
    with open(file_path, "rb") as file:
        part = MIMEApplication(
            file.read(),
            Name=basename(file_path)
        )
    # After the file is closed
    part['Content-Disposition'] = f'attachment; filename="{basename(file_path)}"'
    msg.attach(part)

    try:
        # Set up the SMTP server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()  # Enable security
        server.login(from_email, password)  # Login to the email server

        # Send the email
        server.send_message(msg)
        print("Email sent successfully!")

    except Exception as e:
        print(f"Failed to send email: {e}")

    finally:
        server.quit()

def get_pdf_page_size(pdf_path):
    # Open the PDF file
    pdf = PdfReader(pdf_path)

    # Get the first page
    first_page = pdf.pages[0]

    # Get the page size (width and height)
    page_width = first_page.mediabox[2]
    page_height = first_page.mediabox[3]

    return page_width, page_height
    
def string_width(string, font, font_size):
    # Create a temporary canvas to measure text
    temp_canvas = canvas.Canvas(io.BytesIO())
    temp_canvas.setFont(font, font_size)
    return temp_canvas.stringWidth(string)

def generate_certificates(dataset_path, template_path, output_folder, email, password):
    data = pd.read_csv(dataset_path)

    font = "Helvetica"
    font_size = 30

    width, height = get_pdf_page_size(template_path)
    width = int(width)
    heigh = int(height)

    # Process each name and email
    for index, row in data.iterrows():
        name = row['Name'].strip()
        recipient_email = row['Email'].strip()

        # Create a new PDF with ReportLab
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont(font, font_size)

        # Calculate text width and position
        text_width = (string_width(name, font, font_size))
        x = (width - text_width) / 2
        y = ((height - font_size) / 2) + 40

        # Add the name text at the desired position
        can.drawString(int(x), int(y), name)
        can.save()

        # Move to the beginning of the StringIO buffer
        packet.seek(0)
        new_pdf = PdfReader(packet)

        # Read the existing PDF template
        existing_pdf = PdfReader(open(template_path, "rb"))
        output = PdfWriter()

        # Add the "watermark" (which is the new_pdf) on the existing page
        page = existing_pdf.pages[0]
        page.merge_page(new_pdf.pages[0])
        output.add_page(page)

        # Write the modified content to a new PDF
        output_path = f"{output_folder}/{name}_certificate.pdf"
        with open(output_path, "wb") as outputStream:
            output.write(outputStream)

        print(f"Certificate saved for: {name}")

        # Send the email with the certificate
        subject = "Your Certificate"
        message = f"Dear {name},\n\nPlease find attached your certificate."
        send_email_with_attachment(subject, message, email, recipient_email, password, output_path)
        print(f"Email sent to: {recipient_email}")


# Example usage
email = "pscifebvc@gmail.com"
password = ""
generate_certificates('E-Certificate-Generator\Data.csv', 'E-Certificate-Generator\Template.pdf', "E-Certificate-Generator\Certificates", email, password)