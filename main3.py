import pandas as pd
import cv2
import numpy as np
from PIL import ImageFont, ImageDraw, Image
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage

def clean_folder(folder_name):
    print("Cleaning {} folder...".format(folder_name))
    for certificates in os.listdir("result/{}/".format(folder_name)):
        os.remove("result/{}/{}".format(folder_name, certificates))
    print("Clean-up completed.")

def load_excel_data(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except FileNotFoundError:
        print("Excel file not found.")
        return None

def generate_certificate(template_path, data_frame):
    for index, row in data_frame.iterrows():
        template = cv2.imread(template_path)

        # Modify the template with data from the Excel file
        # For example, change text position and insert relevant data onto the certificate
        cv2.putText(template, f"{row['Student Name']}", (1550, 1300),
                    cv2.FONT_HERSHEY_DUPLEX, 7, (0, 0, 0), 10)
        cv2.putText(template, f"{row['Roll #']}", (1800, 1500),
                    cv2.FONT_HERSHEY_SCRIPT_COMPLEX, 4, (0, 0, 0), 10)
        cv2.putText(template, f"{row['CertificateID']}", (150, 150),
                    cv2.FONT_HERSHEY_DUPLEX, 4, (0, 0, 0), 10)
        cv2.putText(template, f"{row['Branch']}", (2700, 1500),
                    cv2.FONT_HERSHEY_SCRIPT_COMPLEX, 3, (0, 0, 0), 10)

        # Save the modified template as a certificate
        certificate_name = f"{row['Roll #']}-{row['CertificateID']}.jpg"
        cv2.imwrite(f"result/certificates/{certificate_name}", template)

        # Send the generated certificate via email
        send_certificate_email(row['E-Mail ID'], certificate_name)

def send_certificate_email(receiver_email, certificate_name):
    # SMTP configuration to send emails
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    sender_email = 'nadeemahamed916@gmail.com'
    password = 'bftl luvh obwp lmgy'

    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = 'Certificate of TECH NEHATHON 2K23'

    # Attach the certificate in the email
    with open(f"result/certificates/{certificate_name}", 'rb') as file:
        attachment = MIMEImage(file.read())
        attachment.add_header('Content-Disposition', 'attachment', filename=certificate_name)
        message.attach(attachment)

    # Establishing a connection to the SMTP server and sending the email
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())

def main():
    clean_folder("certificates")  # Clean-up old certificates

    # Load Excel data
    file_path = 'Aman.xlsx'
    data = load_excel_data(file_path)

    if data is not None:
        # Generate certificates and send emails
        certificate_template_path = 'certificate.jpg'
        generate_certificate(certificate_template_path, data)

if __name__ == "__main__":
    main()
