import os
import openpyxl
import csv
import smtplib
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time

config_file_path = "G:/Work/autoEmail/config.json"
with open(config_file_path) as config_file:
    config = json.load(config_file)

SMTP_SERVER = config["SMTP_SERVER"]
SMTP_PORT = config["SMTP_PORT"]
SMTP_USERNAME = config["SMTP_USERNAME"]
SMTP_PASSWORD = config["SMTP_PASSWORD"]
SENDER_EMAIL = config["SENDER_EMAIL"]
EMAIL_SUBJECT = config["EMAIL_SUBJECT"]
EMAIL_TEMPLATE = """REMAINDER!!!!

Dear user,

This mail is regarding the credentials of YOJANA - MIS developed by PMD, NCERT.
Yojana web application link - https://yojana.ncert.gov.in

Username: {uname}
Password: Your registered NIC mail id is your first-time login password. You need to change it after login and keep it safe.

For more details, please refer to the user manual on https://yojana.ncert.gov.in/User-Manual.html

Thanks & regards,
Yojana Administrator
PMD, NCERT
"""

# Define a list to store the report data
report_data = []


# Function to append data to the report list
def append_to_report(first_name, last_name, email, status):
    report_data.append([first_name, last_name, email, status])


def send_email(recipient_email, first_name, uname):
    message = MIMEMultipart()
    message["From"] = SENDER_EMAIL
    message["To"] = recipient_email
    message["Subject"] = EMAIL_SUBJECT

    email_body = EMAIL_TEMPLATE.format(first_name=first_name, uname=uname)
    message.attach(MIMEText(email_body, "plain"))

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.login(SMTP_USERNAME, SMTP_PASSWORD)
        server.send_message(message)


def send_emails_from_csv(csv_file):
    with open(csv_file, "r") as file:
        reader = csv.DictReader(file)
        for row in reader:
            uname = row["uname"]
            first_name = row["first_name"]
            last_name = row["last_name"]
            email = row["email"]
            recipient_email = email.strip() if email else ""
            # Within the send_emails_from_csv function
            if recipient_email:
                send_email(recipient_email, first_name, uname)
                append_to_report(first_name, last_name, email, "Sent")
                print(f"Email sent to {recipient_email}.")
                time.sleep(5)
            else:
                append_to_report(first_name, last_name, email, "Skipped")
                print(
                    f"Skipping row: Email address missing for {first_name} + {last_name}"
                )


csv_file = "G:/Work/autoEmail/dummy.csv"

send_emails_from_csv(csv_file)


# Report file path
report_file = "G:/Work/autoEmail/email_report.xlsx"

# Check if the report file exists
if os.path.exists(report_file):
    # Load the existing Excel file
    workbook = openpyxl.load_workbook(report_file)
    # Select the default worksheet (usually named 'Sheet')
    worksheet = workbook.active
else:
    # Create a new Excel workbook and add a worksheet
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    # Add header row
    worksheet.append(["First Name", "Last Name", "Email", "Status"])

# Write data rows
for data_row in report_data:
    worksheet.append(data_row)

# Save the Excel file
workbook.save(report_file)

print(f"Email report updated and saved to {report_file}.")
