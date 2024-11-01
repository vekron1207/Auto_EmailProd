import os
import openpyxl
import csv
import smtplib
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time

# Load configuration
with open("config.json") as config_file:
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

def send_email(recipient_email, first_name, uname):
    message = MIMEMultipart()
    message["From"] = SENDER_EMAIL
    message["To"] = recipient_email
    message["Subject"] = EMAIL_SUBJECT

    email_body = EMAIL_TEMPLATE.format(first_name=first_name, uname=uname)
    message.attach(MIMEText(email_body, "plain"))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.send_message(message)
        return {"first_name": first_name, "email": recipient_email, "status": "Email sent"}
    except Exception as e:
        return {"first_name": first_name, "email": recipient_email, "status": f"Failed to send: {str(e)}"}

def save_results_to_excel(results, report_file="email_report.xlsx"):
    if os.path.exists(report_file):
        workbook = openpyxl.load_workbook(report_file)
        worksheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        # Add header row
        worksheet.append(["First Name", "Email", "Status"])

    # Write each result as a new row
    for result in results:
        worksheet.append([result["first_name"], result["email"], result["status"]])

    workbook.save(report_file)

def send_emails_from_csv(csv_file):
    results = []
    with open(csv_file, "r") as file:
        reader = csv.DictReader(file)
        for row in reader:
            uname = row["uname"]
            first_name = row["first_name"]
            email = row["email"]
            recipient_email = email.strip() if email else ""

            if recipient_email:
                result = send_email(recipient_email, first_name, uname)
                results.append(result)
                print(result["status"])
                time.sleep(5)
            else:
                results.append({"first_name": first_name, "email": "", "status": "Skipped: Missing email"})
                print(f"Skipping row: Email address missing for {first_name}")

    save_results_to_excel(results)
    return results
