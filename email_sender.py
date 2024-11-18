import os
import smtplib
import json
import time
import openpyxl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Load configuration
with open("config.json") as config_file:
    config = json.load(config_file)

SMTP_SERVER = config["SMTP_SERVER"]
SMTP_PORT = config["SMTP_PORT"]
SMTP_USERNAME = config["SMTP_USERNAME"]
SMTP_PASSWORD = config["SMTP_PASSWORD"]
SENDER_EMAIL = config["SENDER_EMAIL"]
EMAIL_SUBJECT = config["EMAIL_SUBJECT"]
EMAIL_BODY = config["EMAIL_BODY"]  # Fixed email body from config

def send_email(recipient_email):
    """
    Send an email to the specified recipient.
    """
    message = MIMEMultipart()
    message["From"] = SENDER_EMAIL
    message["To"] = recipient_email
    message["Subject"] = EMAIL_SUBJECT

    # Attach the fixed email body
    message.attach(MIMEText(EMAIL_BODY, "plain"))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.send_message(message)
        return {"email": recipient_email, "status": "Email sent"}
    except Exception as e:
        return {"email": recipient_email, "status": f"Failed to send: {str(e)}"}

def save_results_to_excel(results, report_file="email_report.xlsx"):
    """
    Save the results of the email operation to an Excel file.
    """
    if os.path.exists(report_file):
        workbook = openpyxl.load_workbook(report_file)
        worksheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        # Add header row
        worksheet.append(["Email", "Status"])

    # Write each result as a new row
    for result in results:
        worksheet.append([result["email"], result["status"]])

    workbook.save(report_file)

def send_emails_from_csv(csv_file):
    """
    Read the list of emails from the first line of a CSV file and send emails to all addresses.
    """
    results = []
    with open(csv_file, "r") as file:
        # Read the first line and split emails by comma
        line = file.readline().strip()
        email_list = line.split(",")

        for email in email_list:
            recipient_email = email.strip()
            if recipient_email:
                result = send_email(recipient_email)
                results.append(result)
                print(result["status"])
                time.sleep(5)  # Add delay to avoid rate-limiting issues
            else:
                results.append({"email": "", "status": "Skipped: Missing email"})
                print("Skipping: Empty email")

    # Save results to an Excel file
    save_results_to_excel(results)
    return results

# Example usage
# send_emails_from_csv("emails.csv")
