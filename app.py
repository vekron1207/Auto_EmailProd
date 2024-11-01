from flask import Flask, render_template, request, jsonify, send_file
from email_sender import send_emails_from_csv
import openpyxl
from io import BytesIO
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Global variable to store results
email_results = []

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send-emails', methods=['POST'])
def send_emails():
    global email_results
    if 'csvFile' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['csvFile']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    # Call the email-sending function and store the results
    email_results = send_emails_from_csv(filepath)
    
    return jsonify({"results": email_results})

@app.route('/download-report')
def download_report():
    global email_results
    
    # Create an Excel workbook and add the data dynamically
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Email Report"
    
    # Add headers
    worksheet.append(["First Name", "Email", "Status"])

    # Add email results to the worksheet
    for result in email_results:
        worksheet.append([result["first_name"], result["email"], result["status"]])

    # Save the workbook to a BytesIO stream
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    # Send the file as a downloadable attachment
    return send_file(output, as_attachment=True, download_name="email_report.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    app.run(debug=True)
