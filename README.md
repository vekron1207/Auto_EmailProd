# Email Sending Script

This script allows you to send emails to multiple recipients using details provided in a CSV file. It's beginner-friendly, and the setup process is straightforward.

---

## Features
- Send emails to a list of recipients from a CSV file.
- Generates an Excel report of email delivery statuses.
- Easy to configure and run.

---

## Prerequisites

### 1. **Install Python**
   - Download and install Python 3.x from the [official website](https://www.python.org/downloads/). 
   - During installation, ensure you check **"Add Python to PATH"**.

### 2. **Install VS Code** (Optional, but recommended)
   - Download and install [Visual Studio Code](https://code.visualstudio.com/).

---

## Setup Instructions

### 1. Clone or Download the Repository
   - Clone the repository or download all files (e.g., `app.py`, `emailsender.py`, etc.) to a folder on your local machine.

### 2. Install Dependencies
   - Open a terminal or command prompt and navigate to the folder containing the downloaded files.
   - Install required Python libraries by running the following command:
     ```
     pip install flask openpyxl
     ```

### 3. Configure the `config.json` File
   - Open the `config.json` file and replace the placeholder values with your own:
     ```json
     {
       "SMTP_SERVER": "smtp.example.com",
       "SMTP_PORT": 465,
       "SMTP_USERNAME": "your_smtp_username",
       "SMTP_PASSWORD": "your_smtp_password",
       "SENDER_EMAIL": "your_email@example.com",
       "EMAIL_SUBJECT": "Your Subject Here",
       "EMAIL_BODY": "Your Email Body Here"
     }
     ```
   - Save the file.

### 4. Prepare the CSV File
   - Create a CSV file (e.g., `emails.csv`) containing a list of email addresses. For example:
     ```
     email
     user1@example.com
     user2@example.com
     ```
   - Place this file in the same folder as the script.

---

## Running the Script

### 1. Start the Application
   - Run the following command to start the Flask app:
     ```
     python app.py
     ```
   - Open your browser and go to `http://127.0.0.1:5000`.

### 2. Send Emails
   - Upload your `emails.csv` file via the web interface.
   - The app will send emails to the listed recipients and display the results.

### 3. Download the Report
   - Once the emails are sent, click the **Download Report** button on the web interface to download an Excel file with the email delivery status.

---

## Notes
- Ensure your SMTP server credentials are correct. If using Gmail, you may need to enable "App Passwords" in your Google account settings.
- If any issues arise, check the terminal for error messages.

---

## Contributing
If you find any issues or have suggestions for improvements, feel free to open an issue or submit a pull request. Or reach to me on varun.kashyap1207@gmail.com

---

## License
This project is licensed under the [MIT License](LICENSE).
