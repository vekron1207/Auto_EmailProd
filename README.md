# Email Script

This Python script allows you to send emails to recipients listed in a CSV file. It reads the email addresses, usernames, and other details from the CSV file and sends customized emails to each recipient.

## Prerequisites

- Python 3.x installed on your system
- CSV file containing the recipient details (e.g., `emails.csv`)
- SMTP server credentials or an SMTP service account for sending emails

## Installation

1. Clone the repository or download the script file (`script.py`) to your local machine.

2. Install the required dependencies by running the following command:
   
```
pip install smtplib
```

## Configuration

1. Open the `script.py` file in a text editor.

2. Update the SMTP server configuration:
- Set the `SMTP_SERVER` variable to your SMTP server address.
- Set the `SMTP_PORT` variable to the appropriate port number for your SMTP server.
- Set the `SMTP_USERNAME` and `SMTP_PASSWORD` variables to your SMTP server credentials or an SMTP service account.

3. Customize the email template:
- Modify the `EMAIL_TEMPLATE` variable to match your desired email content.
- You can use placeholders like `{uname}` and `{first_name}` to dynamically replace the recipient's username and first name in the email template.

## Usage

1. Prepare your CSV file:
- Create a CSV file with the following columns: `uname`, `first_name`, `last_name`, and `email`.
- Enter the recipient details in each row.

2. Run the script:
- Open a terminal or command prompt and navigate to the directory containing the `script.py` file.
- Execute the following command:
  ```
  python script.py
  ```
- The script will read the CSV file, send emails to the recipients, and display a confirmation message for each email sent.

## Contributing

Contributions are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request on the repository.

## License

This project is licensed under the [MIT License](LICENSE).
