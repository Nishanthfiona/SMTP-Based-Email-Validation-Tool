# SMTP-Based Email Validation Tool

This is a Python-based application that validates email addresses using SMTP and bounce-back email detection. It processes emails from an Excel file and generates separate files for valid and invalid email addresses. 

## Prerequisites

1. **Gmail Account**: 
   - You must have a Gmail account to use this tool.
   - Ensure that your Gmail account has 2-step verification enabled.

2. **Gmail App Password**:
   - App passwords are 16-character passcodes that allow less secure apps or devices to access your Google account. 
   - Follow these steps to generate an app password:
     1. Go to [Google Account Security Settings](https://myaccount.google.com/security).
     2. Under **"Signing in to Google"**, select **"App Passwords"**.
     3. Log in with your account credentials.
     4. Under **"Select the app and device you want to generate the app password for"**, choose **Mail** as the app and your preferred device.
     5. Copy the 16-character password provided.

3. **Python Libraries**: Install the required libraries:
   ```bash
   pip install pandas openpyxl
   ```

4. **Input Excel File**: Prepare an Excel file containing the email addresses in one column. The column name should be `Email` (or update the script for your column name).

---

## How to Use

1. **Clone or Download the Repository**:
   ```bash
   git clone https://github.com/YourUsername/SMTP-Based-Email-Validation-Tool.git
   cd SMTP-Based-Email-Validation-Tool
   ```

2. **Prepare the Input File**:
   - Save your list of email addresses in an Excel file, e.g., `emails.xlsx`. Ensure that the email addresses are in a single column, with the header name `Email` (case-sensitive).

3. **Run the Script**:
   - Open the script and update the following variables:
     - `gmail_user`: Your Gmail ID.
     - `gmail_app_password`: The app password generated in the prerequisites step.
     - `input_excel`: Path to your Excel file containing the email list.
     - `start_row` and `end_row`: Specify the range of rows to process.
     - `email_column`: Set the column name containing email addresses (default is `Email`).

   - Run the script:
     ```bash
     python validate_emails.py
     ```

4. **Output Files**:
   - Two output Excel files will be generated:
     - `valid_emails_<start_row>_<end_row>.xlsx`: Contains valid email addresses.
     - `invalid_emails_<start_row>_<end_row>.xlsx`: Contains invalid email addresses.

---

## How It Works

1. **Syntax Validation**:
   - The script first checks if the email address matches a valid email format using a regex.

2. **SMTP Test**:
   - It sends a test email to the address using Gmail's SMTP server.

3. **Bounce-Back Detection**:
   - The script monitors Gmail's inbox for bounce-back emails. If a bounce-back is detected for a test email, it is marked as invalid.

4. **Processing Emails**:
   - Emails are processed in batches based on the `start_row` and `end_row` parameters.

---

## Troubleshooting

- **SMTPAuthenticationError**:
  - Ensure that you are using the correct Gmail credentials and app password.
  - Verify that 2-step verification is enabled in your Google account.

- **"No bounce detected"**:
  - Check the test email address for typos or issues. 
  - Increase the `wait_duration` in the `check_bounce_back()` function to allow more time for bounce-back detection.

- **Library Errors**:
  - Ensure all required libraries are installed:
    ```bash
    pip install pandas openpyxl
    ```

---

## Customization

- Modify the email content by changing the subject and body in the `send_test_email` function:
  ```python
  subject = "Test Email"
  body = "This is a test email to validate your address."
  ```

- Update the email column name in the `process_emails()` function if your input file uses a different header.

---

## Disclaimer

- This tool is designed for educational purposes and must not be used to spam or violate the privacy of others.
- Ensure compliance with all applicable email and data privacy regulations.

---
