"""
README: SMTP-Based Email Validation Tool

This script performs a comprehensive validation of email addresses by leveraging Python's `smtplib`, `re`, 
and `dns.resolver` libraries, along with Gmail's SMTP server for authentication and recipient verification.

**Functionality Overview:**
1. **Syntax Validation:** Checks if the email follows proper formatting using regex.
2. **MX Record Verification:** Confirms that the domain of the email has valid mail exchange (MX) records.
3. **SMTP Validation:** Authenticates with Gmail's SMTP server and performs a recipient address check.

**Input:** 
- An Excel file containing email addresses in a specific column.

**Output:**
- Two Excel files:
  1. **Valid Emails File:** Contains all rows with emails validated as "deliverable" along with validation time.
  2. **Invalid Emails File:** Contains all rows with undeliverable emails along with validation time.

**How It Works:**
1. Reads the input Excel file to fetch the list of emails.
2. Processes each email to perform the three validation steps.
3. Separates the emails into valid and invalid categories.
4. Writes these categories into two separate Excel files for easy access.

**Prerequisites:**
1. Python 3.7 or above.
2. Required libraries: pandas, openpyxl, dnspython, smtplib.
3. A Gmail account with an App Password set up for SMTP access.

**Steps to Run the Script:**
1. Replace the following variables with your information:
    - `gmail_user`: Your Gmail address.
    - `gmail_app_password`: Your Gmail App Password.
    - `input_file`: Path to the Excel file with emails.
    - `email_column`: Name of the column containing email addresses.

2. Adjust `start_row` and `end_row` if processing a specific range.

3. Run the script, and it will generate two output files in the same directory as the input file:
    - **Valid Emails File:** `<input_file_name> valid emails <start_row>-<end_row>.xlsx`
    - **Invalid Emails File:** `<input_file_name> invalid emails <start_row>-<end_row>.xlsx`

4. Each output file includes:
    - Original data from the input file.
    - A "Validation Status" column indicating "Valid" or "Invalid".
    - A "Validation Time (s)" column showing the time taken to validate each email.

**Example:**
If your input file is `email_list.xlsx` with 100 rows and you process rows 1 to 50:
- Valid emails are saved in `email_list valid emails 1-50.xlsx`.
- Invalid emails are saved in `email_list invalid emails 1-50.xlsx`.
