# Email Validation Tool using SMTP

## Overview
This script validates email addresses by performing the following steps:
1. **Syntax Validation**: Ensures the email address is formatted correctly using a regular expression.
2. **MX Record Validation**: Checks if the email's domain has a valid mail server (MX record).
3. **SMTP Validation**: Uses Gmail's SMTP server to check whether the email address exists and is deliverable.

The tool processes a given Excel file containing email addresses and generates **two output files**:
- **Valid Emails File**: Contains rows with deliverable email addresses and validation details.
- **Invalid Emails File**: Contains rows with undeliverable email addresses and validation details.

---

## Features
- **Input File**: An Excel file with a column containing email addresses.
- **Output Files**:
  1. A file listing valid emails with the validation status and time taken for each.
  2. A file listing invalid emails with the same additional details.
- **Customizable Range**: Allows validation of specific rows from the Excel file.
- **Detailed Logs**: Prints status and time taken for each email during validation.

---

## Prerequisites
1. **Python Libraries**:
   - `pandas`: For Excel file handling.
   - `openpyxl`: For writing to Excel files.
   - `dns.resolver`: For checking domain MX records.
   - `smtplib`: For SMTP-based email validation.
2. **Gmail Account**:
   - A Gmail account with an [App Password](https://support.google.com/accounts/answer/185833) enabled.

---

## How to Use
1. Update the following variables in the script:
   - `gmail_user`: Your Gmail email address.
   - `gmail_app_password`: Your Gmail App Password.
   - `input_file`: The path to the input Excel file containing emails.
   - `email_column`: Name of the column in the Excel file that contains email addresses.

2. Adjust `start_row` and `end_row` to define the range of rows to process (if needed).

3. Run the script, and the output files will be generated in the same directory as the input file:
   - `<input_file_name> valid emails <start_row>-<end_row>.xlsx`
   - `<input_file_name> invalid emails <start_row>-<end_row>.xlsx`

---

## Example
### Input File
**email_list.xlsx**
| ID  | Email              |
|------|--------------------|
| 1    | valid@example.com |
| 2    | invalid@fake.com  |

### Running the Script
```python
python validate_emails.py
