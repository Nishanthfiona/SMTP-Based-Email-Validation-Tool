import smtplib
import imaplib
import email
import re
import pandas as pd
from time import time, sleep

# Email regex validation
def is_valid_syntax(email):
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zAHZ0-9-]+\.[a-zA-Z0-9-.]+$'
    return re.match(email_regex, email) is not None

# Check for bounce-back emails
def check_bounce_back(gmail_user, gmail_app_password, test_email, wait_duration=120):
    try:
        # Connect to the IMAP server
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(gmail_user, gmail_app_password)
        mail.select("inbox")

        start_time = time()
        while time() - start_time < wait_duration:
            # Search for unseen messages (waiting for bounce-back)
            status, messages = mail.search(None, 'UNSEEN')
            if status == "OK":
                for num in messages[0].split():
                    status, msg_data = mail.fetch(num, "(RFC822)")
                    if status == "OK":
                        msg = email.message_from_bytes(msg_data[0][1])
                        subject = msg["subject"]
                        body = msg.get_payload(decode=True).decode()

                        # Bounce-check based on subject and content
                        if "bounce" in subject.lower() or "undelivered" in subject.lower():
                            if test_email in body:
                                print(f"Bounce-back detected for {test_email}")
                                return False  # Bounce detected
            sleep(5)  # Wait between checks

        return True  # No bounce detected after waiting
    except Exception as e:
        print(f"Error checking bounce-back: {e}")
        return False

# Send a test email and validate
def send_test_email(test_email, gmail_user, gmail_app_password):
    try:
        # Send a test email
        server = smtplib.SMTP("smtp.gmail.com", 587, timeout=30)
        server.starttls()
        server.login(gmail_user, gmail_app_password)

        # Test email content
        from_email = gmail_user
        to_email = test_email
        subject = "Test Email"
        body = "This is a test email to validate your address."

        message = f"Subject: {subject}\n\n{body}"
        server.sendmail(from_email, to_email, message)
        server.quit()
        return True
    except Exception as e:
        print(f"Error sending test email to {test_email}: {e}")
        return False

# Validate email with SMTP and bounce-back handling
def validate_email(test_email, gmail_user, gmail_app_password):
    if not is_valid_syntax(test_email):
        return False  # Invalid syntax

    print(f"Sending test email to {test_email}...")
    if send_test_email(test_email, gmail_user, gmail_app_password):
        print(f"Initial success for {test_email}. Waiting for final validation.")
        sleep(10)  # Wait for bounce-back to arrive
        if not check_bounce_back(gmail_user, gmail_app_password, test_email):
            print(f"{test_email}: Invalid (Bounce detected)")
            return False
        print(f"{test_email}: Valid and received")
        return True
    else:
        return False

# Main function to handle Excel input/output and timing
def process_emails(input_excel, output_excel, gmail_user, gmail_app_password, start_row, end_row, email_column='Email'):
    # Read input Excel file
    df = pd.read_excel(input_excel)

    # Slice the DataFrame based on the start and end row
    df_subset = df.iloc[start_row-1:end_row]

    # Create lists to store valid and invalid emails
    valid_emails = []
    invalid_emails = []

    start_time = time()  # Start timing the processing

    for idx, row in df_subset.iterrows():
        email_address = row[email_column]  # Use the specified email column
        is_valid = validate_email(email_address, gmail_user, gmail_app_password)
        if is_valid:
            valid_emails.append(row)
        else:
            invalid_emails.append(row)

    # Calculate the time taken
    end_time = time()
    processing_time = end_time - start_time
    print(f"Total processing time: {processing_time:.2f} seconds")

    # Create DataFrame for results and write to Excel
    valid_df = pd.DataFrame(valid_emails)
    invalid_df = pd.DataFrame(invalid_emails)

    # Construct filenames with start and end row information
    valid_output_filename = f"valid_emails_{start_row}_{end_row}.xlsx"
    invalid_output_filename = f"invalid_emails_{start_row}_{end_row}.xlsx"

    # Save the results to Excel files with start and end rows in filenames
    with pd.ExcelWriter(valid_output_filename, engine='openpyxl') as writer:
        valid_df.to_excel(writer, sheet_name='Valid Emails', index=False)

    with pd.ExcelWriter(invalid_output_filename, engine='openpyxl') as writer:
        invalid_df.to_excel(writer, sheet_name='Invalid Emails', index=False)

# Main entry point
if __name__ == "__main__":
    gmail_user = "senthilkumargwgk@gmail.com"
    gmail_app_password = "rlkt fudf juoq cbsh"
    input_excel = "X.xlsx"  # Path to the input Excel file with emails in a column

    # Define the start and end rows (for example: 1 to 10)
    start_row = 1  # Change this to your desired start row
    end_row = 10  # Change this to your desired end row
    email_column = 'Email'  # Change this to your actual email column name if different
    
    process_emails(input_excel, "", gmail_user, gmail_app_password, start_row, end_row, email_column)