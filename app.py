import streamlit as st
import pandas as pd
import smtplib
import imaplib
import email
import re
from io import BytesIO
from time import time, sleep

# Email regex validation
def is_valid_syntax(email):
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
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
                                return False  # Bounce detected
            sleep(5)  # Wait between checks

        return True  # No bounce detected after waiting
    except Exception as e:
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
    except Exception:
        return False

# Validate email with SMTP and bounce-back handling
def validate_email(test_email, gmail_user, gmail_app_password):
    if not is_valid_syntax(test_email):
        return False  # Invalid syntax

    if send_test_email(test_email, gmail_user, gmail_app_password):
        sleep(10)  # Wait for bounce-back to arrive
        if not check_bounce_back(gmail_user, gmail_app_password, test_email):
            return False
        return True
    else:
        return False

# Process emails
def process_emails(df, gmail_user, gmail_app_password, start_row, end_row, email_column):
    df_subset = df.iloc[start_row-1:end_row]

    valid_emails = []
    invalid_emails = []

    for idx, row in df_subset.iterrows():
        email_address = row[email_column]
        is_valid = validate_email(email_address, gmail_user, gmail_app_password)
        if is_valid:
            valid_emails.append(row)
        else:
            invalid_emails.append(row)

    valid_df = pd.DataFrame(valid_emails)
    invalid_df = pd.DataFrame(invalid_emails)

    return valid_df, invalid_df

# Streamlit app
def main():
    st.title("Email Validation App")
    st.write("Validate email addresses by sending test emails and checking for bounce-backs.")

    gmail_user = st.text_input("Enter Gmail Address")
    gmail_app_password = st.text_input("Enter Gmail App Password", type="password")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    start_row = st.number_input("Start Row", min_value=1, value=1)
    end_row = st.number_input("End Row", min_value=1, value=10)
    email_column = st.text_input("Email Column Name", value="Email")

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.write("Uploaded Data Preview:", df.head())

        if st.button("Validate Emails"):
            if not gmail_user or not gmail_app_password:
                st.error("Please enter Gmail credentials.")
                return

            try:
                st.write("Processing emails... This might take some time.")
                valid_df, invalid_df = process_emails(df, gmail_user, gmail_app_password, start_row, end_row, email_column)

                # Convert results to downloadable files
                valid_file = BytesIO()
                invalid_file = BytesIO()

                with pd.ExcelWriter(valid_file, engine='openpyxl') as writer:
                    valid_df.to_excel(writer, sheet_name='Valid Emails', index=False)

                with pd.ExcelWriter(invalid_file, engine='openpyxl') as writer:
                    invalid_df.to_excel(writer, sheet_name='Invalid Emails', index=False)

                st.success("Processing complete!")

                # Show download buttons
                valid_filename = f"valid_emails_{start_row}_{end_row}.xlsx"
                invalid_filename = f"invalid_emails_{start_row}_{end_row}.xlsx"
                
                # Download buttons for both files
                st.download_button("Download Valid Emails", valid_file.getvalue(), valid_filename)
                st.download_button("Download Invalid Emails", invalid_file.getvalue(), invalid_filename)

            except Exception as e:
                st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
