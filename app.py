import smtplib
import imaplib
import email
import re
import pandas as pd
import streamlit as st
from time import time, sleep
from io import BytesIO
from datetime import datetime
import os

# File to store user count
USER_COUNT_FILE = "user_count.txt"

# Function to get the current user count
def get_user_count():
    if os.path.exists(USER_COUNT_FILE):
        with open(USER_COUNT_FILE, "r") as file:
            content = file.read().strip()  # Strip any extra spaces or newlines
            # Default to 0 if the file is empty or contains an invalid value
            count = int(content) if content else 0
    else:
        # Initialize user count to 0 if the file does not exist
        count = 0
        with open(USER_COUNT_FILE, "w") as file:
            file.write(str(count))
    return count


# Function to increment the user count
def increment_user_count():
    count = get_user_count() + 1
    with open(USER_COUNT_FILE, "w") as file:
        file.write(str(count))
    return count



# Email regex validation
def is_valid_syntax(email):
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zAHZ0-9-]+\.[a-zA-Z0-9-.]+$'
    return re.match(email_regex, email) is not None

# Check for bounce-back emails
def check_bounce_back(gmail_user, gmail_app_password, test_email, wait_duration=20):
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
    except Exception as e:
        return False

# Validate email with SMTP and bounce-back handling
def validate_email(test_email, gmail_user, gmail_app_password):
    if not is_valid_syntax(test_email):
        return False, 0  # Invalid syntax, return 0 for validation time

    start_time = time()
    if send_test_email(test_email, gmail_user, gmail_app_password):
        sleep(10)  # Wait for bounce-back to arrive
        if not check_bounce_back(gmail_user, gmail_app_password, test_email):
            end_time = time()
            validation_time = end_time - start_time
            return False, validation_time
        end_time = time()
        validation_time = end_time - start_time
        return True, validation_time
    else:
        end_time = time()
        validation_time = end_time - start_time
        return False, validation_time

# Function to convert DataFrame to Excel
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    output.seek(0)
    return output

# Generate dynamic filename based on start and end row numbers
def generate_filename(input_filename, start_row, end_row, file_type):
    date_str = datetime.now().strftime('%Y-%m-%dT%H-%M-%S')
    return f"{input_filename}_{file_type}_emails_{start_row}_{end_row}.xlsx"

# Main function to handle Excel input/output and timing
def process_emails(input_excel, gmail_user, gmail_app_password, start_row, end_row, email_column='Email'):
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
        is_valid, validation_time = validate_email(email_address, gmail_user, gmail_app_password)
        row_data = row.copy()  # Create a copy of the row to append
        row_data['Validation Status'] = 'Valid' if is_valid else 'Invalid'
        row_data['Validation Time (s)'] = round(validation_time, 2)

        if is_valid:
            valid_emails.append(row_data)
        else:
            invalid_emails.append(row_data)

    # Save results to session state for downloading later
    st.session_state['valid_emails'] = pd.DataFrame(valid_emails)
    st.session_state['invalid_emails'] = pd.DataFrame(invalid_emails)

    # Calculate the time taken
    end_time = time()
    processing_time = end_time - start_time

    # Ensure session state variables are initialized
if 'valid_emails' not in st.session_state:
    st.session_state['valid_emails'] = pd.DataFrame()

if 'invalid_emails' not in st.session_state:
    st.session_state['invalid_emails'] = pd.DataFrame()

# Sidebar content
st.sidebar.title("About Me")
st.sidebar.write("**Nishanth Fiona**")
st.sidebar.write("📊 Data Analyst with 2+ years of experience, learning Data Science.")
st.sidebar.write("📍 Chennai, Tamil Nadu")

# Separator line
st.sidebar.markdown("---")  # This adds a horizontal line as a separator

# Increment user count and display in sidebar directly (no need for __name__ check)
user_count = increment_user_count()  # Increment user count
st.sidebar.metric(label="User Count", value=user_count)  # Display user count in sidebar

# Social links
st.sidebar.write("**Social Links**")

# GitHub-hosted logos
linkedin_logo_url = "https://raw.githubusercontent.com/Nishanthfiona/SMTP-Based-Email-Validation-Tool/refs/heads/main/linkedin.png"
github_logo_url = "https://raw.githubusercontent.com/Nishanthfiona/SMTP-Based-Email-Validation-Tool/refs/heads/main/github-logo.png"
email_logo_url = "https://raw.githubusercontent.com/Nishanthfiona/SMTP-Based-Email-Validation-Tool/refs/heads/main/communication.png"
website_logo_url = "https://raw.githubusercontent.com/Nishanthfiona/SMTP-Based-Email-Validation-Tool/refs/heads/main/internet.png"
medium_logo_url = "https://raw.githubusercontent.com/Nishanthfiona/SMTP-Based-Email-Validation-Tool/refs/heads/main/medium.png"
# Display each social link vertically
st.sidebar.markdown(
    f"""
    <div style="display: flex; align-items: center; margin-bottom: 10px;">
        <img src="{linkedin_logo_url}" width="20" style="margin-right: 10px;">
        <a href="https://linkedin.com/in/nishanthfiona12" target="_blank">LinkedIn</a>
    </div>
    <div style="display: flex; align-items: center; margin-bottom: 10px;">
        <img src="{github_logo_url}" width="20" style="margin-right: 10px;">
        <a href="https://github.com/Nishanthfiona" target="_blank">GitHub</a>
    </div>
    <div style="display: flex; align-items: center; margin-bottom: 10px;">
        <img src="{email_logo_url}" width="20" style="margin-right: 10px;">
        <a href="mailto:fiona12.nf@gmail.com">Email</a>
    </div>
    <div style="display: flex; align-items: center; margin-bottom: 10px;">
        <img src="{website_logo_url}" width="20" style="margin-right: 10px;">
        <a href="https://nishanthfiona.vercel.app/" target="_blank">Website</a>
    </div>
    <div style="display: flex; align-items: center; margin-bottom: 10px;">
        <img src="{medium_logo_url}" width="20" style="margin-right: 10px;">
        <a href="https://medium.com/@nishanthfiona" target="_blank">Blog</a>
    </div>
    """,
    unsafe_allow_html=True,
)



gmail_user = st.text_input("Gmail Address", value="")
gmail_app_password = st.text_input("Gmail App Password", type="password")
input_excel = st.file_uploader("Upload Excel File", type=["xlsx"])

if input_excel:
    df = pd.read_excel(input_excel)
    input_filename = input_excel.name.split('.')[0]  # Extract filename without extension
    st.write("Data Preview", df.head())

    # Dynamically populate the email column selection dropdown with column names
    email_column = st.selectbox("Select Email Column", options=df.columns)

    start_row = st.number_input("Start Row", min_value=1, value=1)
    end_row = st.number_input("End Row", min_value=start_row, value=start_row + 9)

    if st.button("Start Validation"):
        process_emails(input_excel, gmail_user, gmail_app_password, start_row, end_row, email_column=email_column)

    # Display valid and invalid emails
    if not st.session_state['valid_emails'].empty:
        st.subheader("Valid Emails")
        st.dataframe(st.session_state['valid_emails'], use_container_width=True)

        valid_filename = generate_filename(input_filename, start_row, end_row, 'valid')
        valid_emails_excel = to_excel(st.session_state['valid_emails'])
        st.download_button(
            label="Download Valid Emails",
            data=valid_emails_excel,
            file_name=valid_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if not st.session_state['invalid_emails'].empty:
        st.subheader("Invalid Emails")
        st.dataframe(st.session_state['invalid_emails'], use_container_width=True)

        invalid_filename = generate_filename(input_filename, start_row, end_row, 'invalid')
        invalid_emails_excel = to_excel(st.session_state['invalid_emails'])
        st.download_button(
            label="Download Invalid Emails",
            data=invalid_emails_excel,
            file_name=invalid_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
