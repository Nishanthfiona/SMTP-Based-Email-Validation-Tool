import smtplib
import re
import dns.resolver
from time import time, sleep
import pandas as pd
import streamlit as st

# Validate email syntax using regex
def is_valid_syntax(email):
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return re.match(email_regex, email) is not None

# Verify the domain has MX records
def has_valid_mx_record(domain):
    try:
        dns.resolver.resolve(domain, 'MX')
        return True
    except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN):
        return False

# Perform SMTP validation using Gmail's SMTP server
def is_email_valid_smtp(email, gmail_user, gmail_app_password):
    try:
        if not is_valid_syntax(email):
            return False
        
        domain = email.split('@')[1]
        if not has_valid_mx_record(domain):
            return False

        server = smtplib.SMTP("smtp.gmail.com", 587, timeout=30)
        server.starttls()
        server.login(gmail_user, gmail_app_password)
        server.helo()
        server.mail(gmail_user)
        code, message = server.rcpt(email)

        server.quit()
        return code == 250
    except Exception as e:
        return False

# Streamlit interface
def email_verification_app():
    st.title('Email Verification App')

    # User inputs
    gmail_user = st.text_input("Enter your Gmail username", "email you enter will be used to verify other emails")
    gmail_app_password = st.text_input("Enter your Gmail app password(Note: No passwords will be saved )", type="password")
    
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
    
    if uploaded_file is not None:
        # Read the uploaded file
        df = pd.read_excel(uploaded_file)
        
        email_column = st.text_input("Enter the email column name", "Email")
        
        start_row = st.number_input("Start Row", min_value=0, value=0)
        end_row = st.number_input("End Row", min_value=start_row, value=start_row + 1000)

        if st.button('Start Validation'):
            validate_emails(df, email_column, gmail_user, gmail_app_password, start_row, end_row)

# Main function to validate emails
def validate_emails(df, email_column, gmail_user, gmail_app_password, start_row=None, end_row=None):
    start_time = time()

    if start_row is not None and end_row is not None:
        df = df.iloc[start_row:end_row]

    validated_data = []
    invalid_data = []

    for index, row in df.iterrows():
        email = row[email_column]
        email_start_time = time()
        is_valid = is_email_valid_smtp(email, gmail_user, gmail_app_password)
        time_taken = time() - email_start_time

        if is_valid:
            validated_data.append({**row.to_dict(), "Validation Status": "Valid", "Validation Time (s)": time_taken})
        else:
            invalid_data.append({**row.to_dict(), "Validation Status": "Invalid", "Validation Time (s)": time_taken})

        sleep(30)  # Pause to avoid server throttling

    # Show results in Streamlit
    if validated_data:
        st.write("Valid Emails:")
        st.write(pd.DataFrame(validated_data))

    if invalid_data:
        st.write("Invalid Emails:")
        st.write(pd.DataFrame(invalid_data))

    total_time = (time() - start_time) / 3600
    st.write(f"Total time taken: {total_time:.2f} hours")

if __name__ == "__main__":
    email_verification_app()
