import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="📧 Bulk Email Sender", layout="centered")
st.title("📧 Bulk Email Sender")

# Step 1: Upload Excel file
excel_file = st.file_uploader("Upload Excel file", type=["xlsx"])

# Step 2: Input SMTP Host, Port, and Credentials
smtp_host = st.text_input("SMTP Host", value="smtp.gmail.com")
smtp_port = st.number_input("SMTP Port", value=587, step=1)
sender_email = st.text_input("Your Email Address", placeholder="your@email.com")
app_password = st.text_input("App Password", type="password")

# Step 3: Input CC, BCC, Subject, and Message
cc_emails_input = st.text_area("CC Emails (paste multiple, one per line)", height=100, placeholder="cc1@example.com\ncc2@example.com")
bcc_emails_input = st.text_area("BCC Emails (paste multiple, one per line)", height=100, placeholder="bcc1@example.com\nbcc2@example.com")
subject = st.text_input("Email Subject", value="Welcome to Our Platform!")
plain_body = st.text_area("Email Body (Plain Text with Placeholders)", height=250, value="""
Dear {name},

Welcome to our platform! Your account has been successfully created.

Username: {email}
Password: {password}

Please log in and change your password at your earliest convenience.

Regards,
Team
""")

# Process CC and BCC inputs
cc_emails = [email.strip() for email in cc_emails_input.splitlines() if email.strip()]
bcc_emails = [email.strip() for email in bcc_emails_input.splitlines() if email.strip()]

# Step 4: Validate and Show Excel Data
if excel_file:
    df = pd.read_excel(excel_file)
    st.subheader("📄 Preview and Modify Excel Data")
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    if 'Email' not in edited_df.columns or 'Name' not in edited_df.columns:
