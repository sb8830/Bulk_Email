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
cc_emails_input = st.text_area("CC Emails (one per line)", placeholder="cc1@example.com\ncc2@example.com")
bcc_emails_input = st.text_area("BCC Emails (one per line)", placeholder="bcc1@example.com\nbcc2@example.com")
subject = st.text_input("Email Subject", value="Welcome to Our Platform!")
plain_body = st.text_area("Email Body (Plain Text with Placeholders)", value="""
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
        st.error("❗ The Excel file must contain at least 'Name' and 'Email' columns.")
    else:
        # Step 5: Button to send emails
        if st.button("📬 Send Emails"):
            if not (sender_email and app_password and smtp_host and smtp_port):
                st.warning("⚠️ Please provide all SMTP and login details.")
            else:
                success_count = 0
                failed_count = 0
                for index, row in edited_df.iterrows():
                    recipient = row['Email']
                    name = row.get('Name', 'Customer')
                    password = row.get('Password', 'Not Provided')

                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = recipient
                    msg['Subject'] = subject

                    if cc_emails:
                        msg['Cc'] = ", ".join(cc_emails)

                    body_content = plain_body.format(name=name, email=recipient, password=password)
                    msg.attach(MIMEText(body_content, 'plain'))

                    to_addrs = [recipient] + cc_emails + bcc_emails

                    try:
                        with smtplib.SMTP(smtp_host, smtp_port) as server:
                            server.starttls()
                            server.login(sender_email, app_password)
                            server.sendmail(sender_email, to_addrs, msg.as_string())

                        st.success(f"✅ Email sent to {name} ({recipient})")
                        success_count += 1
                    except Exception as e:
                        st.error(f"❌ Failed to send email to {name} ({recipient}): {e}")
                        failed_count += 1

                st.info(f"✅ Total Success: {success_count}, ❌ Total Failed: {failed_count}")
