import pandas as pd
import smtplib
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="📧 Bulk Email Sender", layout="centered")
st.title("📧 Bulk Email Sender")

# Step 1: Upload Excel file
excel_file = st.file_uploader("Upload Excel file", type=["xlsx"])

# Step 2: Input Gmail credentials
sender_email = st.text_input("Your Email Address", placeholder="your@email.com")
app_password = st.text_input("App Password", type="password")

# Step 3: Input CC, BCC, Subject, and Message
cc_emails_input = st.text_area("CC Emails (separate with commas or paste multiple lines)", height=100, placeholder="cc1@example.com\ncc2@example.com")
bcc_emails_input = st.text_area("BCC Emails (separate with commas or paste multiple lines)", height=100, placeholder="bcc1@example.com\nbcc2@example.com")
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

# Helper function to validate email address
def is_valid_email(email):
    return re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", email)

# Process CC and BCC inputs
cc_emails = [email.strip() for line in cc_emails_input.splitlines() for email in line.split(',') if email.strip() and is_valid_email(email.strip())]
bcc_emails = [email.strip() for line in bcc_emails_input.splitlines() for email in line.split(',') if email.strip() and is_valid_email(email.strip())]

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
            if not (sender_email and app_password):
                st.warning("⚠️ Please provide your email and app password.")
            else:
                log_data = []
                success_count = 0
                failed_count = 0
                for index, row in edited_df.iterrows():
                    recipient = row['Email']
                    name = row.get('Name', 'Customer')
                    password = row.get('Password', 'Not Provided')

                    if not is_valid_email(recipient):
                        st.error(f"❌ Invalid email format for {name} ({recipient}), skipping.")
                        failed_count += 1
                        continue

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
                        with smtplib.SMTP("smtp.gmail.com", 587) as server:
                            server.starttls()
                            server.login(sender_email, app_password)
                            response = server.sendmail(sender_email, to_addrs, msg.as_string())

                        if recipient not in response:
                            st.success(f"✅ Email sent to {name} ({recipient})")
                            success_count += 1
                            log_data.append([name, recipient, "Success", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                        else:
                            st.error(f"❌ Failed to send email to {name} ({recipient}): SMTP did not accept recipient")
                            failed_count += 1
                            log_data.append([name, recipient, "SMTP Rejected", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                    except Exception as e:
                        st.error(f"❌ Failed to send email to {name} ({recipient}): {e}")
                        failed_count += 1
                        log_data.append([name, recipient, f"Exception: {e}", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

                st.info(f"✅ Total Success: {success_count}, ❌ Total Failed: {failed_count}")

                # Log to downloadable CSV
                log_df = pd.DataFrame(log_data, columns=["Name", "Email", "Status", "Timestamp"])
                csv_buffer = BytesIO()
                log_df.to_csv(csv_buffer, index=False)
                st.download_button("📥 Download Log CSV", data=csv_buffer.getvalue(), file_name="email_log.csv", mime="text/csv")

                st.warning("⚠️ Note: To track whether an email is opened, you'll need to integrate with a 3rd-party service that provides open tracking or embed a tracking pixel. Gmail and most email clients block read receipts by default for privacy reasons.")
