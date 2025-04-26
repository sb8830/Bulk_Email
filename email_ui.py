import pandas as pd
import smtplib
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
from streamlit_quill import st_quill
from io import BytesIO
from datetime import datetime
import dns.resolver

st.set_page_config(page_title="Bulk Email Sender", layout="wide")
st.title("📧 Bulk Email Sender")

# Step 1: Upload Excel or CSV file
file = st.file_uploader("Upload Excel or CSV file", type=["xlsx", "csv"])
df = None
valid_file = False

if file:
    try:
        if file.name.endswith(".csv"):
            df = pd.read_csv(file)
        elif file.name.endswith(".xlsx"):
            df = pd.read_excel(file)

        # Normalize column names to lowercase for comparison
        df.columns = [col.lower().strip() for col in df.columns]

        required_columns = {'name', 'sender email', 'email id', 'password'}
        if required_columns.issubset(set(df.columns)):
            df.rename(columns={
                'name': 'Name',
                'sender email': 'Email',
                'email id': 'ID',
                'password': 'Password'
            }, inplace=True)
            valid_file = True
        else:
            st.error("❗ File must contain the following columns (case-insensitive): name, sender email, email id, password")
            df = None

    except Exception as e:
        st.error(f"❌ Failed to read file: {e}")
        df = None

# Step 2: Input Gmail credentials
with st.expander("🔐 Email Credentials"):
    sender_email = st.text_input("Your Gmail Address", placeholder="your@email.com")
    app_password = st.text_input("Gmail App Password", type="password")

# Step 3: Input CC, BCC, Subject
with st.expander("📬 Email Settings"):
    cc_emails_input = st.text_area("CC Emails (comma/line-separated)", height=80)
    bcc_emails_input = st.text_area("BCC Emails (comma/line-separated)", height=80)
    subject = st.text_input("Email Subject", value="Welcome to Our Platform!")

# Helper function to validate email address format
def is_valid_email(email):
    return re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", email)

# Optional: MX record validation (basic email existence check)
def email_exists(email):
    domain = email.split('@')[-1]
    try:
        answers = dns.resolver.resolve(domain, 'MX')
        return len(answers) > 0
    except:
        return False

# Process CC and BCC inputs
cc_emails = [email.strip() for line in cc_emails_input.splitlines() for email in line.split(',') if email.strip() and is_valid_email(email.strip())]
bcc_emails = [email.strip() for line in bcc_emails_input.splitlines() for email in line.split(',') if email.strip() and is_valid_email(email.strip())]

# Email body and signature editor
st.subheader("📄 Email Body")
html_body = st_quill(
    value="""
<p><strong>Dear {name},</strong></p>
<p>We are excited to announce that we have created new company email accounts for all employees using Microsoft Outlook!</p>
<p><strong>Your New Email Address:</strong> <span style='background-color: #FFFF00'>{id}</span><br>
<strong>Temporary Password:</strong> <span style='background-color: #90EE90'>{password}</span></p>
<p><strong>Access Links:</strong></p>
<ul>
  <li><a href='https://outlook.office.com/mail/'>Outlook Web App</a></li>
  <li><a href='https://m365.cloud.microsoft/launch/word'>Docs</a></li>
</ul>
<p>Regards,<br><strong>Your Name</strong></p>
<img src='https://yourserver.com/track_open.png?email={email}' width='1' height='1' style='display:none'>
""",
    html=True,
    key="rich_email_body"
)

# Preview
with st.expander("🔍 Preview Final Email with Sample Data"):
    if valid_file:
        preview_filled = html_body.format(name="John Doe", email="john@example.com", id="john@example.com", password="12345678")
        st.markdown(preview_filled, unsafe_allow_html=True)

# Step 5: Load Excel/CSV and send emails
if valid_file:
    df["Send"] = True
    st.subheader("📄 Preview and Modify Data")
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_config={"Send": st.column_config.CheckboxColumn(label="Send", default=True)})

    if st.button("📬 Send Emails"):
        if not (sender_email and app_password):
            st.warning("⚠️ Please provide your email and app password.")
        else:
            log_data = []
            success_count, failed_count = 0, 0
            for index, row in edited_df.iterrows():
                if not row.get("Send", True):
                    continue

                recipient = row['Email']
                name = row.get('Name', 'Customer')
                password = row.get('Password', 'Not Provided')
                user_id = row.get('ID', 'NA')

                if not is_valid_email(recipient) or not email_exists(recipient):
                    st.error(f"❌ Invalid or non-existent email for {name} ({recipient}), skipping.")
                    failed_count += 1
                    continue

                msg = MIMEMultipart("alternative")
                msg['From'] = sender_email
                msg['To'] = recipient
                msg['Subject'] = subject

                if cc_emails:
                    msg['Cc'] = ", ".join(cc_emails)

                filled_body = html_body.format(name=name, email=recipient, id=user_id, password=password)
                msg.attach(MIMEText(filled_body, 'html'))

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
                        st.error(f"❌ SMTP error for {name} ({recipient})")
                        failed_count += 1
                        log_data.append([name, recipient, "SMTP Rejected", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                except Exception as e:
                    st.error(f"❌ Failed for {name} ({recipient}): {e}")
                    failed_count += 1
                    log_data.append([name, recipient, f"Exception: {e}", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

            st.info(f"✅ Sent: {success_count}, ❌ Failed: {failed_count}")

            log_df = pd.DataFrame(log_data, columns=["Name", "Email", "Status", "Timestamp"])
            csv_buffer = BytesIO()
            log_df.to_csv(csv_buffer, index=False)
            st.download_button("📥 Download Log CSV", data=csv_buffer.getvalue(), file_name="email_log.csv", mime="text/csv")
