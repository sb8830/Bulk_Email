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
normalized_columns = {}

if file:
    try:
        if file.name.endswith(".csv"):
            df = pd.read_csv(file)
        elif file.name.endswith(".xlsx"):
            df = pd.read_excel(file)

        # Normalize column names to lowercase
        normalized_columns = {col.lower(): col for col in df.columns}
        required_columns = {'name', 'email', 'password'}

        if required_columns.issubset(normalized_columns):
            df.rename(columns={
                normalized_columns['name']: 'Name',
                normalized_columns['email']: 'Email',
                normalized_columns['password']: 'Password'
            }, inplace=True)
            valid_file = True
        else:
            st.error("❗ File must contain columns: Name, Email, Password")
            df = None
    except Exception as e:
        st.error(f"❌ Failed to read file: {e}")
        df = None

# Step 2: Input Gmail credentials
with st.expander("🔐 Email Credentials"):
    sender_email = st.text_input("Your Gmail Address", placeholder="your@email.com")
    app_password = st.text_input("Gmail App Password", type="password")

# Step 3: CC, BCC, and Subject
with st.expander("📬 Email Settings"):
    cc_input = st.text_area("CC Emails (comma or newline separated)", height=80)
    bcc_input = st.text_area("BCC Emails (comma or newline separated)", height=80)
    subject = st.text_input("Email Subject", value="Welcome to Our Platform!")

# Email validation
def is_valid_email(email):
    return re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", email)

def email_exists(email):
    domain = email.split('@')[-1]
    try:
        answers = dns.resolver.resolve(domain, 'MX')
        return len(answers) > 0
    except:
        return False

# Parse CC/BCC
cc_emails = [e.strip() for line in cc_input.splitlines() for e in line.split(',') if e.strip() and is_valid_email(e.strip())]
bcc_emails = [e.strip() for line in bcc_input.splitlines() for e in line.split(',') if e.strip() and is_valid_email(e.strip())]

# Step 4: Compose Email
st.subheader("📄 Email Body")
html_body = st_quill(
    value="""
<p><strong>Dear {name},</strong></p>
<p style="text-align: justify;">We are excited to announce that we have created new company email accounts for all employees using Microsoft Outlook!</p>
<p><strong>Your New Email Address:</strong> <span style='background-color: #FFFF00'>{email}</span><br>
<strong>Temporary Password:</strong> <span style='background-color: #90EE90'>{password}</span></p>
<p>Access your account via Outlook Web App, set a secure password, and begin using the Microsoft 365 tools.</p>
<p>Regards,<br><strong>Your Name</strong><br>IT Support<br><a href='https://invesmate.com'>invesmate.com</a></p>
<img src='https://yourserver.com/track_open.png?email={email}' width='1' height='1' style='display:none'>
    """,
    html=True,
    key="rich_body"
)

# Step 5: Preview with sample data
with st.expander("🔍 Preview Final Email"):
    if valid_file:
        sample_preview = html_body.format(name="John Doe", email="john@example.com", password="12345678")
        st.markdown(sample_preview, unsafe_allow_html=True)

# Step 6: View/edit data and send emails
if valid_file:
    df["Send"] = True
    st.subheader("📋 Preview and Toggle Send")
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True,
                               column_config={"Send": st.column_config.CheckboxColumn(label="Send", default=True)})

    if st.button("📬 Send Emails Now"):
        if not sender_email or not app_password:
            st.warning("⚠️ Enter your email and app password to send.")
        else:
            success_count = 0
            failed_count = 0
            logs = []

            for idx, row in edited_df.iterrows():
                if not row.get("Send", True):
                    continue

                name = row.get('Name', 'User')
                recipient = row['Email']
                password = row['Password']

                if not is_valid_email(recipient) or not email_exists(recipient):
                    st.error(f"❌ Invalid or non-existent email for {name} ({recipient})")
                    failed_count += 1
                    logs.append([name, recipient, "Invalid Email", datetime.now()])
                    continue

                # Compose message
                msg = MIMEMultipart("alternative")
                msg['From'] = sender_email
                msg['To'] = recipient
                msg['Subject'] = subject
                if cc_emails:
                    msg['Cc'] = ", ".join(cc_emails)

                filled_body = html_body.format(name=name, email=recipient, password=password)
                msg.attach(MIMEText(filled_body, 'html'))
                to_list = [recipient] + cc_emails + bcc_emails

                try:
                    with smtplib.SMTP("smtp.gmail.com", 587) as server:
                        server.starttls()
                        server.login(sender_email, app_password)
                        response = server.sendmail(sender_email, to_list, msg.as_string())

                    if recipient not in response:
                        st.success(f"✅ Sent to {name} ({recipient})")
                        success_count += 1
                        logs.append([name, recipient, "Success", datetime.now()])
                    else:
                        st.error(f"❌ SMTP error for {name} ({recipient})")
                        failed_count += 1
                        logs.append([name, recipient, "SMTP Error", datetime.now()])
                except Exception as e:
                    st.error(f"❌ Exception for {name} ({recipient}): {e}")
                    failed_count += 1
                    logs.append([name, recipient, f"Exception: {e}", datetime.now()])

            st.info(f"📬 Emails sent: {success_count} | ❌ Failed: {failed_count}")

            log_df = pd.DataFrame(logs, columns=["Name", "Email", "Status", "Timestamp"])
            buffer = BytesIO()
            log_df.to_csv(buffer, index=False)
            st.download_button("📥 Download Email Log", buffer.getvalue(), file_name="email_log.csv", mime="text/csv")
