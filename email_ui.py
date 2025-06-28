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

# Demo data file for users to download
sample_data = pd.DataFrame({
    "Name": ["John Doe", "Jane Smith"],
    "Email": ["john.doe@example.com", "jane.smith@example.com"],
    "Password": ["Pass@1234", "Secure#2024"]
})

st.download_button(
    label="📥 Download Demo Excel File",
    data=sample_data.to_csv(index=False).encode(),
    file_name="demo_email_data.csv",
    mime="text/csv"
)

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
<p><strong>Good Evening,</strong></p>
<p>We are excited to announce that we have created new company email accounts for all employees using Microsoft Outlook! This upgrade is part of our ongoing effort to improve communication and collaboration within the company.</p>
<p><strong>Your New Email Address:</strong> {email}<br>
<strong>Temporary Password:</strong> {password}</p>
<p><strong>Important Email Account Transition Information:</strong></p>
<ul>
<li><strong>Google Drive Data and Emails:</strong> All emails and files from existing company domain(*@invesmategroup.com) user accounts have already been migrated to the new Outlook accounts.</li>
<li><strong>Google Workspace Account Deactivation(@invesmategroup.com):</strong> Your existing company domain(@invesmategroup.com) Gmail accounts will be deactivated on 30Th June, 2025. After this date, you will no longer be able to Open account from Gmail.</li>
<li><strong>Individual Gmail Account Deactivation(*.invesmategroup@.com in GSUTE):</strong> You will need to move any important files to your company OneDrive as soon as possible. Your G-Sute accounts will be disabled for company use within The mentioned date if anything missing in your one drive.</li>
</ul>
<p><strong>Accessing Your Accounts:</strong></p>
<ul>
<li>Outlook Web App: <a href='https://outlook.office.com/mail/'>https://outlook.office.com/mail/</a></li>
<li>Microsoft Teams: <a href='https://teams.microsoft.com/v2/'>https://teams.microsoft.com/v2/</a></li>
<li>OneDrive: <a href='https://admininvesmate360-my.sharepoint.com/'>https://admininvesmate360-my.sharepoint.com/</a></li>
<li>Excel: <a href='https://m365.cloud.microsoft/launch/excel'>https://m365.cloud.microsoft/launch/excel</a></li>
<li>Docs: <a href='https://m365.cloud.microsoft/launch/word'>https://m365.cloud.microsoft/launch/word</a></li>
<li>Microsoft Authenticator: <a href='https://play.google.com/store/apps/details?id=com.azure.authenticator'>Google Play</a></li>
</ul>
<p><strong>How to Login to Outlook:</strong></p>
<ol>
<li>Go to the Outlook Web App link provided above.</li>
<li>Enter your new email address.</li>
<li>Enter the temporary password provided above.</li>
<li>You will then be prompted to create a new, more secure password.</li>
<li>Follow the on-screen instructions to complete the login process.</li>
</ol>
<p><strong>Helpful Resources:</strong></p>
<ul>
<li>Outlook Setup</li>
<li>Microsoft Apps Quick Guide</li>
</ul>
<p><strong>Support:</strong> If you have any questions or need assistance with accessing your new account, please do not hesitate to contact us.</p>
<p>We are confident that these new tools will enhance our communication and productivity. We appreciate your cooperation during this transition.</p>
<p><strong>Thanks and Regards,</strong></p>
<img src='https://yourserver.com/track_open.png?email={email}' width='1' height='1' style='display:none'>
""",
    html=True,
    key="rich_body"
)

# Step 5: Preview with sample data
with st.expander("🔍 Preview Final Email"):
    sample_preview = html_body.format(name="John Doe", email="john@example.com", password="Pass@1234")
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
