import pandas as pd
import smtplib
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
from streamlit_quill import st_quill
from io import BytesIO
from datetime import datetime

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

        # Normalize column names to lowercase for comparison
        normalized_columns = {col.lower(): col for col in df.columns}
        required_columns = {'name', 'email', 'password'}

        if required_columns.issubset(normalized_columns):
            # Rename columns for consistency
            df.rename(columns={normalized_columns['name']: 'Name',
                               normalized_columns['email']: 'Email',
                               normalized_columns['password']: 'Password'}, inplace=True)
            valid_file = True
        else:
            st.error("❗ File must contain the following columns: Name, Email, Password")
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

# Helper function to validate email address
def is_valid_email(email):
    return re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", email)

# Process CC and BCC inputs
cc_emails = [email.strip() for line in cc_emails_input.splitlines() for email in line.split(',') if email.strip() and is_valid_email(email.strip())]
bcc_emails = [email.strip() for line in bcc_emails_input.splitlines() for email in line.split(',') if email.strip() and is_valid_email(email.strip())]

# Email body and signature editor
st.subheader("📄 Email Body")
html_body = st_quill(
    value="""
<p><strong>Dear {name},</strong></p>
<p>We're excited to welcome you to our platform! 🎉</p>
<ul>
  <li><span style=\"background-color: #ffff00;\"><strong>Username:</strong></span> {email}</li>
  <li><span style=\"background-color: #90ee90;\"><strong>Password:</strong></span> {password}</li>
</ul>
<p>Please <a href=\"https://yourwebsite.com/login\" target=\"_blank\">click here</a> to log in and change your password.</p>
<p style=\"padding: 10px; border-left: 4px solid #2196F3; background-color: #f1f1f1;\">
  <em>Tip:</em> Keep your login credentials safe and do not share them with others.
</p>
<p style=\"margin-top: 30px;\">
  Best regards,<br>
  <strong>Your Name</strong><br>
  Customer Success Team<br>
  <a href=\"https://yourwebsite.com\">yourwebsite.com</a>
</p>
<img src=\"https://yourserver.com/track_open.png?email={email}\" width=\"1\" height=\"1\" style=\"display:none\">  <!-- Tracking Pixel -->
""",
    html=True,
    key="rich_email_body"
)

# Preview
with st.expander("🔍 Preview Final Email with Sample Data"):
    if valid_file:
        preview_filled = html_body.format(name="John Doe", email="john@example.com", password="12345678")
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

                if not is_valid_email(recipient):
                    st.error(f"❌ Invalid email format for {name} ({recipient}), skipping.")
                    failed_count += 1
                    continue

                msg = MIMEMultipart("alternative")
                msg['From'] = sender_email
                msg['To'] = recipient
                msg['Subject'] = subject

                if cc_emails:
                    msg['Cc'] = ", ".join(cc_emails)

                filled_body = html_body.format(name=name, email=recipient, password=password)
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
