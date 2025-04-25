import pandas as pd
import smtplib
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
from streamlit_quill import st_quill
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Bulk Email Sender", layout="centered")
st.title("📧 Bulk Email Sender")

# Step 1: Upload Excel file
excel_file = st.file_uploader("Upload Excel file", type=["xlsx"])

# Step 2: Input Gmail credentials
sender_email = st.text_input("Your Email Address", placeholder="your@email.com")
app_password = st.text_input("App Password", type="password")

# Step 3: Input CC, BCC, Subject
cc_emails_input = st.text_area("CC Emails (separate with commas or paste multiple lines)", height=100, placeholder="cc1@example.com\ncc2@example.com")
bcc_emails_input = st.text_area("BCC Emails (separate with commas or paste multiple lines)", height=100, placeholder="bcc1@example.com\nbcc2@example.com")
subject = st.text_input("Email Subject", value="Welcome to Our Platform!")

# Step 4: Gmail-style Email Body Editor
st.subheader("📄 Compose Email Body (Rich HTML Format)")
st.markdown("""
Customize your email below using the editor — just like Gmail:
- Use **bold**, *italic*, <u>underline</u>, 🎨 text/background color  
- Create 🔗 hyperlinks  
- Use bullet points, numbered lists  
- Add indentation (Tab / Shift+Tab)
""", unsafe_allow_html=True)

html_body = st_quill(
    value="""
<p><strong>Dear {name},</strong></p>

<p>🎉 Welcome to our platform! Your account has been successfully created. Below are your login credentials:</p>

<ul>
  <li><strong>Username:</strong> {email}</li>
  <li><strong>Password:</strong> {password}</li>
</ul>

<p>Please <a href="https://yourwebsite.com/login" target="_blank">click here to log in</a> and change your password at your earliest convenience.</p>

<p style="background-color:#f4f4f4; padding:10px; border-left: 4px solid #2196F3;"><em>Pro Tip:</em> Bookmark your dashboard for easy access.</p>

<p><em>Best regards,</em><br><strong>The Support Team</strong></p>
""",
    html=True,
    key="rich_editor"
)

# Optional Preview
with st.expander("🔍 Preview Final Email with Sample Data"):
    st.markdown(html_body.format(name="John Doe", email="john@example.com", password="12345678"), unsafe_allow_html=True)

# Email validator
def is_valid_email(email):
    return re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", email)

# Process CC and BCC inputs
cc_emails = [email.strip() for line in cc_emails_input.splitlines() for email in line.split(',') if email.strip() and is_valid_email(email.strip())]
bcc_emails = [email.strip() for line in bcc_emails_input.splitlines() for email in line.split(',') if email.strip() and is_valid_email(email.strip())]

# Step 5: Excel Upload & Send
if excel_file:
    df = pd.read_excel(excel_file)
    df["Send"] = True
    st.subheader("📄 Preview and Modify Excel Data")
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_config={"Send": st.column_config.CheckboxColumn(label="Send", default=True)})

    if 'Email' not in edited_df.columns or 'Name' not in edited_df.columns:
        st.error("❗ The Excel file must contain at least 'Name' and 'Email' columns.")
    else:
        if st.button("📬 Send Emails"):
            if not (sender_email and app_password):
                st.warning("⚠️ Please provide your email and app password.")
            else:
                log_data = []
                success_count = 0
                failed_count = 0
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

                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = recipient
                    msg['Subject'] = subject

                    if cc_emails:
                        msg['Cc'] = ", ".join(cc_emails)

                    try:
                        formatted_html_body = html_body.format(name=name, email=recipient, password=password)
                        msg.attach(MIMEText(formatted_html_body, 'html'))

                        to_addrs = [recipient] + cc_emails + bcc_emails

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

                # Log file
                log_df = pd.DataFrame(log_data, columns=["Name", "Email", "Status", "Timestamp"])
                csv_buffer = BytesIO()
                log_df.to_csv(csv_buffer, index=False)
                st.download_button("📥_
