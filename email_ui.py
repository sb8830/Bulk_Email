import streamlit as st
import pandas as pd
import re
from io import BytesIO
from streamlit_quill import st_quill
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

st.set_page_config(page_title="📧 Bulk Email Sender", layout="wide")
st.title("📧 Bulk Email Sender with Validation & Demo Files")

# Helper functions
def is_valid_email(email):
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return bool(re.match(pattern, email))

def is_strong_password(password):
    pattern = r'^(?=.*[A-Z])(?=.*[a-z])(?=.*\d)(?=.*[@$!%*#?&])[A-Za-z\d@$!#%*?&]{8,}$'
    return bool(re.match(pattern, password))

# Sample data for demo download
demo_data = {
    "Name": ["John Doe", "Jane Smith"],
    "Email": ["john.doe@example.com", "jane.smith@example.com"],
    "Password": ["Pass@1234", "Secure#2024"]
}
demo_df = pd.DataFrame(demo_data)

st.header("📤 Upload User List (CSV or Excel)")
uploaded_file = st.file_uploader("Upload file with Name, Email, Password", type=["csv", "xlsx"])

# Demo file downloads
st.download_button("⬇️ Download Demo CSV", data=demo_df.to_csv(index=False).encode(), file_name="demo_users.csv", mime="text/csv")
excel_buffer = BytesIO()
demo_df.to_excel(excel_buffer, index=False)
st.download_button("⬇️ Download Demo Excel", data=excel_buffer.getvalue(), file_name="demo_users.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Email form
st.subheader("📬 Email Settings")
sender_email = st.text_input("Your Gmail Address")
app_password = st.text_input("Gmail App Password", type="password")
subject = st.text_input("Subject", value="Welcome to the Team!")

# Rich Text Editor for Body
st.subheader("📄 Email Body")
default_body = """
<p>Dear {name},</p>
<p>Welcome to the team! Your new company email account is:</p>
<p><strong>Email:</strong> {email}<br>
<strong>Temporary Password:</strong> {password}</p>
<p>Regards,<br>Your IT Team</p>
"""
html_body = st_quill(value=default_body, html=True)

# Process uploaded file
if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        if not {'Name', 'Email', 'Password'}.issubset(df.columns):
            st.error("❌ The file must contain 'Name', 'Email', and 'Password' columns.")
        else:
            st.success("✅ File is valid! Preview and send emails below.")
            df['Send'] = True
            df = st.data_editor(df, num_rows="dynamic")

            if st.button("📬 Send Emails"):
                if not (sender_email and app_password):
                    st.warning("⚠️ Please enter your Gmail and app password.")
                else:
                    success_count = 0
                    fail_count = 0
                    log_data = []
                    for _, row in df.iterrows():
                        if not row['Send']:
                            continue
                        name = row['Name']
                        recipient = row['Email']
                        password = row['Password']

                        if not is_valid_email(recipient):
                            log_data.append([name, recipient, "Invalid Email"])
                            fail_count += 1
                            continue

                        msg = MIMEMultipart()
                        msg['From'] = sender_email
                        msg['To'] = recipient
                        msg['Subject'] = subject
                        final_body = html_body.format(name=name, email=recipient, password=password)
                        msg.attach(MIMEText(final_body, 'html'))

                        try:
                            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                                server.starttls()
                                server.login(sender_email, app_password)
                                server.sendmail(sender_email, recipient, msg.as_string())
                            log_data.append([name, recipient, "Success"])
                            success_count += 1
                        except Exception as e:
                            log_data.append([name, recipient, f"Failed: {e}"])
                            fail_count += 1

                    st.success(f"✅ Sent: {success_count}, ❌ Failed: {fail_count}")
                    log_df = pd.DataFrame(log_data, columns=["Name", "Email", "Status"])
                    log_csv = BytesIO()
                    log_df.to_csv(log_csv, index=False)
                    st.download_button("📥 Download Email Log", data=log_csv.getvalue(), file_name="email_log.csv", mime="text/csv")
    except Exception as e:
        st.error(f"⚠️ Error: {e}")
