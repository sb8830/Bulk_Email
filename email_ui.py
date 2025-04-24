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

# Step 2: Input Gmail credentials
sender_email = st.text_input("Your Gmail", placeholder="your@gmail.com")
app_password = st.text_input("App Password", type="password")

# Step 3: Input email subject and message
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
                st.warning("⚠️ Please provide Gmail and App Password.")
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

                    body_content = plain_body.format(name=name, email=recipient, password=password)
                    msg.attach(MIMEText(body_content, 'plain'))

                    try:
                        with smtplib.SMTP("smtp.gmail.com", 587) as server:
                            server.starttls()
                            server.login(sender_email, app_password)
                            server.send_message(msg)

                        st.success(f"✅ Email sent to {name} ({recipient})")
                        success_count += 1
                    except Exception as e:
                        st.error(f"❌ Failed to send email to {name} ({recipient}): {e}")
                        failed_count += 1

                st.info(f"✅ Total Success: {success_count}, ❌ Total Failed: {failed_count}")
