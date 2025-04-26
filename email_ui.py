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
import time

# Page settings
st.set_page_config(page_title="📧 Bulk Email Sender", layout="wide")
st.title("📧 Bulk Email Sender")
st.sidebar.image("https://www.invesmate.com/assets/images/logo.png", width=200)
st.sidebar.markdown("Developed by Invesmate Admin Team")

# Step 1: Upload file
st.header("1️⃣ Upload Recipient Data")
file = st.file_uploader("Upload an Excel or CSV file", type=["xlsx", "csv"])
if "data" not in st.session_state:
    st.session_state.data = None
valid_file = False

if file:
    try:
        if file.name.endswith(".csv"):
            df = pd.read_csv(file)
        elif file.name.endswith(".xlsx"):
            df = pd.read_excel(file)

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
            st.session_state.data = df.copy()
            st.success("✅ File uploaded and validated successfully!")
        else:
            st.error("❗ Required columns missing: name, sender email, email id, password")
            st.session_state.data = None
    except Exception as e:
        st.error(f"❌ Failed to read file: {e}")
        st.session_state.data = None

# Step 2: Email credentials
st.header("2️⃣ Setup Email Credentials")
with st.expander("🔒 Gmail Login"):
    sender_email = st.text_input("Gmail Address", placeholder="you@gmail.com")
    app_password = st.text_input("Gmail App Password", type="password")

# Step 3: Email Settings
st.header("3️⃣ Configure Email Settings")
with st.expander("✉️ CC / BCC / Subject Settings"):
    cc_emails_input = st.text_area("CC Emails (comma/line-separated)", height=70)
    bcc_emails_input = st.text_area("BCC Emails (comma/line-separated)", height=70)
    subject = st.text_input("Email Subject", value="Welcome to Invesmate!")
    delay = st.slider("⏱ Delay between emails (seconds)", min_value=0, max_value=60, value=2)

# Helpers
def is_valid_email(email):
    return re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", str(email))

def email_exists(email):
    try:
        domain = email.split('@')[-1]
        answers = dns.resolver.resolve(domain, 'MX')
        return len(answers) > 0
    except:
        return False

# Process CC, BCC
cc_emails = [email.strip() for line in cc_emails_input.splitlines() for email in line.split(',') if email.strip() and is_valid_email(email.strip())]
bcc_emails = [email.strip() for line in bcc_emails_input.splitlines() for email in line.split(',') if email.strip() and is_valid_email(email.strip())]

# Step 4: Compose email
st.header("4️⃣ Compose Email Body")
html_body = st_quill(
    value="""
<p><strong>Dear {name},</strong></p>
<p>Welcome to Invesmate! Your company account has been created.</p>
<p><strong>Email:</strong> {id}<br><strong>Temporary Password:</strong> {password}</p>
<p>🔗 Access your account: <a href='https://outlook.office.com/mail/'>Outlook</a></p>
<p>For any help, contact Admin Support.</p>
<p>Regards,<br><strong>Invesmate Team</strong></p>
<img src='https://www.invesmate.com/tracking_open.png?email={email}' width='1' height='1' style='display:none'>
""",
    html=True,
    key="rich_email_body"
)

with st.expander("👁 Preview Sample Email"):
    if st.session_state.data is not None:
        preview = html_body.format(name="John Doe", email="john@example.com", id="john@example.com", password="pass123")
        st.markdown(preview, unsafe_allow_html=True)

# Step 5: Edit & Send
if st.session_state.data is not None:
    st.header("5️⃣ Review and Send Emails")
    st.session_state.data["Send"] = st.session_state.data.apply(
        lambda row: all(pd.notna([row['Name'], row['Email'], row['ID'], row['Password']]))
                    and is_valid_email(row['Email'])
                    and is_valid_email(row['ID']), axis=1
    )

    st.subheader("🛠 Edit or Add Recipients")

    def highlight_invalid(val):
        if not is_valid_email(val):
            return 'background-color: #FFD6D6'
        return ''

    st.dataframe(
        st.session_state.data.style.applymap(highlight_invalid, subset=['ID']),
        use_container_width=True
    )

    edited_df = st.data_editor(
        st.session_state.data,
        num_rows="dynamic",
        use_container_width=True,
        column_config={"Send": st.column_config.CheckboxColumn(label="Send", default=True)}
    )

    if st.button("➕ Add New Entry"):
        new_row = {
            'Name': '',
            'Email': '',
            'ID': '',
            'Password': '',
            'Send': True
        }
        st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)

    st.session_state.data = edited_df

    if st.button("🚀 Send Bulk Emails"):
        if not (sender_email and app_password):
            st.warning("⚠️ Please provide Gmail address and app password.")
        else:
            log_data = []
            success, failure = 0, 0
            for index, row in st.session_state.data.iterrows():
                if not row.get("Send", True):
                    continue

                recipient = row['Email']
                name = row.get('Name', 'User')
                user_id = row.get('ID')
                password = row.get('Password')

                if not is_valid_email(recipient) or not email_exists(recipient) or not is_valid_email(user_id):
                    st.error(f"❌ Skipping {name}: Invalid email or ID.")
                    failure += 1
                    continue

                msg = MIMEMultipart("alternative")
                msg['From'] = sender_email
                msg['To'] = recipient
                msg['Subject'] = subject

                if cc_emails:
                    msg['Cc'] = ", ".join(cc_emails)

                filled_body = html_body.format(name=name, email=recipient, id=user_id, password=password)
                msg.attach(MIMEText(filled_body, 'html'))

                to_addresses = [recipient] + cc_emails + bcc_emails

                try:
                    with smtplib.SMTP("smtp.gmail.com", 587) as server:
                        server.starttls()
                        server.login(sender_email, app_password)
                        server.sendmail(sender_email, to_addresses, msg.as_string())
                    st.success(f"✅ Sent email to {name} ({recipient})")
                    success += 1
                    log_data.append([name, recipient, "Success", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                except Exception as e:
                    st.error(f"❌ Failed to send to {name}: {e}")
                    failure += 1
                    log_data.append([name, recipient, f"Failed: {e}", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

                time.sleep(delay)

            st.info(f"📢 Summary: {success} Success ✅ | {failure} Failed ❌")

            log_df = pd.DataFrame(log_data, columns=["Name", "Email", "Status", "Timestamp"])
            buffer = BytesIO()
            log_df.to_csv(buffer, index=False)
            st.download_button("📥 Download Email Log", data=buffer.getvalue(), file_name="email_log.csv", mime="text/csv")
