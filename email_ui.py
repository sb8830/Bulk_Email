import streamlit as st
import pandas as pd
import smtplib
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from streamlit_quill import st_quill
from io import BytesIO
from datetime import datetime
import time

# Page setup
st.set_page_config(page_title="📧 Bulk Email Sender", layout="wide")
st.title("📧 Bulk Email Sender")
st.sidebar.image("https://www.invesmate.com/assets/images/logo.png", width=200)
st.sidebar.markdown("Developed by Invesmate Admin Team")

# Dummy credentials
USER_CREDENTIALS = {
    "admin@invesmate.com": {"password": "admin123", "role": "admin"},
    "user@invesmate.com": {"password": "user123", "role": "user"},
}

# Initialize session state
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user_email = None
    st.session_state.user_role = None
    st.session_state.login_attempt = False
    st.session_state.data = None

# Login logic
def login_form():
    with st.form("Login"):
        st.subheader("🔐 Please log in to continue")
        email = st.text_input("Email")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")
        if submitted:
            user = USER_CREDENTIALS.get(email)
            if user and user["password"] == password:
                st.session_state.logged_in = True
                st.session_state.user_email = email
                st.session_state.user_role = user["role"]
            else:
                st.session_state.login_attempt = True

# Email validator
def is_valid_email(email):
    return bool(re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", str(email)))

# Highlight invalid cells
def highlight_invalid_cells(row):
    styles = [''] * len(row)
    if not is_valid_email(row['Email']):
        styles[row.index.get_loc('Email')] = 'background-color: #FFD6D6;'
    if not is_valid_email(row['ID']):
        styles[row.index.get_loc('ID')] = 'background-color: #FFD6D6;'
    if pd.isna(row['Password']) or row['Password'] == '':
        styles[row.index.get_loc('Password')] = 'background-color: #FFD6D6;'
    return styles

# Main App
def run_app():
    st.sidebar.success(f"✅ Logged in as {st.session_state.user_email} ({st.session_state.user_role})")

    st.header("1️⃣ Upload Recipient Data")
    file = st.file_uploader("Upload an Excel or CSV file", type=["xlsx", "csv"])
    if file:
        try:
            df = pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)
            df.columns = [col.lower().strip() for col in df.columns]
            required = {'name', 'sender email', 'email id', 'password'}
            if required.issubset(set(df.columns)):
                df.rename(columns={
                    'name': 'Name',
                    'sender email': 'Email',
                    'email id': 'ID',
                    'password': 'Password'
                }, inplace=True)
                df['Send'] = True
                st.session_state.data = df
                st.success("✅ File uploaded and validated successfully!")
            else:
                st.error("❗ Required columns: name, sender email, email id, password")
        except Exception as e:
            st.error(f"❌ Failed to read file: {e}")

    st.header("2️⃣ Compose Email Body")
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

    if st.session_state.user_role == "admin":
        st.header("3️⃣ Email Credentials and Settings")
        sender_email = st.text_input("Gmail Address", placeholder="you@gmail.com")
        app_password = st.text_input("Gmail App Password", type="password")

        with st.expander("✉️ CC / BCC / Subject Settings"):
            cc_input = st.text_area("CC Emails", height=50)
            bcc_input = st.text_area("BCC Emails", height=50)
            subject = st.text_input("Email Subject", value="Welcome to Invesmate!")
            delay = st.slider("Delay between emails (sec)", 0, 60, 2)

        cc_emails = [e.strip() for l in cc_input.splitlines() for e in l.split(',') if is_valid_email(e.strip())]
        bcc_emails = [e.strip() for l in bcc_input.splitlines() for e in l.split(',') if is_valid_email(e.strip())]

    if st.session_state.data is not None:
        st.header("4️⃣ Review & Send Emails")
        edited_df = st.data_editor(
            st.session_state.data,
            num_rows="dynamic",
            use_container_width=True,
            column_config={"Send": st.column_config.CheckboxColumn(label="Send", default=True)}
        )
        st.session_state.data = edited_df.copy()
        styled_df = edited_df.style.apply(highlight_invalid_cells, axis=1)
        st.dataframe(styled_df, use_container_width=True)

        if st.session_state.user_role == "admin":
            if st.button("🚀 Send Bulk Emails"):
                if not (sender_email and app_password):
                    st.warning("⚠️ Enter Gmail & App Password.")
                else:
                    progress = st.progress(0)
                    log_data = []
                    success, failure = 0, 0
                    to_send = edited_df[edited_df['Send'] == True]
                    total = len(to_send)

                    for i, (_, row) in enumerate(to_send.iterrows()):
                        recipient = row['Email']
                        name = row['Name']
                        user_id = row['ID']
                        pwd = row['Password']

                        if not (is_valid_email(recipient) and is_valid_email(user_id) and pwd):
                            st.error(f"❌ Skipping {name} ({recipient}) - Invalid data.")
                            failure += 1
                            continue

                        try:
                            msg = MIMEMultipart("alternative")
                            msg['From'] = sender_email
                            msg['To'] = recipient
                            msg['Subject'] = subject
                            if cc_emails: msg['Cc'] = ", ".join(cc_emails)
                            body = html_body.format(name=name, email=recipient, id=user_id, password=pwd)
                            msg.attach(MIMEText(body, 'html'))

                            all_recipients = [recipient] + cc_emails + bcc_emails

                            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                                server.starttls()
                                server.login(sender_email, app_password)
                                server.sendmail(sender_email, all_recipients, msg.as_string())

                            st.success(f"✅ Sent to {name}")
                            success += 1
                            log_data.append([name, recipient, "Success", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

                        except Exception as e:
                            st.error(f"❌ Failed to send to {name}: {str(e)}")
                            failure += 1
                            log_data.append([name, recipient, f"Failed: {str(e)}", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

                        time.sleep(delay)
                        progress.progress((i + 1) / total)

                    st.info(f"✅ Summary: {success} Sent | ❌ {failure} Failed")
                    log_df = pd.DataFrame(log_data, columns=["Name", "Email", "Status", "Timestamp"])
                    buffer = BytesIO()
                    log_df.to_csv(buffer, index=False)
                    st.download_button("📥 Download Log", data=buffer.getvalue(), file_name="email_log.csv", mime="text/csv")

# Routing
if not st.session_state.logged_in:
    login_form()
    if st.session_state.login_attempt:
        st.error("❌ Invalid login, try again.")
else:
    run_app()
