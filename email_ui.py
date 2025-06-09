import streamlit as st
import pandas as pd
import smtplib
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from io import BytesIO
from datetime import datetime
import time
from streamlit_quill import st_quill

# Page config must be first
st.set_page_config(page_title="📧 Bulk Email Sender", layout="wide")

# Sample users (in production, use a secure store)
USERS = {
    "admin": {"password": "admin123", "role": "admin"},
    "user": {"password": "user123", "role": "user"}
}

# Helper Functions
def is_valid_email(email):
    return bool(re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", str(email)))

def highlight_invalid_cells(row):
    styles = [''] * len(row)
    if not is_valid_email(row['Email']):
        styles[row.index.get_loc('Email')] = 'background-color: #FFD6D6;'
    if not is_valid_email(row['ID']):
        styles[row.index.get_loc('ID')] = 'background-color: #FFD6D6;'
    if pd.isna(row['Password']) or row['Password'] == '':
        styles[row.index.get_loc('Password')] = 'background-color: #FFD6D6;'
    return styles

def logout():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.success("Logged out successfully.")
    st.experimental_rerun()

def login_page():
    st.title("🔐 Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username in USERS and USERS[username]["password"] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.role = USERS[username]["role"]
            st.success("Login successful!")
            st.experimental_rerun()
        else:
            st.error("Invalid credentials")

def run_app():
    if "logged_in" not in st.session_state:
        login_page()
        return

    st.sidebar.title("Welcome")
    st.sidebar.write(f"👤 Logged in as: `{st.session_state.username}` ({st.session_state.role})")
    if st.sidebar.button("🔓 Logout"):
        logout()

    st.title("📧 Bulk Email Sender")
    st.sidebar.image("https://www.invesmate.com/assets/images/logo.png", width=200)
    st.sidebar.markdown("Developed by Invesmate Admin Team")

    # Step 1: Upload
    st.header("1️⃣ Upload Recipient Data")
    file = st.file_uploader("Upload an Excel or CSV file", type=["xlsx", "csv"])
    if "data" not in st.session_state:
        st.session_state.data = None

    if file:
        try:
            df = pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)
            df.columns = [col.lower().strip() for col in df.columns]
            required = {'name', 'sender email', 'email id', 'password'}
            if required.issubset(set(df.columns)):
                df.rename(columns={'name': 'Name', 'sender email': 'Email', 'email id': 'ID', 'password': 'Password'}, inplace=True)
                df['Send'] = True
                st.session_state.data = df
                st.success("✅ File uploaded and validated successfully!")
            else:
                st.error("❗ Required columns missing: name, sender email, email id, password")
        except Exception as e:
            st.error(f"❌ Failed to read file: {e}")

    # Step 2: Credentials
    st.header("2️⃣ Setup Email Credentials")
    with st.expander("🔒 Gmail Login"):
        sender_email = st.text_input("Gmail Address", placeholder="you@gmail.com")
        app_password = st.text_input("Gmail App Password", type="password")

    # Step 3: Email Settings
    st.header("3️⃣ Configure Email Settings")
    with st.expander("✉️ CC / BCC / Subject Settings"):
        cc_input = st.text_area("CC Emails", height=50, key="cc_input")
        bcc_input = st.text_area("BCC Emails", height=50, key="bcc_input")
        subject = st.text_input("Email Subject", value="Welcome to Invesmate!")
        delay = st.slider("⏱ Delay between emails (seconds)", 0, 60, 2)

    cc_emails = [e.strip() for line in cc_input.splitlines() for e in line.split(',') if is_valid_email(e.strip())]
    bcc_emails = [e.strip() for line in bcc_input.splitlines() for e in line.split(',') if is_valid_email(e.strip())]

    # Step 4: Compose Email
    st.header("4️⃣ Compose Email Body")
    html_body = st_quill(value="""
    <p><strong>Dear {name},</strong></p>
    <p>Welcome to Invesmate! Your company account has been created.</p>
    <p><strong>Email:</strong> {id}<br><strong>Temporary Password:</strong> {password}</p>
    <p>🔗 Access your account: <a href='https://outlook.office.com/mail/'>Outlook</a></p>
    <p>For any help, contact Admin Support.</p>
    <p>Regards,<br><strong>Invesmate Team</strong></p>
    <img src='https://www.invesmate.com/tracking_open.png?email={email}' width='1' height='1' style='display:none'>
    """, html=True, key="rich_email_body")

    if st.session_state.data is not None:
        st.header("5️⃣ Review, Edit, and Send")
        edited_df = st.data_editor(st.session_state.data, num_rows="dynamic", use_container_width=True,
                                   column_config={"Send": st.column_config.CheckboxColumn(label="Send", default=True)})
        st.session_state.data = edited_df.copy()

        styled_df = edited_df.style.apply(highlight_invalid_cells, axis=1)
        st.dataframe(styled_df, use_container_width=True)

        if st.button("🚀 Send Bulk Emails"):
            if not (sender_email and app_password):
                st.warning("⚠️ Please provide Gmail address and app password.")
            else:
                progress = st.progress(0)
                log_data = []
                success, failure = 0, 0
                to_send_df = st.session_state.data[st.session_state.data['Send'] == True]
                total = len(to_send_df)

                for i, (index, row) in enumerate(to_send_df.iterrows()):
                    recipient = row['Email']
                    name = row['Name']
                    user_id = row['ID']
                    pwd = row['Password']

                    if not (is_valid_email(recipient) and is_valid_email(user_id) and pwd):
                        st.error(f"❌ Skipping {name} ({recipient}): Invalid Email/ID/Password.")
                        failure += 1
                        continue

                    try:
                        msg = MIMEMultipart("alternative")
                        msg['From'] = sender_email
                        msg['To'] = recipient
                        msg['Subject'] = subject
                        if cc_emails:
                            msg['Cc'] = ", ".join(cc_emails)

                        filled_body = html_body.format(name=name, email=recipient, id=user_id, password=pwd)
                        msg.attach(MIMEText(filled_body, 'html'))

                        to_addresses = [recipient] + cc_emails + bcc_emails

                        with smtplib.SMTP("smtp.gmail.com", 587) as server:
                            server.starttls()
                            server.login(sender_email, app_password)
                            server.sendmail(sender_email, to_addresses, msg.as_string())

                        st.success(f"✅ Sent to {name} ({recipient})")
                        success += 1
                        log_data.append([name, recipient, "Success", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

                    except Exception as e:
                        st.error(f"❌ Failed to send to {name}: {str(e)}")
                        failure += 1
                        log_data.append([name, recipient, f"Failed: {str(e)}", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

                    time.sleep(delay)
                    progress.progress((i + 1) / total)

                st.info(f"📢 Summary: {success} Sent | {failure} Failed")

                log_df = pd.DataFrame(log_data, columns=["Name", "Email", "Status", "Timestamp"])
                buffer = BytesIO()
                log_df.to_csv(buffer, index=False)
                st.download_button("📥 Download Log File", data=buffer.getvalue(), file_name="email_log.csv", mime="text/csv")

# Main entry
if __name__ == "__main__":
    run_app()
