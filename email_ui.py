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

# Set this at the very top
st.set_page_config(page_title="📧 Bulk Email Sender", layout="wide")

# Dummy user database
USER_DB = {
    "admin": {"password": "admin123", "role": "admin"},
    "user": {"password": "user123", "role": "user"}
}

# Helper functions
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

def login_page():
    st.title("🔐 Login to Bulk Email Sender")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        user = USER_DB.get(username)
        if user and password == user["password"]:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.user_role = user["role"]
            st.experimental_rerun()
            return
        else:
            st.error("Invalid username or password")

def logout():
    if st.sidebar.button("Logout"):
        st.session_state.clear()
        st.experimental_rerun()
        return

def run_app():
    st.sidebar.image("data:image/png;base64,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", width=200)
    st.sidebar.markdown("Developed by Invesmate Admin Team")
    logout()

    role = st.session_state.user_role
    st.title("📧 Bulk Email Sender")

    # Step 1: Upload file
    st.header("1️⃣ Upload Recipient Data")
    file = st.file_uploader("Upload an Excel or CSV file", type=["xlsx", "csv"])
    if "data" not in st.session_state:
        st.session_state.data = None

    if file:
        try:
            df = pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)
            df.columns = [col.lower().strip() for col in df.columns]
            required_columns = {'name', 'sender email', 'email id', 'password'}
            if required_columns.issubset(set(df.columns)):
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
                st.error("❗ Required columns missing: name, sender email, email id, password")
        except Exception as e:
            st.error(f"❌ Failed to read file: {e}")

    # Step 2: Email credentials
    st.header("2️⃣ Setup Email Credentials")
    with st.expander("🔒 Gmail Login"):
        sender_email = st.text_input("Gmail Address", placeholder="you@gmail.com")
        app_password = st.text_input("Gmail App Password", type="password")

    # Step 3: Email Settings
    st.header("3️⃣ Configure Email Settings")
    with st.expander("✉️ CC / BCC / Subject Settings"):
        cc_emails_input = st.text_area("CC Emails", height=50)
        bcc_emails_input = st.text_area("BCC Emails", height=50)
        subject = st.text_input("Email Subject", value="Welcome to Invesmate!")
        delay = st.slider("⏱ Delay between emails (seconds)", 0, 60, 2)

    cc_emails = [e.strip() for l in cc_emails_input.splitlines() for e in l.split(',') if is_valid_email(e.strip())]
    bcc_emails = [e.strip() for l in bcc_emails_input.splitlines() for e in l.split(',') if is_valid_email(e.strip())]

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

    # Step 5: Preview, Edit, and Send
    if st.session_state.data is not None:
        st.header("5️⃣ Review, Edit, and Send")
        edited_df = st.data_editor(
            st.session_state.data,
            num_rows="dynamic",
            use_container_width=True,
            column_config={"Send": st.column_config.CheckboxColumn(label="Send", default=True)}
        )
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

# Entry point
if __name__ == "__main__":
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        login_page()
    else:
        run_app()
