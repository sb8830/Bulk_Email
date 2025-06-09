import pandas as pd
import smtplib
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
from streamlit_quill import st_quill
from io import BytesIO
from datetime import datetime
import time

# 1. Page config - MUST be at the very top before any widgets
st.set_page_config(page_title="📧 Bulk Email Sender with Login", layout="wide")

# --- Login and Logout Functions ---
def login():
    st.title("🔐 Login Panel")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        # Example hardcoded credentials
        if username == "admin" and password == "admin123":
            st.session_state["logged_in"] = True
            st.session_state["user_role"] = "admin"
            st.experimental_rerun()
            return
        elif username == "user" and password == "user123":
            st.session_state["logged_in"] = True
            st.session_state["user_role"] = "user"
            st.experimental_rerun()
            return
        else:
            st.error("❌ Invalid username or password")

def logout():
    if st.button("Logout"):
        for key in ["logged_in", "user_role"]:
            if key in st.session_state:
                del st.session_state[key]
        st.experimental_rerun()
        return

# --- Email validation helper ---
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

# --- Main Bulk Email Sender UI (for admin only) ---
def bulk_email_sender():
    st.title("📧 Bulk Email Sender - Admin Panel")
    st.sidebar.image("https://www.invesmate.com/assets/images/logo.png", width=200)
    st.sidebar.markdown("Developed by Invesmate Admin Team")

    # Step 1: Upload file
    st.header("1️⃣ Upload Recipient Data")
    file = st.file_uploader("Upload an Excel or CSV file", type=["xlsx", "csv"])
    if "data" not in st.session_state:
        st.session_state.data = None

    if file:
        try:
            if file.name.endswith(".csv"):
                df = pd.read_csv(file)
            else:
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
        cc_emails_input = st.text_area("CC Emails (comma/line-separated)", height=70)
        bcc_emails_input = st.text_area("BCC Emails (comma/line-separated)", height=70)
        subject = st.text_input("Email Subject", value="Welcome to Invesmate!")
        delay = st.slider("⏱ Delay between emails (seconds)", 0, 60, 2)

    # Process CC, BCC emails
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

# --- Main User Panel ---
def user_panel():
    st.title("👤 User Panel")
    st.write("Welcome normal user! Limited access here.")
    st.write("You can customize this panel as per your needs.")

# --- Main app ---
def run_app():
    if "logged_in" not in st.session_state or not st.session_state["logged_in"]:
        login()
    else:
        # Show Logout button
        logout()

        # Show different UI based on role
        if st.session_state.get("user_role") == "admin":
            bulk_email_sender()
        else:
            user_panel()

if __name__ == "__main__":
    run_app()
