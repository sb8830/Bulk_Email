import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import re
from st_aggrid import AgGrid
from st_quill import st_quill

# --------------- CONFIGURE EMAIL CREDENTIALS ---------------- #
SENDER_EMAIL = "YOUR_EMAIL@gmail.com"
SENDER_PASSWORD = "YOUR_APP_PASSWORD"  # Use App Password from Google
USE_SSL = False  # Set to True to use SSL on port 465

# --------------- Email Validation ---------------- #
def is_valid_email(email):
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email)

# --------------- Send Email Function ---------------- #
def send_email(to_email, subject, body):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))

        if USE_SSL:
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        else:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()

        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        server.quit()
        return 'Success ✅'
    except Exception as e:
        print(f"[Error] {to_email}: {e}")
        return f'❌ {str(e)}'

# --------------- Streamlit UI ---------------- #
st.set_page_config(page_title="Bulk Email Sender", layout="centered")
st.title("📧 Bulk Email Sender")

# ---------- File Upload ----------- #
st.sidebar.header("Upload Excel or CSV")
uploaded_file = st.sidebar.file_uploader("Upload File", type=["xlsx", "csv"])

# ---------- File Processing ----------- #
if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.success("✅ File uploaded successfully!")
        st.subheader("📄 Preview of Uploaded Data")
        AgGrid(df)

        if 'Email' not in df.columns:
            st.error("❌ Please make sure your file includes an 'Email' column.")
        else:
            # Compose Email
            st.subheader("✍️ Compose Your Email")
            subject = st.text_input("Subject", max_chars=200)
            body = st_quill(label="Email Body (Rich HTML Supported)")

            # Send Emails
            if st.button("📤 Send Emails Now"):
                if not subject or not body:
                    st.warning("⚠️ Subject and Email Body are required.")
                else:
                    status_list = []
                    for idx, row in df.iterrows():
                        email = str(row['Email']).strip()
                        if pd.isna(email) or not is_valid_email(email):
                            status = "Invalid Email ❌"
                        else:
                            status = send_email(email, subject, body)
                        status_list.append(status)

                    df['Status'] = status_list
                    st.success("✅ All emails processed!")
                    st.dataframe(df[['Email', 'Status']])
                    st.download_button("📥 Download Log", df.to_csv(index=False), "email_log.csv", "text/csv")

    except Exception as e:
        st.error(f"❌ File processing error: {e}")
else:
    st.info("⬆️ Upload your Excel/CSV file to begin.")
