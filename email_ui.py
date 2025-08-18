import streamlit as st
import pandas as pd
import smtplib
import ssl
import time
import uuid
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO

# ======================
# Streamlit Page Config
# ======================
st.set_page_config(page_title="Bulk Email Sender", layout="wide")
st.title("📧 Bulk Email Sender with Personalization & Tracking")

# ======================
# Sidebar - SMTP Settings
# ======================
st.sidebar.header("SMTP Settings")

smtp_server = st.sidebar.text_input("SMTP Server", value="smtp.gmail.com")
smtp_port = st.sidebar.number_input("SMTP Port", value=465, step=1)
sender_email = st.sidebar.text_input("Sender Email")
sender_password = st.sidebar.text_input("Sender Password", type="password")
test_mode = st.sidebar.checkbox("Test Mode (emails not actually sent)", value=False)

# ======================
# Upload Recipient Data
# ======================
st.subheader("1️⃣ Upload Recipient Data")

uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

df = None
if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        st.success("✅ File uploaded successfully!")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")

# ======================
# Email Composition
# ======================
st.subheader("2️⃣ Compose Email")

subject = st.text_input("Email Subject")
body_html = st.text_area("Email Body (HTML supported)", height=250,
                         placeholder="Example: Dear {Name}, your login password is {Password}")

st.caption("🔑 You can use placeholders like `{Name}`, `{Email}`, `{Password}` (must match column names in uploaded file).")

# ======================
# Preview
# ======================
st.subheader("3️⃣ Preview Email for First Recipient")

if df is not None and not df.empty:
    preview_row = df.iloc[0].to_dict()
    preview_body = body_html
    for col in df.columns:
        preview_body = preview_body.replace(f"{{{col}}}", str(preview_row.get(col, "")))
    
    st.markdown(f"**Subject:** {subject}")
    st.markdown("---")
    st.markdown(preview_body, unsafe_allow_html=True)

# ======================
# Send Emails
# ======================
st.subheader("4️⃣ Send Emails")

if st.button("🚀 Start Sending Emails"):
    if df is None:
        st.error("❌ Please upload a file first.")
    elif not sender_email or not sender_password:
        st.error("❌ Please enter SMTP credentials.")
    elif not subject or not body_html:
        st.error("❌ Please compose your email.")
    else:
        context = ssl.create_default_context()
        results = []

        if not test_mode:
            try:
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context)
                server.login(sender_email, sender_password)
            except Exception as e:
                st.error(f"❌ Failed to connect to SMTP server: {e}")
                server = None
        else:
            server = None

        progress = st.progress(0)
        status_box = st.empty()

        for i, row in df.iterrows():
            recipient_email = str(row.get("Email", "")).strip()
            if not recipient_email or "@" not in recipient_email:
                results.append({"Recipient": recipient_email, "Status": "❌ Invalid Email"})
                continue

            personalized_body = body_html
            for col in df.columns:
                personalized_body = personalized_body.replace(f"{{{col}}}", str(row.get(col, "")))
            
            # Add tracking pixel
            tracking_id = str(uuid.uuid4())
            personalized_body += f'<img src="https://via.placeholder.com/1x1.png?uid={tracking_id}" width="1" height="1" />'

            msg = MIMEMultipart("alternative")
            msg["From"] = sender_email
            msg["To"] = recipient_email
            msg["Subject"] = subject
            msg.attach(MIMEText(personalized_body, "html"))

            try:
                if not test_mode and server:
                    server.sendmail(sender_email, recipient_email, msg.as_string())
                results.append({"Recipient": recipient_email, "Status": "✅ Sent"})
            except Exception as e:
                results.append({"Recipient": recipient_email, "Status": f"❌ Failed: {e}"})

            progress.progress((i + 1) / len(df))
            status_box.text(f"Processed {i+1}/{len(df)}")

            time.sleep(1)  # small delay to avoid SMTP rate limits

        if server:
            server.quit()

        # Show results
        st.subheader("📊 Sending Results")
        results_df = pd.DataFrame(results)
        st.dataframe(results_df)

        # Download log
        output = BytesIO()
        results_df.to_csv(output, index=False)
        st.download_button("📥 Download Log as CSV", data=output.getvalue(),
                           file_name="email_results.csv", mime="text/csv")

