import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st

st.title("📧 Bulk Email Sender")

excel_file = st.file_uploader("Upload Excel file", type=["xlsx"])
sender_email = st.text_input("Your Gmail")
app_password = st.text_input("App Password", type="password")

if st.button("Send Emails"):
    if not (excel_file and sender_email and app_password):
        st.warning("Please fill all the fields.")
    else:
        try:
            df = pd.read_excel(excel_file)
            for index, row in df.iterrows():
                recipient = row['Email']
                name = row.get('Name', 'Customer')

                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient
                msg['Subject'] = "Your Subject Here"

                body = f"""Dear {name},

This is a test email sent using Python.

Regards,
Your Name"""
                msg.attach(MIMEText(body, 'plain'))

                with smtplib.SMTP("smtp.gmail.com", 587) as server:
                    server.starttls()
                    server.login(sender_email, app_password)
                    server.send_message(msg)

            st.success("Emails sent successfully!")
        except Exception as e:
            st.error(f"Error: {e}")

