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

# Page settings
st.set_page_config(page_title="📧 Bulk Email Sender", layout="wide")
st.title("📧 Bulk Email Sender")
st.sidebar.image("https://www.invesmate.com/assets/images/logo.png", width=200)
st.sidebar.markdown("Developed by Invesmate Admin Team")

# Step 1: Upload or Create file
st.header("1️⃣ Upload or Create Recipient Data")
file = st.file_uploader("Upload an Excel or CSV file", type=["xlsx", "csv"])
if "data" not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=["Name", "Email", "ID", "Password", "Send"])

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
            st.success("✅ File
