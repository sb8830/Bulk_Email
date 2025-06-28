import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Function to validate email
def is_valid_email(email):
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return bool(re.match(pattern, email))

# Function to validate password
def is_strong_password(password):
    pattern = r'^(?=.*[A-Z])(?=.*[a-z])(?=.*\d)(?=.*[@$!%*#?&])[A-Za-z\d@$!#%*?&]{8,}$'
    return bool(re.match(pattern, password))

# Streamlit App
st.title("🔐 Bulk User Uploader & Validator")

# Upload section
st.header("📤 Upload Your User File")
uploaded_file = st.file_uploader("Upload an Excel or CSV file with columns: Name, Email, Password", type=["csv", "xlsx"])

# Demo files below upload
demo_data = {
    "Name": ["John Doe", "Jane Smith", "Rahul Roy"],
    "Email": ["john.doe@example.com", "jane.smith@example.com", "rahul.roy@example.com"],
    "Password": ["Pass@1234", "Welcome#567", "India@2025"]
}
demo_df = pd.DataFrame(demo_data)

# Download demo buttons
st.download_button("⬇️ Download Demo CSV", data=demo_df.to_csv(index=False).encode(), file_name="demo_users.csv", mime="text/csv")
excel_buffer = BytesIO()
demo_df.to_excel(excel_buffer, index=False)
st.download_button("⬇️ Download Demo Excel", data=excel_buffer.getvalue(), file_name="demo_users.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Process uploaded file
if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        required_cols = {"Name", "Email", "Password"}
        if not required_cols.issubset(df.columns):
            st.error("❌ The file must contain 'Name', 'Email', and 'Password' columns.")
        else:
            st.success("✅ File uploaded and contains all required columns!")

            # Validate each row
            results = []
            for idx, row in df.iterrows():
                name, email, password = row['Name'], row['Email'], row['Password']
                email_ok = is_valid_email(str(email))
                pass_ok = is_strong_password(str(password))

                status = "✅ Valid" if email_ok and pass_ok else "❌ Invalid"
                error_msg = ""
                if not email_ok:
                    error_msg += "Invalid Email. "
                if not pass_ok:
                    error_msg += "Weak Password."

                results.append({
                    "Name": name,
                    "Email": email,
                    "Password": password,
                    "Status": status,
                    "Remarks": error_msg.strip()
                })

            results_df = pd.DataFrame(results)
            st.dataframe(results_df)

            # Allow download of the result
            csv_buffer = BytesIO()
            results_df.to_csv(csv_buffer, index=False)
            st.download_button("📥 Download Validation Report", data=csv_buffer.getvalue(), file_name="validation_report.csv", mime="text/csv")

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
