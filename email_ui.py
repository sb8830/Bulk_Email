file = st.file_uploader("Upload an Excel or CSV file", type=["xlsx", "csv"])
if "data" not in st.session_state:
st.session_state.data = None
valid_file = False

if file:
try:
if file.name.endswith(".csv"):
df = pd.read_csv(file)
        elif file.name.endswith(".xlsx"):
        else:
df = pd.read_excel(file)

df.columns = [col.lower().strip() for col in df.columns]

required_columns = {'name', 'sender email', 'email id', 'password'}
if required_columns.issubset(set(df.columns)):
df.rename(columns={
@@ -40,15 +38,13 @@
'email id': 'ID',
'password': 'Password'
}, inplace=True)
            valid_file = True
            st.session_state.data = df.copy()
            df['Send'] = True
            st.session_state.data = df
st.success("✅ File uploaded and validated successfully!")
else:
st.error("❗ Required columns missing: name, sender email, email id, password")
            st.session_state.data = None
except Exception as e:
st.error(f"❌ Failed to read file: {e}")
        st.session_state.data = None

# Step 2: Email credentials
st.header("2️⃣ Setup Email Credentials")
@@ -62,23 +58,23 @@
cc_emails_input = st.text_area("CC Emails (comma/line-separated)", height=70)
bcc_emails_input = st.text_area("BCC Emails (comma/line-separated)", height=70)
subject = st.text_input("Email Subject", value="Welcome to Invesmate!")
    delay = st.slider("⏱ Delay between emails (seconds)", min_value=0, max_value=60, value=2)
    delay = st.slider("⏱ Delay between emails (seconds)", 0, 60, 2)

# Helpers
def is_valid_email(email):
    return re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", str(email))
    return bool(re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", str(email)))

def email_exists(email):
try:
domain = email.split('@')[-1]
        answers = dns.resolver.resolve(domain, 'MX')
        return len(answers) > 0
        dns.resolver.resolve(domain, 'MX')
        return True
except:
return False

# Process CC, BCC
cc_emails = [email.strip() for line in cc_emails_input.splitlines() for email in line.split(',') if email.strip() and is_valid_email(email.strip())]
bcc_emails = [email.strip() for line in bcc_emails_input.splitlines() for email in line.split(',') if email.strip() and is_valid_email(email.strip())]
cc_emails = [e.strip() for l in cc_emails_input.splitlines() for e in l.split(',') if is_valid_email(e.strip())]
bcc_emails = [e.strip() for l in bcc_emails_input.splitlines() for e in l.split(',') if is_valid_email(e.strip())]

# Step 4: Compose email
st.header("4️⃣ Compose Email Body")
@@ -96,50 +92,28 @@ def email_exists(email):
key="rich_email_body"
)

with st.expander("👁 Preview Sample Email"):
    if st.session_state.data is not None:
        preview = html_body.format(name="John Doe", email="john@example.com", id="john@example.com", password="pass123")
        st.markdown(preview, unsafe_allow_html=True)
# Step 5: Preview & Edit
def highlight_invalid(val, col_name):
    if col_name in ['Email', 'ID'] and not is_valid_email(val):
        return 'background-color: #FFD6D6'
    if pd.isna(val) or val == '':
        return 'background-color: #FFD6D6'
    return ''

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

    st.header("5️⃣ Review, Edit, and Send")
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
    styled = edited_df.style.applymap(lambda v: highlight_invalid(v, 'Email'), subset=['Email'])\
                               .applymap(lambda v: highlight_invalid(v, 'ID'), subset=['ID'])
    st.dataframe(styled, use_container_width=True)

    st.session_state.data = edited_df
    st.session_state.data = edited_df.copy()

if st.button("🚀 Send Bulk Emails"):
if not (sender_email and app_password):
@@ -152,46 +126,47 @@ def highlight_invalid(val):
continue

recipient = row['Email']
                name = row.get('Name', 'User')
                user_id = row.get('ID')
                password = row.get('Password')
                name = row['Name']
                user_id = row['ID']
                pwd = row['Password']

                if not is_valid_email(recipient) or not email_exists(recipient) or not is_valid_email(user_id):
                    st.error(f"❌ Skipping {name}: Invalid email or ID.")
                if not (is_valid_email(recipient) and is_valid_email(user_id)):
                    st.error(f"❌ Skipping {name}: Invalid Email/ID format.")
failure += 1
continue

                msg = MIMEMultipart("alternative")
                msg['From'] = sender_email
                msg['To'] = recipient
                msg['Subject'] = subject

                if cc_emails:
                    msg['Cc'] = ", ".join(cc_emails)
                try:
                    msg = MIMEMultipart("alternative")
                    msg['From'] = sender_email
                    msg['To'] = recipient
                    msg['Subject'] = subject
                    if cc_emails:
                        msg['Cc'] = ", ".join(cc_emails)

                filled_body = html_body.format(name=name, email=recipient, id=user_id, password=password)
                msg.attach(MIMEText(filled_body, 'html'))
                    filled_body = html_body.format(name=name, email=recipient, id=user_id, password=pwd)
                    msg.attach(MIMEText(filled_body, 'html'))

                to_addresses = [recipient] + cc_emails + bcc_emails
                    to_addresses = [recipient] + cc_emails + bcc_emails

                try:
with smtplib.SMTP("smtp.gmail.com", 587) as server:
server.starttls()
server.login(sender_email, app_password)
server.sendmail(sender_email, to_addresses, msg.as_string())
                    st.success(f"✅ Sent email to {name} ({recipient})")

                    st.success(f"✅ Sent to {name} ({recipient})")
success += 1
log_data.append([name, recipient, "Success", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

except Exception as e:
                    st.error(f"❌ Failed to send to {name}: {e}")
                    st.error(f"❌ Failed to send to {name}: {str(e)}")
failure += 1
                    log_data.append([name, recipient, f"Failed: {e}", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                    log_data.append([name, recipient, f"Failed: {str(e)}", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

time.sleep(delay)

            st.info(f"📢 Summary: {success} Success ✅ | {failure} Failed ❌")
            st.info(f"📢 Summary: {success} Sent | {failure} Failed")

log_df = pd.DataFrame(log_data, columns=["Name", "Email", "Status", "Timestamp"])
buffer = BytesIO()
log_df.to_csv(buffer, index=False)
            st.download_button("📥 Download Email Log", data=buffer.getvalue(), file_name="email_log.csv", mime="text/csv")
            st.download_button("📥 Download Log File", data=buffer.getvalue(), file_name="email_log.csv", mime="text/csv")
