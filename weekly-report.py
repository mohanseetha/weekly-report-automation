import pymongo
import certifi
import pandas as pd
from datetime import datetime, timedelta
import smtplib
import os
import json
from email.message import EmailMessage
from dotenv import load_dotenv

load_dotenv()
MONGO_URI = os.getenv("MONGO_URI")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT"))
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
ALL_MAIL = os.getenv("ALL_MAIL")
dept_mappings = json.loads(os.getenv("DEPT_MAPPINGS"))

client = pymongo.MongoClient(MONGO_URI, tlsCAFile=certifi.where())
db = client["studentDB"]
collection = db["latecomers"]

today = datetime.today()
monday = today - timedelta(days=today.weekday())
saturday = monday + timedelta(days=5)

data = list(collection.find())
if not data:
    exit(0)

df = pd.DataFrame(data)
try:
    df.drop(['_id', '__v'], axis=1, inplace=True, errors='ignore')
    df['date'] = pd.to_datetime(df['date'])
except:
    exit(1)

df_week = df[(df['date'] >= monday) & (df['date'] <= saturday)]
if df_week.empty:
    exit(0)

df_week['date_str'] = df_week['date'].dt.strftime('%d/%m/%y')
student_counts = (
    df_week.groupby(['pin', 'name', 'department'])
    .agg(
        late_count=('date_str', lambda x: len(set(x))),
        repeated_dates=('date_str', lambda x: ', '.join(sorted(set(x))))
    )
    .reset_index()
)
df_filtered = student_counts[student_counts['late_count'] >= 3]
if df_filtered.empty:
    exit(0)

saved_files = {}
consolidated_filename = f"Weekly_Latecomers_{monday.strftime('%Y-%m-%d')}_to_{saturday.strftime('%Y-%m-%d')}.xlsx"

with pd.ExcelWriter(consolidated_filename, engine="xlsxwriter") as writer:
    for dept, email in dept_mappings.items():
        df_dept = df_filtered[df_filtered['department'] == dept]
        if not df_dept.empty:
            dept_filename = f"{dept}_weekly_latecomers_{monday.strftime('%Y-%m-%d')}.xlsx"
            df_dept.to_excel(writer, sheet_name=dept, index=False)
            df_dept.to_excel(dept_filename, index=False)
            saved_files[dept] = dept_filename
    df_filtered.to_excel(writer, sheet_name="Consolidated", index=False)

def send_email(receiver_email, subject, body, attachment_path):
    if not os.path.exists(attachment_path):
        return
    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver_email
    msg["Subject"] = subject
    msg.set_content(body)
    with open(attachment_path, "rb") as file:
        msg.add_attachment(file.read(), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           filename=os.path.basename(attachment_path))
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
    except:
        pass
    os.remove(attachment_path)

for dept, email in dept_mappings.items():
    if email and dept in saved_files:
        send_email(email, f"Weekly Latecomers Report - {dept} ({monday.strftime('%Y-%m-%d')} to {saturday.strftime('%Y-%m-%d')})",
                   "Attached is the list of students who were late on 3 or more unique days this week.", saved_files[dept])

send_email(ALL_MAIL, f"Weekly Latecomers Consolidated Report ({monday.strftime('%Y-%m-%d')} to {saturday.strftime('%Y-%m-%d')})",
           "Attached is the consolidated latecomers' report for all departments.", consolidated_filename)
