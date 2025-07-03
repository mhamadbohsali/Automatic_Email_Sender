# ===============================
# LOCAL TEST MODE - EMAIL SENDER TOOL v2 (Outlook HTML + Inline Image)
# ===============================
# Requirements:
#   pip install pyodbc schedule pywin32
#   Outlook must be installed
# ===============================

import pyodbc
import schedule
import time
import win32com.client

# -------------------------------
# STEP 1: Connect to SQL Server and fetch users
# -------------------------------
def get_users():
    conn_str = (
        "DRIVER={ODBC Driver 18 for SQL Server};"
        "SERVER= *Your Local Server IP* ;"
        "DATABASE= *Database Name*;"
        "UID= *Database Username*;"
        "PWD= *Database Password*;"
        "TrustServerCertificate=yes;"
        "Encrypt=yes;"
    )

    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT FullName, Email 
            FROM dbo.ADUsers 
            WHERE Country = 'XXXX' 
              AND Email = 'xxx.xxx@xxx.com'
        """)
        rows = cursor.fetchall()
        conn.close()
        return [{'FullName': row[0], 'Email': row[1]} for row in rows]

    except Exception as e:
        print(f"SQL Connection Error: {e}")
        return []

# -------------------------------
# STEP 2: Send personalized HTML email with inline image via Outlook
# -------------------------------
def send_html_email_with_image(to_email, full_name):
    try:
        # Setup Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        # Subject and recipient
        mail.Subject = "Value Awards Nomination Reminder"
        mail.To = to_email

        # Path to header image
        image_path = "C:/script/1.png"
        image_cid = "header001"

        # Build the personalized HTML body
        html_body = f"""
        <html>
        <body>
            <p><img src="cid:{image_cid}" alt="Header Image"></p>

            <p>Dear {full_name},</p>

            <p>Just a reminder to nominate people for the <b>Value Awards</b>.</p>

            <p>If you think your nominee demonstrates the criteria of our values and that their contribution stands out, 
            please consider nominating them for our Value Awards.</p>

            <p>Please click on the link below and nominate your colleague you feel is deserving of any of these awards!</p>

            <p><a href="urlLink">
            urlLink</a></p>

            <p>Good Luck!<br>
            <b>Reward & Recognition</b></p>
        </body>
        </html>
        """

        mail.HTMLBody = html_body

        # Embed the image
        attachment = mail.Attachments.Add(image_path)
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", image_cid
        )

        # Send
        mail.Send()
        print(f"Sent to {to_email}")

    except Exception as e:
        print(f"Failed to send to {to_email}: {e}")

# -------------------------------
# STEP 3: Scheduled Job
# -------------------------------
def job():
    print("Running email job...")
    users = get_users()
    for user in users:
        send_html_email_with_image(user['Email'], user['FullName'])

# -------------------------------
# STEP 4: Run scheduler
# -------------------------------
#schedule.every(1).minutes.do(job)

#print("LOCAL TEST MODE: Scheduler is running...")
#while True:
   # schedule.run_pending()
   # time.sleep(1)
# ===============================