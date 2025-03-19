import smtplib
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Function to Select CSV File
def select_csv_file():
    root = tk.Tk()
    root.withdraw()  # Hide root window
    file_path = filedialog.askopenfilename(title="Select Recipients CSV", filetypes=[("CSV Files", "*.csv")])
    return file_path

# Get User Inputs
SENDER_EMAIL = simpledialog.askstring("Input", "Enter your sender email:")
SENDER_PASSWORD = simpledialog.askstring("Input", "Enter your email password (Use App Password if needed):", show="*")
EMAIL_SUBJECT = simpledialog.askstring("Input", "Enter email subject:")

# Select CSV File
csv_file = select_csv_file()
if not csv_file:
    print("No CSV file selected. Exiting...")
    exit()

# Load Recipients from CSV
try:
    df = pd.read_csv(csv_file)
    if "email" not in df.columns or "name" not in df.columns:
        print("CSV file must contain 'email' and 'name' columns. Exiting...")
        exit()
except Exception as e:
    print("Error reading CSV file:", e)
    exit()

# SMTP Configuration
SMTP_SERVER = "smtp.gmail.com"  # Change for Outlook/Yahoo/etc.
SMTP_PORT = 587

EMAIL_BODY = """
Dear {name},

This is an automated email sent using Python.
We appreciate your time and attention.

Best Regards,
Your Name
"""

# Connect to SMTP Server
try:
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()  # Secure Connection
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    print("\nLogged in successfully.")

    # Sending Emails
    for index, row in df.iterrows():
        recipient_email = row["email"]
        recipient_name = row["name"]
        
        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = recipient_email
        msg["Subject"] = EMAIL_SUBJECT
        
        # Personalizing Email
        personalized_body = EMAIL_BODY.format(name=recipient_name)
        msg.attach(MIMEText(personalized_body, "plain"))

        # Send Email
        server.sendmail(SENDER_EMAIL, recipient_email, msg.as_string())
        print(f"✅ Email sent to {recipient_email}")

    server.quit()
    print("\n✅✅✅ All emails sent successfully! ✅✅✅")

except Exception as e:
    print("❌ Error:", e)
