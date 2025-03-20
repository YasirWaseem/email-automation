import smtplib
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime
import json
import os

# ---------------------------- LOGIC: Email Sending and CSV Handling ---------------------------- #

class EmailSender:
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    DAILY_LIMIT = 500
    DATA_FILE = "email_limit.json"

    def __init__(self):
        self.sender_email = ""
        self.sender_password = ""
        self.email_subject = ""
        self.email_body = ""
        self.csv_file = None
        self.df = None
        self.sent_today = self.load_email_count()

    def set_credentials(self, email, password):
        self.sender_email = email
        self.sender_password = password

    def load_csv(self, file_path):
        try:
            df = pd.read_csv(file_path)
            if "email" not in df.columns or "name" not in df.columns:
                return None, "CSV must contain 'email' and 'name' columns."
            self.df = df
            self.csv_file = file_path
            return df, None
        except Exception as e:
            return None, str(e)

    def set_email_content(self, subject, body):
        self.email_subject = subject
        self.email_body = body

    def load_email_count(self):
        """Loads the email count from a file to track the daily limit."""
        if os.path.exists(self.DATA_FILE):
            try:
                with open(self.DATA_FILE, "r") as file:
                    data = json.load(file)
                    saved_date = data.get("date", "")
                    if saved_date == str(datetime.date.today()):
                        return data.get("count", 0)
            except Exception as e:
                print(f"Error loading email count: {e}")

        return 0  # Reset if it's a new day or error occurs

    def save_email_count(self, count):
        """Saves the current email count to a file."""
        data = {"date": str(datetime.date.today()), "count": count}
        try:
            with open(self.DATA_FILE, "w") as file:
                json.dump(data, file)
        except Exception as e:
            print(f"Error saving email count: {e}")

    def send_emails(self, progress_callback):
        if self.df is not None and not self.df.empty:
            if self.sent_today >= self.DAILY_LIMIT:
                return None, f"Daily email limit of {self.DAILY_LIMIT} reached! Try again tomorrow."

            try:
                server = smtplib.SMTP(self.SMTP_SERVER, self.SMTP_PORT)
                server.starttls()
                server.login(self.sender_email, self.sender_password)

                emails_sent = 0
                for index, row in self.df.iterrows():
                    if self.sent_today + emails_sent >= self.DAILY_LIMIT:
                        progress_callback(f"⚠️ Limit reached! Sent {emails_sent} emails today.")
                        break

                    recipient_email = row["email"]
                    recipient_name = row["name"]
                    personalized_body = self.email_body.replace("{name}", recipient_name)

                    msg = MIMEMultipart()
                    msg["From"] = self.sender_email
                    msg["To"] = recipient_email
                    msg["Subject"] = self.email_subject
                    msg.attach(MIMEText(personalized_body, "plain"))

                    server.sendmail(self.sender_email, recipient_email, msg.as_string())
                    emails_sent += 1
                    progress_callback(f"✅ Sent to {recipient_email}")

                server.quit()

                self.sent_today += emails_sent
                self.save_email_count(self.sent_today)

                return f"Sent {emails_sent} emails successfully! Remaining limit: {self.DAILY_LIMIT - self.sent_today}", None
            except Exception as e:
                return None, str(e)
        else:
            return None, "CSV file is empty or not loaded."

# ---------------------------- UI CLASSES ---------------------------- #

class EmailApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Email Automation App")
        self.geometry("500x500")
        self.email_sender = EmailSender()

        self.container = tk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self.frames = {}
        for F in (LoginPage, CSVPage, EmailPage, SendEmailPage):
            frame = F(self.container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(LoginPage)

    def show_frame(self, frame_class):
        frame = self.frames[frame_class]
        frame.tkraise()

class LoginPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        tk.Label(self, text="Enter Your Email", font=("Arial", 14)).pack(pady=10)

        self.email_entry = ttk.Entry(self, width=40)
        self.email_entry.pack(pady=5)

        tk.Label(self, text="Enter Your Password", font=("Arial", 14)).pack(pady=10)
        self.password_entry = ttk.Entry(self, width=40, show="*")
        self.password_entry.pack(pady=5)

        self.login_btn = ttk.Button(self, text="Next", command=self.save_credentials)
        self.login_btn.pack(pady=20)

    def save_credentials(self):
        email = self.email_entry.get()
        password = self.password_entry.get()
        if email and password:
            self.controller.email_sender.set_credentials(email, password)
            self.controller.show_frame(CSVPage)
        else:
            messagebox.showerror("Error", "Both fields are required!")

class CSVPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        tk.Label(self, text="Select CSV File", font=("Arial", 14)).pack(pady=20)

        self.select_btn = ttk.Button(self, text="Browse", command=self.select_csv)
        self.select_btn.pack(pady=10)

        self.csv_label = tk.Label(self, text="No file selected", fg="red")
        self.csv_label.pack()

        self.next_btn = ttk.Button(self, text="Next", command=self.validate_csv)
        self.next_btn.pack(pady=20)

    def select_csv(self):
        file_path = filedialog.askopenfilename(title="Select CSV", filetypes=[("CSV Files", "*.csv")])
        if file_path:
            self.csv_label.config(text=f"Selected: {file_path}", fg="green")

    def validate_csv(self):
        file_path = self.csv_label.cget("text").replace("Selected: ", "")
        df, error = self.controller.email_sender.load_csv(file_path)
        if error:
            messagebox.showerror("Error", error)
        else:
            self.controller.show_frame(EmailPage)

class EmailPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        tk.Label(self, text="Email Subject:", font=("Arial", 14)).pack(pady=10)

        self.subject_entry = ttk.Entry(self, width=50)
        self.subject_entry.pack(pady=5)

        tk.Label(self, text="Email Body:", font=("Arial", 14)).pack(pady=10)

        self.body_text = scrolledtext.ScrolledText(self, width=60, height=10)
        self.body_text.pack(pady=5)

        self.next_btn = ttk.Button(self, text="Next", command=self.save_email_content)
        self.next_btn.pack(pady=20)

    def save_email_content(self):
        subject = self.subject_entry.get()
        body = self.body_text.get("1.0", tk.END).strip()
        if subject and body:
            self.controller.email_sender.set_email_content(subject, body)
            self.controller.show_frame(SendEmailPage)
        else:
            messagebox.showerror("Error", "Subject and body cannot be empty!")

class SendEmailPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.send_button = ttk.Button(self, text="Send Emails", command=self.send_emails)
        self.send_button.pack(pady=10)
        
        # Scrolled text widget for progress messages
        self.progress_text = scrolledtext.ScrolledText(self, width=60, height=10, state="disabled")
        self.progress_text.pack(pady=10)

    def update_progress(self, message):
        self.progress_text.configure(state="normal")
        self.progress_text.insert(tk.END, message + "\n")
        self.progress_text.see(tk.END)
        self.progress_text.configure(state="disabled")

    def send_emails(self):
        # Disable the send button to prevent multiple clicks
        self.send_button.config(state="disabled")
        # Clear previous messages
        self.progress_text.configure(state="normal")
        self.progress_text.delete("1.0", tk.END)
        self.progress_text.configure(state="disabled")
        
        # Call send_emails with the update_progress callback
        success, error = self.controller.email_sender.send_emails(self.update_progress)
        if error:
            messagebox.showerror("Error", error)
        else:
            messagebox.showinfo("Success", success)
        # Re-enable the button after sending emails
        self.send_button.config(state="normal")

if __name__ == "__main__":
    app = EmailApp()
    app.mainloop()
