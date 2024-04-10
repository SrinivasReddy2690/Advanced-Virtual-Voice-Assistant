import tkinter as tk
import smtplib
from dotenv import load_dotenv
import os
import re

load_dotenv(dotenv_path='..\\Data\\.env')

sender_id = os.getenv('EMAIL_ID')
password = os.getenv('PASSWORD')

def send_email(reciever_id, subject, body):

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(sender_id, password)
    message = "\r\n".join([
        f"From: {sender_id}",
        f"To: {reciever_id}",
        f"Subject: {subject}",
        "",
        f"{body}"
    ])
    try:
        s.sendmail(sender_id, reciever_id, message)
    except smtplib.SMTPRecipientsRefused:
        print("INVALID EMAIL ADDRESS")
        return False
    except smtplib.SMTPException:
        return False
    s.quit()
    return True

def check_email(email):
    verifier = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    if re.fullmatch(verifier, email):
        return True
    else:
        return False
    