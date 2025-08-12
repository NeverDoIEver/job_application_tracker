import win32com.client
import pythoncom
import pandas as pd
import os
from datetime import datetime

EXCEL_FILE = "email_log.xlsx"

# Ensure Excel file exists
if not os.path.exists(EXCEL_FILE):
    pd.DataFrame(columns=["Type", "To", "From", "Subject", "Time", "BodySnippet"]).to_excel(EXCEL_FILE, index=False)

# Event handler for incoming emails
class InboxEventHandler:
    def OnItemAdd(self, item):
        try:
            df = pd.read_excel(EXCEL_FILE)
            df.loc[len(df)] = {
                "Type": "Received",
                "To": item.To,
                "CC": item.CC,
                "Subject": item.Subject,
                "Sent On": item.SentOn.strftime("%d.%m.%Y"),
                "Body": item.Body[:200]  # Truncate long body
            }
            df.to_excel(EXCEL_FILE, index=False)
            print(f"[INBOX] New email from {item.SenderEmailAddress} logged.")
        except Exception as e:
            print(f"[INBOX] Error logging email: {e}")

# Event handler for sent emails
class SentEventHandler:
    def OnItemAdd(self, item):
        try:
            df = pd.read_excel(EXCEL_FILE)
            df.loc[len(df)] = {
                "Type": "Sent",
                "To": item.To,
                "CC": item.CC,
                "Subject": item.Subject,
                "Sent On": item.SentOn.strftime("%d.%m.%Y"),
                "Body": item.Body[:200]  # Truncate long body
            }
            df.to_excel(EXCEL_FILE, index=False)
            print(f"[SENT] New sent email to {item.To} logged.")
        except Exception as e:
            print(f"[SENT] Error logging sent email: {e}")

# ✅ Correct way to connect to Outlook
outlook_app = win32com.client.Dispatch("Outlook.Application")
namespace = outlook_app.GetNamespace("MAPI")

# Get Inbox and Sent Items
inbox_items = namespace.GetDefaultFolder(6).Items  # 6 = Inbox
sent_items = namespace.GetDefaultFolder(5).Items   # 5 = Sent Items

# Attach event handlers
win32com.client.WithEvents(inbox_items, InboxEventHandler)
win32com.client.WithEvents(sent_items, SentEventHandler)

print(" Tracking incoming and outgoing emails...\nPress CTRL+C to stop.")

# Keep script running
while True:
    pythoncom.PumpWaitingMessages()
