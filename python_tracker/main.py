import win32com.client
import re
import os
import json
import datetime
import schedule
import time
import nltk
from nltk import word_tokenize

nltk.download('punkt_tab')

def simple_tokenizer(whole_article):
    token_list = word_tokenize(whole_article)
    return " ".join(token_list)


sent_log_file = 'log_file_sent.json'
receive_log_file = 'log_file_received.json'

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")


def get_latest_email(email_type):
    print(email_type[:-4])
    if "sent" == email_type[:-4]:
        folder = outlook.GetDefaultFolder(5).Items
        folder.Sort("[ReceivedTime]", True)
        return folder.GetFirst()
    else:
        mailbox = outlook.Folders["davide.ferrara@ost.ch"]
        folder = mailbox.Folders["neu_Bewerbungen"].Items
        folder.Sort("[ReceivedTime]", True)
        return folder.GetFirst()


def load_log(filename):
    if os.path.exists(filename):
        with open(filename, "r", encoding="utf-8") as log:
            try:
                return json.load(log)
            except json.JSONDecodeError:
                return []
    return []


def save_log(filename, data):
    existing_data = load_log(filename)
    existing_data.extend(data)
    with open(filename, "w", encoding="utf-8") as log_file:
        json.dump(existing_data, log_file, ensure_ascii=False, indent=2)


def clean_text(raw_text):
    body_clean1 = re.sub(r'\[cid:.*?\\]', '', raw_text, flags=re.IGNORECASE)
    body_clean2 = re.sub(r'(?:[A-Za-z0-9+/]{80,}={0,2}\s*)+', '', body_clean1)
    body_clean3 = re.sub(r'http[s]?://\S+\.(?:png|jpg|jpeg|gif)\S*', '', body_clean2, flags=re.IGNORECASE)
    body_clean4 = re.sub(r'\s+', ' ', body_clean3).strip()
    return body_clean4.strip()


def extract_new_emails(email_type):
    new_data = []
    latest_email = get_latest_email(email_type)
    try:
        body_text = clean_text(latest_email.Body)
        print(body_text)
        new_data.append({
            "To": latest_email.To,
            "CC": latest_email.CC,
            "Subject": latest_email.Subject,
            "SentOn": latest_email.SentOn.strftime("%d-%m-%Y %H:%M:%S"),
            "Body": body_text[:1000]
            })
    except Exception as e:
        print(f"Skipped email due to error: {e}")
    return new_data


def get_last_date(log_file):
    log_data = load_log(log_file)
    if not log_data:
        return
    last_entry = log_data[-1]['SentOn']
    try:
        return last_entry
    except (KeyError, ValueError):
        return None


def add_new_email(log_file_name):
    if not os.path.exists(log_file_name):
        latest = extract_new_emails(log_file_name)
        save_log(log_file_name, latest)
        print("saved")
    else:
        current_email_date = get_last_date(log_file_name)
        new_mail = extract_new_emails(log_file_name)
        new_email_date = new_mail[0]['SentOn']
        if new_email_date > current_email_date:
            save_log(log_file_name, new_mail)


def track_new_emails():
    print(f"Checking for new emails at {datetime.datetime.now()}")

    add_new_email(sent_log_file)
    add_new_email(receive_log_file)


track_new_emails()

schedule.every(1).minutes.do(track_new_emails)

print("Email tracker running. Press CTRL+C to stop.")
while True:
    schedule.run_pending()
    time.sleep(1)
