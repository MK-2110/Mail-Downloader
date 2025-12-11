from __future__ import print_function
import os
import base64
import pickle
import time
from pathlib import Path

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
SENDER_EMAIL = "noreply@icegate.gov.in"
OUTPUT_FOLDER = "downloaded_final_leo"
CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.pkl"


os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def authenticate_gmail():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES
            )
            creds = flow.run_local_server(port=8080)

        with open(TOKEN_FILE, "wb") as token:
            pickle.dump(creds, token)

    return build("gmail", "v1", credentials=creds)


def get_subject(payload):
    for h in payload.get("headers", []):
        if h.get("name", "").lower() == "subject":
            return h.get("value", "")
    return ""


def safe_get_attachment(service, msg_id, att_id, retries=3):
    for i in range(retries):
        try:
            return service.users().messages().attachments().get(
                userId="me", messageId=msg_id, id=att_id
            ).execute()
        except Exception as e:
            print(f" Retry {i+1}: {e}")
            time.sleep(1)
    return None


def download_attachments_from_message(service, msg_id):
    try:
        message = service.users().messages().get(
            userId="me", id=msg_id, format="full"
        ).execute()
    except HttpError:
        return 0

    payload = message.get("payload", {})
    subject = get_subject(payload)


    if "final" not in subject.lower() or "leo" not in subject.lower():
        print(" Skipped (Not Final LEO mail):", subject)
        return 0

    print(" Final LEO Mail Found:", subject)

    parts = []

    def collect_parts(p):
        if not p:
            return
        if p.get("parts"):
            for sp in p.get("parts"):
                collect_parts(sp)
        else:
            parts.append(p)

    collect_parts(payload)

    count = 0
    for part in parts:
        filename = part.get("filename")
        if not filename or not filename.lower().endswith(".pdf"):
            continue

        body = part.get("body", {})
        att_id = body.get("attachmentId")

        if not att_id:
            continue

        attachment = safe_get_attachment(service, msg_id, att_id)
        if not attachment:
            continue

        data = attachment.get("data")
        if not data:
            continue

        file_data = base64.urlsafe_b64decode(data.encode("UTF-8"))

        save_path = Path(OUTPUT_FOLDER) / filename
        with open(save_path, "wb") as f:
            f.write(file_data)

        print(" Final LEO PDF Downloaded:", save_path)
        count += 1

    return count


def main():
    service = authenticate_gmail()

    START_DATE = "2021/01/01"
    END_DATE = "2025/12/31"

    query = f'from:{SENDER_EMAIL} subject:"Final LEO" has:attachment filename:pdf after:{START_DATE} before:{END_DATE}'

    print(" Gmail query:\n", query)

    total = 0
    nextPageToken = None

    while True:
        response = service.users().messages().list(
            userId="me", q=query, maxResults=100, pageToken=nextPageToken
        ).execute()

        messages = response.get("messages", [])
        print(f" Found {len(messages)} Final LEO mails")

        for m in messages:
            total += download_attachments_from_message(service, m["id"])

        nextPageToken = response.get("nextPageToken")
        if not nextPageToken:
            break

    print(f"\n COMPLETED — Total FINAL LEO PDFs saved: {total}")


if __name__ == "__main__":
    main()
