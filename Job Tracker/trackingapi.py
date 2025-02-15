from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pandas as pd
import os
import re
import time
import logging
import json
from google.generativeai import configure, GenerativeModel
import base64
from bs4 import BeautifulSoup
import google.api_core.exceptions

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Configure Gemini API
configure(api_key=os.getenv('GEMINI_API_KEY'))
model = GenerativeModel('gemini-2.0-flash')

SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
KEYWORDS = {
    'Rejected': ['decision to not move forward', 'unable to give you further consideration','not be moving forward','you have not been selected','not to move forward'],
    'In Progress': ['you are selected', 'moving forward', 'application is under'],
    'Applied': ['thank you for applying', 'your application has been received','successfully received your application','successfully submitted','your application was sent','Application submitted','received your application']
}

def generate_with_retry(prompt):
    for attempt in range(5):
        try:
            result = model.generate_content(prompt).text.strip()
            return result.split(',')[0].strip() 
        except google.api_core.exceptions.ResourceExhausted:
            logging.warning("Quota reached. Retrying in 15 seconds...")
            time.sleep(15)
    raise Exception("Quota limit reached. Try later.")

def extract_body(message):
    body = ''
    if 'parts' in message['payload']:
        for part in message['payload']['parts']:
            if part.get('parts'):
                body += extract_body({'payload': part})
            elif part['mimeType'] in ['text/plain', 'text/html']:
                data = part.get('body', {}).get('data')
                if data:
                    body += base64.urlsafe_b64decode(data).decode('utf-8')
    else:
        data = message['payload'].get('body', {}).get('data')
        if data:
            body = base64.urlsafe_b64decode(data).decode('utf-8')
    return BeautifulSoup(body, 'html.parser').get_text(separator='\n', strip=True)

def categorize_email(body):
    body_lower = body.lower()
    if any(keyword in body_lower for keyword in KEYWORDS['Rejected']):
        return 'Rejected'
    elif any(keyword in body_lower for keyword in KEYWORDS['In Progress']):
        return 'In Progress'
    elif any(keyword in body_lower for keyword in KEYWORDS['Applied']):
        return 'Applied'
    return 'Unknown'

def update_database(data):
    file_path = 'JobTracker.xlsx'
    df = pd.DataFrame([data])
    if os.path.exists(file_path):
        existing_df = pd.read_excel(file_path)
        combined_df = pd.concat([existing_df, df], ignore_index=True)
        combined_df.to_excel(file_path, index=False)
    else:
        df.to_excel(file_path, index=False)

def apply_label(service, msg_id, label_name):
    labels = service.users().labels().list(userId='me').execute().get('labels', [])
    label_id = next((label['id'] for label in labels if label['name'].lower() == label_name.lower()), None)
    if label_id:
        service.users().messages().modify(userId='me', id=msg_id, body={'addLabelIds': [label_id]}).execute()

def extract_all_emails(service, messages):
    email_data = []
    for msg in messages:
        message = service.users().messages().get(userId='me', id=msg['id'], format='full').execute()
        body = extract_body(message)
        subject = next((h['value'] for h in message['payload']['headers'] if h['name'] == 'Subject'), 'N/A')
        sender = next((h['value'] for h in message['payload']['headers'] if h['name'] == 'From'), 'N/A')
        job_title = generate_with_retry(f"Extract only the job title: '{subject} {body}'")
        company_name = generate_with_retry(f"Extract only the company name: '{subject} {body}'")
        status = categorize_email(body)

        if status != 'Unknown':  # Only add if status is not 'Unknown'
            entry = {
                'STATUS': status,
                'JOB TITLE': job_title,
                'COMPANY': company_name,
                'PLATFORM': sender.split('<')[0].strip(),
                'EMAIL ID': sender.split('<')[-1].strip('>'),
            }
            email_data.append(entry)
            update_database(entry)
            apply_label(service, msg['id'], status)

    with open('email_data.json', 'w') as json_file:
        json.dump(email_data, json_file, indent=4)
    logging.info(f"Saved {len(email_data)} valid emails to email_data.json")

def authenticate_gmail():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(os.getenv('GMAIL_CREDENTIALS_PATH'), SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def main():
    logging.info("Starting Job Application Tracker System")
    service = authenticate_gmail()
    messages = service.users().messages().list(userId='me', maxResults=100).execute().get('messages', [])
    extract_all_emails(service, messages)
    logging.info("Job Application Tracker System completed")

if __name__ == '__main__':
    main()
