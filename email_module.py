import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os

def authenticate_gmail_api():
    # Load in Gmail API credentials
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')

    # Have the user log in if the credentials aren't available or are no longer valid
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', ['https://www.googleapis.com/auth/gmail.compose'])
            creds = flow.run_local_server(port=0)

        # Save credentials so you don't have to log in every single time
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds

def create_and_draft_email(to_email, subject, body, signature, attachment_path, cc_emails, vendor):
    if os.path.exists(f'{attachment_path}'):
        # Authenticate with the API
        creds = authenticate_gmail_api()
        service = build('gmail','v1', credentials=creds)

        message = MIMEMultipart()
        message['to'] = ', '.join(to_email)
        message['subject'] = subject
        message['Cc'] = ', '.join(cc_emails)

        body_text_part = MIMEText(body, 'plain')
        message.attach(body_text_part)

        body_html_part = MIMEText(signature, 'html')
        message.attach(body_html_part)

        # Attach the file
        with open(attachment_path, 'rb') as file:
            attach = MIMEApplication(file.read(), Name=os.path.basename(attachment_path))
            attach.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
            message.attach(attach)

        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")

        # Draft the email
        try:
            create_message = {'message': {'raw': raw_message}}
            service.users().drafts().create(userId="me", body=create_message).execute()
            print(f"Email to {vendor} successfully created, continuing...")
        except Exception as e:
            print(f"Error creating email: {e}")
    else:
        print(f"No file present, skipping email to {vendor}...")