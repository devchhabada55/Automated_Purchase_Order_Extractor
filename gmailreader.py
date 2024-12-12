from datetime import timedelta
import os
import base64
import openpyxl
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from bs4 import BeautifulSoup
from datetime import datetime
import re

def extract_emails_to_excel(start_date, end_date):
    """
    Extract emails and save details in a structured Excel file.

    Args:
        start_date (str): Start date in 'YYYY-MM-DD' format
        end_date (str): End date in 'YYYY-MM-DD' format
    """
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

    # Gmail Authentication
    def gmail_authenticate():
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'your-credentials', SCOPES
                )
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        return build('gmail', 'v1', credentials=creds)

    # Function to decode email body
    def decode_message_body(part):
        try:
            if 'data' in part['body']:
                return base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
        except Exception:
            return None
        return None

    # Recursive function to extract body content
    def extract_body(payload):
        body_content = []
        if 'body' in payload and 'data' in payload['body']:
            decoded_body = decode_message_body(payload)
            if decoded_body:
                body_content.append(decoded_body)
        if 'parts' in payload:
            for part in payload['parts']:
                body_content.extend(extract_body(part))
        return body_content

    # Helper function to clean and parse email date
    def parse_email_date(date_header):
        try:
            clean_date = re.sub(r'\s*\(.*\)$', '', date_header)
            return datetime.strptime(clean_date, '%a, %d %b %Y %H:%M:%S %z').strftime('%Y-%m-%d')
        except Exception:
            return 'Unknown Date'

    # Extract email details
    def extract_email_details(service):
        adjusted_end_date = (datetime.strptime(end_date, '%Y-%m-%d') + timedelta(days=1)).strftime('%Y-%m-%d')
        query = f"after:{start_date} before:{adjusted_end_date}"
        results = service.users().messages().list(userId='me', q=query, maxResults=100).execute()
        messages = results.get('messages', [])
        data = []

        for msg in messages:
            msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
            payload = msg_data['payload']
            headers = payload.get('headers', [])
            
            # Extract sender, subject, and date
            sender = next((header['value'] for header in headers if header['name'] == 'From'), 'Unknown')
            subject = next((header['value'] for header in headers if header['name'] == 'Subject'), 'No Subject')
            date_header = next((header['value'] for header in headers if header['name'] == 'Date'), 'Unknown Date')
            email_date = parse_email_date(date_header)

            # Extract email body (both text and HTML)
            body_parts = extract_body(payload)
            body = "\n\n".join(body_parts).strip() if body_parts else "No Content Available"

            # Parse attachments
            attachments = []
            if 'parts' in payload:
                for part in payload['parts']:
                    if part['filename'] and part['body'].get('attachmentId'):
                        attachment_id = part['body']['attachmentId']
                        attachment = service.users().messages().attachments().get(
                            userId='me', messageId=msg['id'], id=attachment_id).execute()
                        attachment_data = base64.urlsafe_b64decode(attachment['data'])
                        filename = part['filename']
                        attachment_type = part['mimeType']

                        # Save attachment locally
                        filepath = os.path.join("attachments", filename)
                        os.makedirs("attachments", exist_ok=True)
                        with open(filepath, 'wb') as f:
                            f.write(attachment_data)

                        attachments.append({'filename': filename, 'type': attachment_type, 'link': filepath})

            for attachment in attachments:
                data.append({
                    'sender_email': sender,
                    'subject': subject,
                    'body': body,
                    'date': email_date,
                    'filename': attachment['filename'],
                    'attachment_type': attachment['type'],
                    'attachment_link': attachment['link']
                })

            if not attachments:
                data.append({
                    'sender_email': sender,
                    'subject': subject,
                    'body': body,
                    'date': email_date,
                    'filename': 'No attachment',
                    'attachment_type': 'None',
                    'attachment_link': 'N/A'
                })

        return data

    # Write to Excel
    def write_to_excel(data, filename='emails_data-testcase.xlsx'):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.append(['Sender Email', 'Subject', 'Body', 'Date', 'Filename', 'Attachment Link', 'Attachment Type'])

        for entry in data:
            sheet.append([
                entry['sender_email'], entry['subject'], entry['body'], entry['date'],
                entry['filename'], entry['attachment_link'], entry['attachment_type']
            ])

        wb.save(filename)
        print(f"Data saved to {filename}")

    # Main Execution
    try:
        service = gmail_authenticate()
        email_data = extract_email_details(service)
        if email_data:
            write_to_excel(email_data)
            print(f"Extracted {len(email_data)} emails.")
        else:
            print("No emails found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


# Example usage
extract_emails_to_excel("2024-11-30", "2024-12-02")
