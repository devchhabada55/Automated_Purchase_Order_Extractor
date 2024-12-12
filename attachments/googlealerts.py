import os
import base64
import openpyxl
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from bs4 import BeautifulSoup
import re
from datetime import datetime, timedelta

def run_google_alert_extraction(start_date, end_date):
    """
    Function to authenticate Gmail, extract Google Alerts, and save them into an Excel file.
    
    Args:
        start_date (str): Start date in 'YYYY-MM-DD' format
        end_date (str): End date in 'YYYY-MM-DD' format
    """
    # SCOPES for accessing Gmail
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
                    'C:\\Users\\Dev.chhabada\\OneDrive - Bekaert\\Desktop\\new\\client_secret_833632377056-e6ffq13j5o3numhiha15j6frtlp217jc.apps.googleusercontent.com.json',
                    SCOPES
                )
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        return build('gmail', 'v1', credentials=creds)

    # Function to clean text
    def clean_text(text):
        text = text.strip()
        text = re.sub(r'\s+', ' ', text)  # Remove multiple spaces and newlines
        return text

    # Extract relevant parts from Google Alert email
    def extract_google_alerts(service):
        # Convert input dates to datetime objects
        try:
            start_dt = datetime.strptime(start_date,  '%Y-%m-%d')
            end_dt = datetime.strptime(end_date, '%Y-%m-%d')
        except ValueError as e:
            raise ValueError("Invalid date format. Please use YYYY-MM-DD format.") from e

        # Validate date range
        if end_dt < start_dt:
            raise ValueError("End date must be after start date.")

        # Format dates for Gmail query
        start_query_date = start_dt.strftime('%Y/%m/%d')
        end_query_date = end_dt.strftime('%Y/%m/%d')

        # Gmail query to fetch emails within date range
        query = f"from:googlealerts-noreply@google.com subject:Google Alert - data center after:{start_query_date} before:{end_query_date}"

        results = service.users().messages().list(userId='me', q=query, maxResults=100).execute()
        messages = results.get('messages', [])
        data = []

        if not messages:
            print(f"No messages found between {start_date} and {end_date}.")
            return data

        for msg in messages:
            msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
            msg_date = datetime.fromtimestamp(int(msg_data['internalDate']) / 1000).strftime('%Y-%m-%d')
            parts = msg_data['payload'].get('parts', [])

            for part in parts:
                if part['mimeType'] == 'text/html':
                    msg_html = base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
                    soup = BeautifulSoup(msg_html, 'html.parser')

                    # Find all links (for titles) and associated text
                    news_items = soup.find_all('a', href=True)

                    for link in news_items:
                        if "flag as irrelevant" not in link.text.lower() and link.text.strip():
                            title = clean_text(link.text)  # Clean the title text
                            article_url = link['href']  # Extract the URL

                            # Find article summary by inspecting the neighboring content
                            brief = None
                            parent_td = link.find_parent('td')
                            if parent_td:
                                next_sibling = parent_td.find_next_sibling('td')
                                if next_sibling:
                                    brief = clean_text(next_sibling.get_text())

                            brief = brief if brief else "No summary available"

                            # Extract the source from nearby text or HTML structure
                            source = None
                            if parent_td:
                                sibling = parent_td.find_next('td')
                                if sibling:
                                    source = sibling.get_text()
                            source = clean_text(source) if source else 'Unknown'

                            # Append the extracted data
                            data.append({
                                'title': title,
                                'link': article_url,
                                'date': msg_date,
                                'source': 'Google-Alerts'
                            })

        return data

    # Write extracted data to Excel
    def write_to_excel(data):
        file_name = f'data-alerts_{start_date}_to_{end_date}.xlsx'
        if os.path.exists(file_name):
            wb = openpyxl.load_workbook(file_name)
            sheet = wb.active
        else:
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.append(["title", "link", "date", "source"])  # Add header

        for entry in data:
            sheet.append([entry['title'], entry['link'], entry['date'], entry['source']])

        wb.save(file_name)
        print(f"Data saved to {file_name}")

    # Main execution
    try:
        service = gmail_authenticate()
        alerts_data = extract_google_alerts(service)
        if alerts_data:
            write_to_excel(alerts_data)
            print(f"Successfully extracted {len(alerts_data)} alerts between {start_date} and {end_date}.")
        else:
            print("No data extracted for the specified date range.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    
google_alerts_df = run_google_alert_extraction("2024-11-08","2024-11-21")