import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle

# Define scopes needed
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_gspread_client():
    creds = None

    # Load token if exists
    if os.path.exists('creds/token.pickle'):
        with open('creds/token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If no valid token, authenticate and save new one
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'creds/oauth_credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the token
        os.makedirs('creds', exist_ok=True)
        with open('creds/token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # Return an authorized gspread client
    return gspread.authorize(creds)

# Connect to Sheets
gc = get_gspread_client()

# Replace with your actual sheet name
spreadsheet = gc.open("דאטה אפשיט אופיס").sheet1

# Test: Print all rows
rows = spreadsheet.get_all_values()
for row in rows:
    print(row)
