import gspread
from google.oauth2.service_account import Credentials
import os

# Define scopes needed
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_gspread_client():
    # Path to your service account key JSON file
    SERVICE_ACCOUNT_FILE = 'creds/service_account.json'

    # Create credentials using service account file and scopes
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES
    )

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
