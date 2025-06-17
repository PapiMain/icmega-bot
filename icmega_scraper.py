import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from dotenv import load_dotenv
import os

load_dotenv()  # Load variables from .env file

USER1_EMAIL = os.getenv("ICMEGA_USER1_EMAIL")
USER1_PASSWORD = os.getenv("ICMEGA_USER1_PASSWORD")
USER2_EMAIL = os.getenv("ICMEGA_USER2_EMAIL")
USER2_PASSWORD = os.getenv("ICMEGA_USER2_PASSWORD")


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


def get_date_range_from_sheet(sheet):
    date_col_index = 4  # Column 'תאריך' is the 5th column, 0-based index
    rows = sheet.get_all_values()[1:]  # Skip header

    dates = []
    for row in rows:
        try:
            date_str = row[date_col_index]
            if date_str:
                day, month, year = map(int, date_str.split("/"))
                dates.append(datetime.date(year, month, day))
        except:
            continue

    if not dates:
        return None, None

    return min(dates), max(dates)

# Connect to sheet and load data
gc = get_gspread_client()
sheet = gc.open("דאטה אפשיט אופיס").worksheet("כרטיסים")

start_date, end_date = get_date_range_from_sheet(sheet)
print("📅 Date range from sheet:", start_date, "to", end_date)
