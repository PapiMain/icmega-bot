from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials
import time

import os
from dotenv import load_dotenv
load_dotenv()

# --- Load credentials ---
EMAIL = os.getenv("ICMEGA_USER1_EMAIL")
PASSWORD = os.getenv("ICMEGA_USER1_PASSWORD")
print("Loaded credentials:", EMAIL, "*" * len(PASSWORD) if PASSWORD else "None")

# --- Google Sheets setup ---
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'creds/service_account.json'

def get_gspread_client():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return gspread.authorize(creds)

def get_date_range_from_sheet(sheet):
    date_col_index = 3  # Column 'תאריך'
    rows = sheet.get_all_values()[1:]
    dates = []
    today = date.today()

    for row in rows:
        try:
            date_str = row[date_col_index]
            if date_str:
                day, month, year = map(int, date_str.split("/"))
                d = date(year, month, day)
                if d >= today:
                    dates.append(d)
        except:
            continue

    return (min(dates), max(dates)) if dates else (None, None)


# --- Selenium login ---
def login_to_icmega():
    print("🚀 Launching browser...")
    options = Options()
    options.add_argument("--start-maximized")
    service = Service("chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=options)

    driver.get("https://center.icmega.co.il/login.aspx?_theme=A")
    print("🌐 Page opened.")

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "recall_UserName"))
        )
        driver.find_element(By.NAME, "recall_UserName").send_keys(EMAIL)
        driver.find_element(By.NAME, "recall_UserPassword").send_keys(PASSWORD)
        driver.execute_script("SubmitForm('L');")
        print("✅ Logged in.")
    except Exception as e:
        print("❌ Login failed:", e)

    return driver

# --- Go to search page and insert date range ---
def go_to_search_and_enter_dates(driver, start_date, end_date):
    print("📄 Navigating to search page...")
    driver.get("https://center.icmega.co.il/mn_search.aspx?_TableName=sapak_product_barcode&sidebar=23")

    try:
        wait = WebDriverWait(driver, 10)

        print("⏳ Waiting for date fields...")
        wait.until(EC.presence_of_element_located((By.NAME, "event_start_date_from")))

        print("✅ Found date fields, checking checkbox...")
        checkbox = driver.find_element(By.NAME, "ChkOption")
        if not checkbox.is_selected():
            checkbox.click()

        print("📝 Filling in date range...")
        driver.find_element(By.NAME, "event_start_date_from").send_keys(start_date.strftime("%d/%m/%Y"))
        driver.find_element(By.NAME, "event_start_date_to").send_keys(end_date.strftime("%d/%m/%Y"))

        print("🔎 Clicking 'חפש' (search) button...")
        search_button = driver.find_element(By.XPATH, "//a[contains(text(),'חפש')]")
        search_button.click()

        print("✅ Search submitted.")
    except Exception as e:
        print("❌ Failed during search step:", e)

# --- Get all allocation links ---
def get_all_allocation_links(driver):
    print("🔍 Looking for allocation links...")

    try:
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(text(),'הקצאה')]")))

        rows = driver.find_elements(By.CSS_SELECTOR, "table.table-bordered tr")[1:]  # skip header row

        allocation_data = []
        for row in rows:
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) < 14:
                    continue  # not enough cells, skip

                name = cells[2].text.strip()  # קוד הצגה
                location = cells[3].text.strip()  # קוד אולם
                date = cells[5].text.strip()  # תאריך האירוע

                # Look for הקצאה link
                links = cells[13].find_elements(By.TAG_NAME, "a")
                for link in links:
                    href = link.get_attribute("href")
                    if href and "sapak_theatre_program.aspx" in href:
                        allocation_data.append({
                            "link": href,
                            "name": name,
                            "location": location,
                            "date": date
                        })

            except Exception as e:
                print("❌ Error reading row:", e)
                continue

        print(f"✅ Found {len(allocation_data)} allocation links.")
        return allocation_data

    except Exception as e:
        print("❌ Error finding allocation links:", e)
        return []



# --- Extract ticket data from allocation page ---
TARGET_ORGS = ['מגה לאן', 'חבר', 'קרנות השוטרים']

def extract_org_ticket_data(driver, event, wait_time=10):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import time

    driver.get(event["link"])
    driver.save_screenshot("debug_event_page.png")

    try:
        # Wait for <ul> to appear
        WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "ul.list-group-horizontal > li"))
        )
    except Exception as e:
        print("❌ Timeout waiting for org list:", e)
        return []

    # Wait up to 6 seconds for any of the numbers to become something other than 0/0
    for _ in range(12):  # check every 0.5s for 6s max
        org_elements = driver.find_elements(By.CSS_SELECTOR, 'ul.list-group-horizontal > li')
        valid_numbers = []
        for el in org_elements:
            try:
                a_tag = el.find_element(By.TAG_NAME, 'a')
                text = a_tag.text.strip()
                if '(' in text and ')' in text:
                    ticket_part = text.split('(')[-1].strip(')')
                    sold_str, total_str = ticket_part.split('/')
                    sold = int(sold_str)
                    total = int(total_str)
                    if total > 0:
                        valid_numbers.append((sold, total))
            except Exception:
                continue
        if valid_numbers:
            break
        time.sleep(0.5)

    # Now extract the org data as usual
    org_data = []
    for org_el in org_elements:
        try:
            a_tag = org_el.find_element(By.TAG_NAME, 'a')
            text = a_tag.text.strip()
            print("🔎 Raw text:", repr(text))

            if '(' not in text or ')' not in text:
                continue

            name_part, ticket_part = text.split('(')
            org_name = name_part.strip()
            if org_name not in TARGET_ORGS:
                continue

            ticket_part = ticket_part.strip(')')
            sold_str, total_str = ticket_part.split('/')
            sold, total = int(sold_str), int(total_str)

            print(f"✅ Parsed: {org_name} sold={sold} total={total}")
            org_data.append({
                "link": event["link"],
                "name": event["name"],
                "location": event["location"],
                "date": event["date"],
                "organization": org_name,
                "sold": sold,
                "total": total
            })
        except Exception as e:
            print("Error parsing org:", e)
            continue

    return org_data


# --- Main execution flow ---
if __name__ == "__main__":
    gc = get_gspread_client()
    sheet = gc.open("דאטה אפשיט אופיס").worksheet("כרטיסים")
    start_date, end_date = get_date_range_from_sheet(sheet)

    if not start_date or not end_date:
        print("❌ No valid dates found in the sheet.")
    else:
        print(f"📅 Date range from sheet: {start_date} to {end_date}")
        driver = login_to_icmega()
        go_to_search_and_enter_dates(driver, start_date, end_date)

        allocation_links = get_all_allocation_links(driver)
        all_ticket_data = []

        for item in allocation_links:
            url = item["link"] if isinstance(item, dict) else item
            if isinstance(url, str) and url.startswith("http"):
                ticket_data = extract_org_ticket_data(driver, item)
                all_ticket_data.extend(ticket_data)
            else:
                print(f"❌ Invalid URL: {url}")

            time.sleep(1)

        print("🎉 All extracted data:")
        for item in all_ticket_data:
            print(item)


