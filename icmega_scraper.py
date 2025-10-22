from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tabulate import tabulate
from datetime import datetime, date
from google.oauth2.service_account import Credentials
import gspread
from gspread.utils import rowcol_to_a1
import chromedriver_autoinstaller
import time
import pytz

import os
from dotenv import load_dotenv
load_dotenv()

# --- Load credentials ---
EMAIL1 = os.getenv("ICMEGA_USER1_EMAIL")
PASSWORD1 = os.getenv("ICMEGA_USER1_PASSWORD")
EMAIL2 = os.getenv("ICMEGA_USER2_EMAIL")
PASSWORD2 = os.getenv("ICMEGA_USER2_PASSWORD")

print("Loaded credentials:")
print(f"User1: {EMAIL1}, password: {'*' * len(PASSWORD1) if PASSWORD1 else None}")
print(f"User2: {EMAIL2}, password: {'*' * len(PASSWORD2) if PASSWORD2 else None}")

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
def login_to_icmega(email, password):
    print("🚀 Launching browser...")

    # Automatically install matching chromedriver
    chromedriver_autoinstaller.install()

    # Setup headless Chrome options for CI (GitHub Actions)
    options = Options()
    options.add_argument("--headless=new")  # modern headless mode
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(options=options)

    driver.get("https://center.icmega.co.il/login.aspx?_theme=A")
    print("🌐 Page opened.")

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "recall_UserName"))
        )
        driver.find_element(By.NAME, "recall_UserName").send_keys(email)
        driver.find_element(By.NAME, "recall_UserPassword").send_keys(password)
        driver.execute_script("SubmitForm('L');")
        print("✅ Logged in.")
    except Exception as e:
        print("❌ Login failed:", e)
        driver.quit()
        return None

    return driver

# --- Go to search page and insert date range ---
def go_to_search_and_enter_dates(driver, start_date, end_date, user_email="unknown_user"):
    print("📄 Navigating to search page...")

    try:
        wait = WebDriverWait(driver, 10)

        print("📂 Waiting for menu item 'מועדים והקצאות'...")
        menu_item = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[.//span[text()='מועדים והקצאות']]"))
        )
        menu_item.click()
        print("✅ Menu item clicked.")

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
        return True

    except Exception as e:
        print("❌ Failed during search step:", str(e))
        os.makedirs("artifacts", exist_ok=True)
        screenshot_file = f"artifacts/search_error_{user_email}.png"
        # Save screenshot for debugging
        driver.save_screenshot(screenshot_file)
        print(f"📸 Screenshot saved: {screenshot_file}")
        return False


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

    driver.get(event["link"])
    # driver.save_screenshot("debug_event_page.png")
    
    # ✅ Wait for AngularJS to finish before looking for ticket elements
    try:
        driver.execute_async_script("""
            var callback = arguments[arguments.length - 1];
            if (window.angular) {
                var el = document.querySelector('body');
                if (angular.element(el).injector()) {
                    angular.element(el).injector().get('$browser')
                        .notifyWhenNoOutstandingRequests(callback);
                } else {
                    callback(true);
                }
            } else {
                callback(true);
            }
        """)
    except Exception as e:
        print("⚠️ Skipping Angular wait:", e)
        
    try:
        # Wait for <ul> to appear
        WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "ul.list-group-horizontal > li"))
        )
    except Exception as e:
        print("❌ Timeout waiting for org list:", e)
        return []

    # Retry loop: wait until all TARGET_ORGS have total > 0 and (sold/total) format is present
    max_retries = 8  # 8 seconds max (1s * 8)

    for attempt in range(max_retries):
        org_elements = driver.find_elements(By.CSS_SELECTOR, 'ul.list-group-horizontal > li')
        found_orgs = {}
        for el in org_elements:
            try:
                a_tag = el.find_element(By.TAG_NAME, 'a')
                text = a_tag.text.strip()
                if '(' not in text or ')' not in text:
                    continue

                name_part, ticket_part = text.split('(')
                org_name = name_part.strip()
                if org_name not in TARGET_ORGS:
                    continue

                sold_str, total_str = ticket_part.strip(')').split('/')
                sold, total = int(sold_str), int(total_str)
                if total > 0:
                    found_orgs[org_name] = (sold, total)
            except Exception:
                continue

        if all(org in found_orgs for org in TARGET_ORGS):
            break  # All target orgs have valid data
        time.sleep(1)
    else:
        print("⚠️ Warning: Not all orgs loaded in time")

    # Now extract the org data as usual
    org_data = []
    for org_el in org_elements:
        try:
            a_tag = org_el.find_element(By.TAG_NAME, 'a')
            text = a_tag.text.strip()
            # print("🔎 Raw text:", repr(text))

            if '(' not in text or ')' not in text:
                continue

            name_part, ticket_part = text.split('(')
            org_name = name_part.strip()
            if org_name not in TARGET_ORGS:
                continue

            ticket_part = ticket_part.strip(')')
            sold_str, total_str = ticket_part.split('/')
            sold, total = int(sold_str), int(total_str)

            # print(f"✅ Parsed: {org_name} sold={sold} total={total}")
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

# --- Run for each user ---
def run_for_user(email, password, start_date, end_date):
    print(f"Starting process for user: {email}")

    driver = login_to_icmega(email, password)
    if not driver:
        print(f"Skipping user {email} due to login failure.")
        return []

 # ✅ Check if date entry and search succeeded
    success = go_to_search_and_enter_dates(driver, start_date, end_date, user_email=email)
    if not success:
        print(f"🚫 Skipping user {email} due to search page error.")
        driver.quit()
        return []
    
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

    driver.quit()
    print(f"Completed process for user: {email}")
    return all_ticket_data

# --- Update tickets ---
def update_sheet_with_ticket_data(sheet, all_ticket_data):
    print("📥 Updating Google Sheet with ticket data...")

    records = sheet.get_all_records()
    headers = sheet.row_values(1)
    # name_col = headers.index("הפקה")
    # location_col = headers.index("אולם")
    # date_col = headers.index("תאריך")
    # org_col = headers.index("ארגון")
    sold_col = headers.index("נמכרו")
    total_col = headers.index("קיבלו")
    updated_col = headers.index("עודכן לאחרונה")

    updates = []
    updated_rows = []
    updated_ticket_data = []
    not_updated = []
    unique_events = set()
    israel_tz = pytz.timezone("Asia/Jerusalem")
    now_israel = datetime.now(israel_tz).strftime('%d/%m/%Y %H:%M')

    for ticket in all_ticket_data:
        # Strip time if exists (e.g. '30/07/25 17:30' → '30/07/25')
        ticket_date_raw = ticket["date"].split()[0]
        # Get current time in Israel
        # Normalize to dd/mm/yyyy
        try:
            dt = datetime.strptime(ticket_date_raw, "%d/%m/%y") if len(ticket_date_raw.split("/")[-1]) == 2 else datetime.strptime(ticket_date_raw, "%d/%m/%Y")
            ticket_date = dt.strftime("%d/%m/%Y")
        except Exception as e:
            print(f"⚠️ Could not parse ticket date '{ticket['date']}':", e)
            continue

        found = False
        for i, row in enumerate(records, start=2):  # i is actual sheet row index
            if (
                row.get("הפקה") == ticket["name"]
                and row.get("תאריך") == ticket_date
                and row.get("ארגון") == ticket["organization"]
            ):
                # Prepare cell updates
                updates.extend([
                    {'range': rowcol_to_a1(i, sold_col + 1), 'values': [[ticket["sold"]]]},
                    {'range': rowcol_to_a1(i, total_col + 1), 'values': [[ticket["total"]]]},
                    {'range': rowcol_to_a1(i, updated_col + 1), 'values': [[now_israel]]}
                ])
                updated_rows.append(i)
                updated_ticket_data.append(ticket)
                # Count unique (name, date) pairs that were updated
                unique_events.add((ticket["name"], ticket_date))
                found = True
                break

        if not found:
            not_updated.append(ticket)

    if updates:
            sheet.batch_update(updates)

    print(f"✅ Updated {len(updated_rows)} rows in sheet.")
    print(f"🗂️  That covers {len(unique_events)} unique events.")
    print("🟩 Row numbers updated:", updated_rows)

     # 🧾 Build and print table of updated rows
    if updated_ticket_data:
        print("\n📊 Table of updated ticket data:")
        print(tabulate(updated_ticket_data, headers="keys", tablefmt="grid", stralign="center"))

    if not_updated:
        print(f"\n⚠️ {len(not_updated)} items were NOT matched in the sheet:")
        print(tabulate(not_updated, headers="keys", tablefmt="grid", stralign="center"))
    else:
        print("✅ All items matched and updated successfully.")

# --- Main execution flow ---
if __name__ == "__main__":
    gc = get_gspread_client()
    sheet = gc.open("דאטה אפשיט אופיס").worksheet("כרטיסים")
    start_date, end_date = get_date_range_from_sheet(sheet)

    if not start_date or not end_date:
        print("❌ No valid dates found in the sheet.")
    else:
        print(f"📅 Date range from sheet: {start_date} to {end_date}")

        # Run for both users
        user1_data = run_for_user(EMAIL1, PASSWORD1, start_date, end_date)
        user2_data = run_for_user(EMAIL2, PASSWORD2, start_date, end_date)

        all_ticket_data = user1_data + user2_data

        # Print data as a clean table
        print("🎉 All extracted data from both users:")

    if all_ticket_data:

        update_sheet_with_ticket_data(sheet, all_ticket_data)
        

        # Print rows with total == 0 separately
        zero_total = [row for row in all_ticket_data if row.get("total") == 0]

        if zero_total:
            print("\n⚠️ Events with 0 total tickets:")
            zero_table = tabulate(zero_total, headers="keys", tablefmt="grid", stralign="center")
            print(zero_table)
        else:
            print("\n✅ No events with total = 0")
    else:
        print("No data found.")



