#playwright codegen https://mxconnect.com/#/login
import csv, logging
from playwright.sync_api import sync_playwright,Playwright
from Pages.login_page import LoginPage
from Pages.dashboard_page import DashboardPage
from datetime import datetime, timedelta


def load_data(csv_path: str):
    try:
        with open(csv_path, mode='r', newline='', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            row = next(csv_reader, None)  # 直接取第一行（可能是 None）
            if row is None:
                raise ValueError("No data in CSV")
            owner_name = (row.get("Owner Name") or "").strip()
            try:
                name_parts = owner_name.split()
                row['first_name'] = name_parts[0]
                row['last_name'] = name_parts[1]
            except:
                row['first_name'] = row['last_name'] = owner_name
                
            for key in ("Home State", "State Issued", "State"):
                val = (row.get(key) or "").strip()
                row[key] = val.split(" (")[0] if val else ""
                
            __dob = row['Date of Birth'].split("-")
            row['Date of Birth'] = f"{__dob[1]}{__dob[2]}{__dob[0]}"
            yesterday = datetime.now() - timedelta(days=1)
            row['Yesterday'] = yesterday.strftime("%m%d%Y")
        return row
    except:
        raise Exception

def main(data:dict):
    
    username = "Brianl@zbspos.com"
    password = "Brianl@12345"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1440, "height": 800})
        page = context.new_page()
        page.set_default_timeout(30000)
        login_page = LoginPage(page)
        dashboard = DashboardPage(page)
        login_page.login(username, password)
        dashboard.new_app(data)

        context.close()
        browser.close()

if __name__ == "__main__":
    # Configure logging
    log_file = f"log.log"
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    path = "TestData/Merchant Open Account (list_of_open_account).csv"
    data = load_data(path)
    main(data)
    
