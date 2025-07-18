# tests/test_download_statements.py
#playwright codegen https://mxconnect.com/#/login
import csv
from playwright.sync_api import sync_playwright,Playwright
from Pages.login_page import LoginPage
from Pages.dashboard_page import DashboardPage
import logging


def load_mids(csv_path: str):
    with open(csv_path, newline="") as csvfile:
        return [row["MID"] for row in csv.DictReader(csvfile)]

def statement_download():
    mids = load_mids("TestData/Merchant Open Account (list_of_open_account).csv")
    if not mids:
        logging.error("No MIDs found in the CSV file.")
        return
    username = "Brianl@zbspos.com"
    password = "1990100900pW@"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1440, "height": 800})
        page = context.new_page()
        page.set_default_timeout(15000)
        login_page = LoginPage(page)
        dashboard = DashboardPage(page)

        login_page.goto()
        login_page.login(username, password)
        failures = []
        for mid in mids:
            try:
                logging.info(f"Started - MID: {[mid]}")
                dashboard.search_mid(mid)
                dashboard.download_pdf(mid,f"statements/{mid}.pdf")
            except Exception as e:
                failures.append(mid)
                logging.error(f"Error processing MID {mid}: {e}")
        logging.info(f"Fallures: {failures}")
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
    statement_download()
