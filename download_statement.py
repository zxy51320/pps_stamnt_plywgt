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
    username = "Brianl@zbspos.com"
    password = "1990100900pW@"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1440, "height": 900})
        page = context.new_page()

        login_page = LoginPage(page)
        dashboard = DashboardPage(page)

        login_page.goto()
        login_page.login(username, password)

        for mid in mids:
            print(f"正在处理 MID: {mid}")
            dashboard.search_mid(mid)
            dashboard.navigate_to_statements()
            dashboard.download_pdf(f"statements/{mid}.pdf")

        context.close()
        browser.close()

statement_download()
