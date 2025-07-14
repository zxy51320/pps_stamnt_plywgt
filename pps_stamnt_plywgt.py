import time
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://mxconnect.com/#/login")
    page.get_by_role("textbox", name="Username").click()
    page.get_by_role("textbox", name="Username").fill("123@123.com")
    page.get_by_role("textbox", name="Password").click()
    page.get_by_role("textbox", name="Password").fill("12345@")
    page.get_by_role("button", name="SIGN IN").click()
    page.get_by_role("textbox", name="Filter...").click()
    page.get_by_role("textbox", name="Filter...").fill("432423432432")
    page.get_by_role("tab", name="Merchant").click()
    page.get_by_role("button", name="Go To").click()
    page.get_by_role("tab", name="TSYS").click()
    page.get_by_role("tab", name="Reports").click()
    page.get_by_test_id("MerchantDetails-WaitSpinner").locator("svg").nth(3).click()
    time.sleep(1)
    page.get_by_test_id("MerchantDetails-WaitSpinner").locator("svg").nth(3).click()
    page.get_by_role("tab", name="Statements - MXC").click()
    page.locator("td:nth-child(8)").first.click()
    page.get_by_role("button", name="Summary").click()
    page.get_by_role("option", name="Key 3").click()
    with page.expect_download() as download_info:
        page.get_by_role("button", name="PDF").click()
    download = download_info.value
    download.save_as("statement.pdf")
    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
