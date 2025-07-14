# pages/dashboard_page.py
import time
from playwright.sync_api import Page

class DashboardPage:
    def __init__(self, page: Page):
        self.page = page

    def search_mid(self, mid: str):
        self.page.get_by_role("textbox", name="Filter...").click()
        self.page.get_by_role("textbox", name="Filter...").fill(mid)
        self.page.get_by_role("tab", name="Merchant", exact=True).click()
        self.page.get_by_role("button", name="Go To").click()

    def navigate_to_statements(self):
        self.page.get_by_role("tab", name="TSYS").click()
        self.page.get_by_role("tab", name="Reports").click()
        try:
            self.page.get_by_test_id("MerchantDetails-WaitSpinner").locator("svg").nth(3).click()
            time.sleep(1)  # 等待内容加载
            self.page.get_by_test_id("MerchantDetails-WaitSpinner").locator("svg").nth(3).click()
        except Exception as e:
            pass
        self.page.get_by_role("tab", name="Statements - MXC").click()

    def download_pdf(self, save_path: str):
        self.page.locator("td:nth-child(8)").first.click()
        self.page.get_by_role("button", name="Summary").click()
        self.page.get_by_role("option", name="Key 3").click()
        with self.page.expect_download() as download_info:
            self.page.get_by_role("button", name="PDF").click()
        download = download_info.value
        download.save_as(save_path)
        #self.page.keyboard.press("Escape")
        self.page.get_by_role("button", name="").click() # 关闭下载对话框
        time.sleep(1)  # 等待下载完成
