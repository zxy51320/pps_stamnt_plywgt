# pages/dashboard_page.py
import time
from playwright.sync_api import Page
import logging

class DashboardPage:
    def __init__(self, page: Page):
        self.page = page

    def search_mid(self, mid: str):
        try:
            self.page.get_by_role("textbox", name="Filter...").click()
            self.page.get_by_role("textbox", name="Filter...").fill(mid)
            self.page.get_by_role("tab", name="Merchant", exact=True).click(timeout=3000)
            self.page.get_by_role("button", name="Go To").click(timeout=3000)
        except Exception as e:
            logging.error(f"Can not find MID {mid}")
            raise e

    def download_pdf(self,mid: str,save_path: str):
        first_four = mid[:4]
        #815219362887
        #56317155030542
        #554402040703142
        #8739781940608005
        if first_four == "8739" or first_four == "5631":
            self.page.get_by_role("tab", name="TSYS").click()
            logging.info(f"MID: {mid} for TSYS")
        elif first_four == "8152":
            self.page.get_by_role("tab", name="Fiserv North").click()
            logging.info(f"MID: {mid} for Fiserv North")
        elif first_four == "5544":
            self.page.get_by_role("tab", name="Fiserv - Omaha").click()
            logging.info(f"MID: {mid} for Fiserv - Omaha")
        self.page.get_by_role("tab", name="Reports").click()
        while not self.page.get_by_role("tab", name="Statements").is_visible():
            self.page.get_by_test_id("MerchantDetails-WaitSpinner").locator("svg").nth(3).click()
        self.page.get_by_role("tab", name="Statements").click()
        try:
            self.page.locator(".src-components-dynamic-grid-styles--cell").first.click(timeout=3000)
        except Exception as e:
            logging.error(f"No statements found for MID {mid}")
            raise e
        try:
            self.page.get_by_role("button", name="Summary").click(timeout=3000)
            self.page.get_by_role("option", name="Key 3").click(timeout=3000)
            logging.info(f"Key 3")
        except Exception as e:
            logging.info(f"Lagacy")
        with self.page.expect_download() as download_info:
            self.page.get_by_role("button", name="PDF").click()
        download = download_info.value
        download.save_as(save_path)
        self.page.keyboard.press("Escape") # 关闭下载对话框
        #self.page.get_by_role("button", name="").click() 
        time.sleep(2)  # 控制下载频率
        logging.info(f"Downloaded statement for MID {mid}")
