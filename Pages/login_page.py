# pages/login_page.py
from playwright.sync_api import Page

class LoginPage:
    def __init__(self, page: Page):
        self.page = page

    def goto(self):
        self.page.goto("https://mxconnect.com/#/login")

    def login(self, username: str, password: str):
        self.page.get_by_role("textbox", name="Username").fill(username)
        self.page.get_by_role("textbox", name="Password").fill(password)
        self.page.get_by_role("button", name="SIGN IN").click()