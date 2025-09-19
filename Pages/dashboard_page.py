# pages/dashboard_page.py
import time, re, logging
from playwright.sync_api import Page
from datetime import datetime
from dateutil.relativedelta import relativedelta


class DashboardPage:
    
    def __init__(self, page: Page):
        self.page = page
        self.now = datetime.now()
        self.last_month_date = self.now - relativedelta(months=1)
        self.last_month_name = self.last_month_date.strftime('%B')

    def search_mid(self, mid: str):
        try:
            self.page.get_by_role("textbox", name="Filter...").click()
            self.page.get_by_role("textbox", name="Filter...").fill(mid)
            self.page.get_by_role("tab", name="Merchant", exact=True).click(timeout=3000)
            self.page.get_by_role("button", name="Go To").click(timeout=3000)
        except Exception as e:
            logging.error(f"Can not find MID {mid}")
            raise e

    def download_pdf(self,mid: str):
        time.sleep(2)  # Wait for the page to load
        self.page.get_by_role("tab", name="SUMMARY").click()
        if self.page.get_by_text("Closed").first.is_visible():
            logging.error(f"Merchant {mid} is closed, skipping download.")
            return
        # Determine the tab to click based on the MID prefix
        first_four = mid[:4]
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
        # Find Statements tab
        while not self.page.get_by_role("tab", name="Statements").is_visible():
            self.page.get_by_test_id("MerchantDetails-WaitSpinner").locator("svg").nth(3).click()
        self.page.get_by_role("tab", name="Statements").click()
        try:
            if self.page.get_by_role("tab", name="Statements - Legacy").is_visible():
                self.page.locator(".src-components-dynamic-grid-styles--cell").first.click()
            else:
                self.page.get_by_role("cell", name=f"{self.last_month_name}").click()
        except Exception:
            logging.error(f"No {self.last_month_name} statement found for MID {mid}")
            raise Exception(f"No {self.last_month_name} statement found for MID {mid}")
        time.sleep(5)
        # Check if Legacy and make sure the statement is for the correct MID
        if self.page.get_by_role("button", name="Key 3").is_visible() and self.page.locator("iframe").content_frame.get_by_text(mid).is_visible():
            logging.info(f"Auto Key 3")
        elif self.page.get_by_role("cell", name=f"MERCHANT : {mid}", exact=True).is_visible():
            logging.info(f"Lagacy")
        elif self.page.get_by_role("button", name="Summary").is_visible() and self.page.locator("iframe").content_frame.get_by_text(mid).is_visible():
            self.page.get_by_role("button", name="Summary").click()
            self.page.get_by_role("option", name="Key 3").click()
            logging.info(f"Manual Key 3")
        else:
            logging.error(f"Cannot find the statement for MID {mid}")
            raise Exception(f"Cannot find the statement for MID {mid}")
        time.sleep(2)
        with self.page.expect_download() as download_info:
            self.page.get_by_role("button", name="PDF").click()
        download = download_info.value
        download.save_as(f"statements/{mid}.pdf")
        time.sleep(1)
        logging.info(f"Downloaded statement for MID {mid}")
        self.page.reload()
        
        
    def new_app(self,data: dict):
        self.page.get_by_role("navigation").get_by_text("Merchant").click()
        self.page.get_by_text("Applications").click()
        self.page.get_by_role("button", name="Add Application").click()
        # Location Search Tab
        self.page.get_by_role("textbox", name="Select a Partner").click()
        if "China" in data['Department']:
            if "Cash Discount" in data['Pricing Type']:
                self.page.get_by_role("textbox", name="Select a Partner").fill("Team FZ (CD)")
                time.sleep(1)
                self.page.get_by_text("P542582451267889").click()
                __caseType = 0
            else:
                self.page.get_by_role("textbox", name="Select a Partner").fill("Team FZ")
                time.sleep(1)
                self.page.get_by_text("P927165793457905").click()
                __caseType = 1
        else:
            self.page.get_by_role("textbox", name="Select a Partner").fill("ZBS-Kevin.")
            time.sleep(1)
            self.page.get_by_text("P445234827259245").click()
            __caseType = 2
        try:
            # Location Tab
            self.page.get_by_role("button", name="Location", exact=True).click()
            self.page.get_by_role("textbox", name="Location Name (DBA)").fill(data['DBA'])
            self.page.get_by_role("textbox", name="Street (PO Box not allowed)").fill(data['Street'])
            self.page.get_by_role("textbox", name="City").fill(data['City'])
            self.page.get_by_role("textbox", name="State").fill(data['State'])
            self.page.get_by_role("option", name=data['State']).click()
            self.page.get_by_role("textbox", name="Postal Code").fill(data['ZIP'])
            self.page.get_by_placeholder("Business Established").type(data['Yesterday'])
            self.page.get_by_role("textbox", name="Location Phone").fill(data['Business Phone'])
            if "Massage" in data['Business Type']:
                self.page.get_by_role("textbox", name="Market Segment").fill("Services")
                self.page.get_by_role("option", name="Services").click()
                self.page.get_by_role("textbox", name="Industry Type").fill("Massage")
                self.page.get_by_role("option", name="- MASSAGE PARLORS").click()
            else:
                self.page.get_by_role("textbox", name="Market Segment").fill("Restaurant")
                self.page.get_by_role("option", name="Restaurant").click()
                self.page.get_by_role("textbox", name="Industry Type").fill("- EATING PLACES, RESTAURANTS")
                self.page.get_by_role("option", name="- EATING PLACES, RESTAURANTS").click()
            logging.info('Location Tab finished')
            # Products Tab
            self.page.get_by_role("button", name="Products").click()
            time.sleep(5)
            self.page.locator("div").filter(has_text=re.compile(r"^TSYSSelect$")).get_by_role("button").click()
            self.page.get_by_role("button", name="Configure").first.wait_for()
            time.sleep(5)
            self.page.locator("div").filter(has_text=re.compile(r"^MX MerchantSelect$")).get_by_role("button").click()
            self.page.get_by_role("button", name="Configure").nth(1).wait_for()
            time.sleep(5)
            logging.info("Added TSYS and MX_Merchant")
            self.page.get_by_role("button", name="Configure").first.click()
            self.page.get_by_role("textbox", name="Average Purchase ($)").fill(data['Average Sales Amount'])
            self.page.get_by_role("textbox", name="Average Monthly Sales ($)").fill(data['Estimated Monthly Sale Volume'])
            self.page.get_by_role("textbox", name="Swiped, EMV %").fill("98")
            self.page.get_by_role("textbox", name="Keyed, MOTO %").fill("2")
            if __caseType == 2:
                self.page.get_by_role("button", name="Global Welcome Kit - TSYS").click()
                self.page.get_by_role("option", name="TSYS - Welcome Kit").click()
                time.sleep(1)
                self.page.get_by_role("button", name="ZBS Kevin Tsys Stmt Branding").click()
                self.page.get_by_role("option", name="PIS - Wholesale Processing").click()
            else:
                self.page.get_by_role("button", name="Global Welcome Kit - TSYS").click()
                self.page.get_by_role("option", name="Welcome Kit - TSYS", exact=True).click()
                if self.page.get_by_role("button", name="Team FZ TSYS Statement Branding").is_visible():
                    self.page.get_by_role("button", name="Team FZ TSYS Statement Branding").click()
                    self.page.get_by_role("option", name="PIS - Wholesale Processing").click()

            self.page.get_by_role("tab", name="Features").click()
            self.page.get_by_role("combobox").select_option("combined - all batches")
            self.page.get_by_role("tab", name="Pricing").click()
            # Set Rates
            logging.info("Setting Rates")
            self.page.locator("label:has-text('Select the Pricing Template')").locator(
                "xpath=following-sibling::div[contains(@class,'MuiInputBase-root')]//div[@role='button' and @aria-haspopup='listbox']").click()
            if data['Pricing Type'] == "Interchange":
                self.page.get_by_role("option", name="WPN- Pass Through").click()
                self.page.get_by_role("cell", name="0.10 Percent (%)").get_by_role("textbox").fill(data['V/M/D Rate(_.__%)'])
                self.page.get_by_role("row", name="Authorization : Any: Visa,").get_by_role("textbox").fill(data['V/M/D Fee($_.___)'])
                self.page.get_by_role("cell", name="0.55 Percent (%)").get_by_role("textbox").fill(data['Amex Rate(_.__%)'])
                self.page.get_by_role("cell", name="0.10 Per Event ($)").get_by_role("textbox").fill(data['Amex Fee($_.__)'])
                self.page.get_by_role("row", name="Recurring : Service: Monthly").get_by_role("textbox").fill(data['Monthly Fee'])
            elif data['Pricing Type'] == "Flat Rate":
                self.page.get_by_role("option", name="True Flat Rate").click()
                if float(data['V/M/D Rate(_.__%)']) == 0:
                    self.page.get_by_role("row", name="Sale : Any: MasterCard,").get_by_role("button").first.click()
                    self.page.get_by_role("button", name="Percent (%)").click()
                    self.page.get_by_role("option", name="Per Event ($)").click()
                    self.page.get_by_role("button", name="Save").click()
                    self.page.get_by_role("row", name="Sale : Any: MasterCard,").get_by_role("textbox").fill(data['V/M/D Fee($_.___)'])
                    self.page.get_by_role("cell", name="35.00 Per Event ($)").get_by_role("textbox").fill(data['Monthly Fee'])
                else:
                    self.page.get_by_role("row", name="Sale : Any: MasterCard,").get_by_role("textbox").fill(data['V/M/D Rate(_.__%)'])
                    self.page.get_by_role("cell", name="35.00 Per Event ($)").get_by_role("textbox").fill(data['Monthly Fee'])
            elif data['Pricing Type'] == "Cash Discount (by Percentage %)":
                self.page.get_by_role("option", name="True Flat Rate").click()
                __cdPrice = f"{(1 - (100 / (100 + float(data['Cash Discount Rate']))))*100:.2f}"
                self.page.get_by_role("row", name="Sale : Any: MasterCard,").get_by_role("textbox").fill(__cdPrice)
                self.page.get_by_role("cell", name="35.00 Per Event ($)").get_by_role("textbox").fill(data['Monthly Fee'])
            elif data['Pricing Type'] == "Cash Discount (by Flat Fee $)":
                self.page.get_by_role("option", name="True Flat Rate").click()
                self.page.get_by_role("row", name="Sale : Any: MasterCard,").get_by_role("button").first.click()
                self.page.get_by_role("button", name="Percent (%)").click()
                self.page.get_by_role("option", name="Per Event ($)").click()
                self.page.get_by_role("button", name="Save").click()
                self.page.get_by_role("row", name="Sale : Any: MasterCard,").get_by_role("textbox").fill(data['Cash Discount Rate'])
                self.page.get_by_role("cell", name="35.00 Per Event ($)").get_by_role("textbox").fill(data['Monthly Fee'])
            # Billing Tab
            self.page.get_by_role("tab", name="Billing").click()
            logging.info('Setting Bank Info')
            self.page.get_by_role("textbox", name="Routing #").first.fill(data['Bank Routing'])
            time.sleep(1)
            self.page.get_by_role("textbox", name="Routing #").nth(1).fill(data['Bank Routing'])
            time.sleep(1)
            self.page.get_by_role("textbox", name="Account #").first.fill(data['Bank Account'])
            self.page.get_by_role("textbox", name="Account #").nth(1).fill(data['Bank Account'])
            # Equipment Tab
            self.page.get_by_role("tab", name="Equipment").click()
            self.page.get_by_text("Add Equipment").click()
            self.page.locator("input[placeholder='Search by Type, Vendor, Product or Feature']:enabled").fill("S80 (Class B PPS Build)")
            self.page.locator(".src-components-form-combobox-styles--optionsContainer:visible").get_by_text("S80 (Class B PPS Build)").click()
            self.page.locator("input[placeholder='Select an application...']:enabled").click()
            self.page.locator(".src-components-form-combobox-styles--optionsContainer:visible").get_by_text("RAO Stage Only PAX B+").click()
            self.page.get_by_role("banner").get_by_role("button").click()
            logging.info('Products Tab finished')
            # Legal Tab
            self.page.get_by_role("button", name="Legal").click()
            self.page.get_by_role("textbox", name="Client Legal Entity Name").fill(data['Legal Name of Business'])
            self.page.get_by_role("checkbox", name="Copy Location Address").check()
            if "LLC" in data['Legal Type']:
                self.page.get_by_role("textbox", name="Legal Entity Type").fill("LLC")
                self.page.get_by_role("option", name="LLC", exact=True).click()
            elif "Corporation" in data['Legal Type']:
                self.page.get_by_role("textbox", name="Legal Entity Type").fill("S-Corp")
                self.page.get_by_role("option", name="S-Corp", exact=True).click()
            else:
                self.page.get_by_role("textbox", name="Legal Entity Type").fill("Individual or Sole Proprietor")
                self.page.get_by_role("option", name="Individual or Sole Proprietor", exact=True).click()
            
            self.page.get_by_role("textbox", name="Legal Entity Email").fill(data['Email'])
            self.page.get_by_role("textbox", name="State of Organization").fill(data['State'])
            self.page.get_by_role("option", name=data['State']).click()
            self.page.get_by_placeholder("Entity Formation Date").type(data['Yesterday'])
            self.page.get_by_role("textbox", name="Tax Id Type *").click()
            if len(data['Tax ID'].split('-')[0]) == 2:
                self.page.get_by_role("option", name="EIN").click()
                self.page.get_by_role("textbox", name="Tax ID", exact=True).type(data['Tax ID'])
            else:
                self.page.get_by_role("option", name="SSN").click()
                self.page.get_by_role("textbox", name="SSN", exact=True).type(data['Tax ID'])
            logging.info('Legal Tab finished')
            # Ownership Tab
            self.page.get_by_role("button", name="Ownership").click()
            self.page.get_by_role("textbox", name="First Name").fill(data['first_name'])
            self.page.get_by_role("textbox", name="Last Name").fill(data['last_name'])
            self.page.get_by_role("textbox", name="Home Address (PO Box not").fill(data['Home Street'])
            self.page.get_by_role("textbox", name="City").fill(data['Home City'])
            self.page.get_by_role("button", name="Open").first.click()
            self.page.get_by_role("option", name=data['Home State']).click()
            self.page.get_by_role("textbox", name="Postal Code").fill(data['Home ZIP'])
            self.page.get_by_placeholder("Date of Birth").type(data['Date of Birth'])
            self.page.get_by_role("textbox", name="Email Address").fill(data['Email'])
            self.page.get_by_role("textbox", name="Phone Number").fill(data['Mobile'])
            self.page.get_by_role("textbox", name="SSN").fill(data['Social Security Number'])
            if len(data['Driver License Number']) > 1:
                self.page.get_by_role("textbox", name="Driver's License").fill(data['Driver License Number'])
                self.page.get_by_role("button", name="Open").nth(2).click()
                self.page.get_by_role("option", name=data['State Issued']).click()
            logging.info('Ownership Tab finished')
            self.page.get_by_role("button", name="Save").click()
            time.sleep(3)
        except Exception as e:
            logging.error(f"Error: {e}. Save draft befor exit")
            self.page.get_by_role("button", name="Save").click()
            time.sleep(3)
