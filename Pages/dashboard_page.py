# pages/dashboard_page.py
from ast import stmt
import time, re, logging, os
from playwright.sync_api import Page
from datetime import datetime
from dateutil.relativedelta import relativedelta
from playwright.sync_api import expect
from pypdf import PdfReader, PdfWriter


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

    def download_pdf(self,mid: str, log_cb):
        try:
            time.sleep(2)  # Wait for the page to load
            self.page.get_by_role("tab", name="SUMMARY").click()
            """
            if self.page.get_by_text("Closed").first.is_visible():
                log_cb(f"Merchant {mid} is closed, skipping download.")
                return "closed", "Merchant closed"
            """
            loc = self.page.get_by_text(re.compile(r"\b[A-Z]{2}\s+\d{5}(?:-\d{4})?\b")).nth(1)
            text = loc.inner_text()
            loc_zip = re.search(r'\b(?:[A-Z]{2})\s+(\d{5})(?:-\d{4})?\b', text).group(1)
            log_cb(f"Location ZIP for MID {mid}: {loc_zip}")
            # Determine the tab to click based on the MID prefix
            first_four = mid[:4]
            if first_four == "8739" or first_four == "5631":
                self.page.get_by_role("tab", name="TSYS").click()
                log_cb(f"MID: {mid} for TSYS")
            elif first_four == "8152":
                self.page.get_by_role("tab", name="Fiserv North").click()
                log_cb(f"MID: {mid} for Fiserv North")
            elif first_four == "5544":
                self.page.get_by_role("tab", name="Fiserv - Omaha").click()
                log_cb(f"MID: {mid} for Fiserv - Omaha")

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
                msg = f"No {self.last_month_name} statement found for MID {mid}"
                log_cb(msg)
                return "failed", msg

            time.sleep(5)

            # Make sure the statement is for the correct MID
            try:
                frame = self.page.frame_locator("iframe")
                expect(frame.get_by_text(mid)).to_be_visible(timeout=10000)
                found = True
            except Exception:
                try:
                    expect(self.page.get_by_role("cell", name=f"MERCHANT : {mid}", exact=True)).to_be_visible(timeout=10000)
                    found = True
                except Exception:
                    try:
                        pattern = re.compile(
                            rf"{mid[:4]}\s+{mid[4:8]}\s+{mid[-7:]}"
                        )
                        stmt = self.page.locator("div.src-components-statement-detail-styles--contentContainer pre")
                        expect(stmt).to_contain_text(pattern, timeout=10000)
                        found = True
                    except Exception:
                        found = False
            if found:
                with self.page.expect_download() as download_info:
                    self.page.get_by_role("button", name="PDF").click()
                download = download_info.value
                download.save_as(f"{self.last_month_name}_statements/{mid}.pdf")
                self.encrypt_pdf(f"{self.last_month_name}_statements/{mid}.pdf", loc_zip)
                """
                pyminizip.compress(f"{self.last_month_name}_statements/{mid}.pdf", None,
                                  f"{self.last_month_name}_statements/{mid}.zip", loc_zip, 5)
                try:
                    os.remove(f"{self.last_month_name}_statements/{mid}.pdf")
                except OSError as e:
                    logging.warning(f"Failed to delete: {f"{self.last_month_name}_statements/{mid}.pdf"} -> {e}")
                """
                log_cb(f"Downloaded statement for MID {mid}")
                self.page.reload(wait_until="load")
                return "downloaded", "ok"
            else:
                msg = f"Cannot find the statement content for MID {mid}"
                log_cb(msg)
                return "failed", msg

        except Exception as e:
            msg = f"Unexpected error for MID {mid}: {e}"
            log_cb(msg)
            return "failed", msg
        
    def encrypt_pdf(self, loc: str, password: str):
        reader = PdfReader(loc)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        writer.encrypt(user_password = password, owner_password = "zbspos_admin")

        with open(loc, "wb") as f:
            writer.write(f)
        
    def new_app(self,data: dict, quantity: int):
        self.page.get_by_role("navigation").get_by_text("Merchant").click()
        self.page.locator('.src-components-sidebar-styles--subMenu').get_by_text("Applications",exact=True).click()
        self.page.get_by_role("button", name="Add Application").click()
        # Location Search Tab
        self.page.get_by_role("textbox", name="Select a Partner").click()
        if "China" in data['Department']:
            if "Cash Discount" in data['Pricing Type'] or "Dual Pricing" in data['Pricing Type']:
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
            try:
                self.page.get_by_role("textbox", name="State").fill(data['State'])
                self.page.get_by_role("option", name=data['State']).click()
            except:
                pass
            self.page.get_by_role("textbox", name="Postal Code").fill(data['ZIP'])
            self.page.get_by_placeholder("Business Established").type(data['Yesterday'])
            self.page.get_by_role("textbox", name="Location Phone").fill(data['Business Phone'])
            if "Massage" in data['Business Type']:
                self.page.get_by_role("textbox", name="Market Segment").fill("Services")
                self.page.get_by_role("option", name="Services").click()
                self.page.get_by_role("textbox", name="Industry Type").fill("Massage")
                self.page.get_by_role("option", name="- MASSAGE PARLORS").click()
            elif "Restaurant" in data['Business Type']:
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
            __welcomeKit = self.page.locator('div.MuiFormControl-root', has_text="Welcome Kit").get_by_role("button")
            __statementBranding = self.page.locator('div.MuiFormControl-root', has_text="Statement Branding").get_by_role("button")
            if __caseType == 2:
                __welcomeKit.click()
                self.page.get_by_role("option", name="TSYS - Welcome Kit").click()
                time.sleep(1)
                __statementBranding.click()
                self.page.get_by_role("option", name="PIS - Wholesale Processing").click()
            else:
                __welcomeKit.click()
                self.page.get_by_role("option", name="Welcome Kit - TSYS", exact=True).click()
                __statementBranding.click()
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
            elif data['Pricing Type'] == "Cash Discount (by Percentage %)" or data['Pricing Type'] == "Dual Pricing (by Percentage %)":
                self.page.get_by_role("option", name="True Flat Rate").click()
                __cdPrice = f"{(1 - (100 / (100 + float(data['Cash Discount Rate']))))*100:.2f}"
                self.page.get_by_role("row", name="Sale : Any: MasterCard,").get_by_role("textbox").fill(__cdPrice)
                self.page.get_by_role("cell", name="35.00 Per Event ($)").get_by_role("textbox").fill(data['Monthly Fee'])
            elif data['Pricing Type'] == "Cash Discount (by Flat Fee $)" or data['Pricing Type'] == "Dual Pricing (by Flat Fee $)":
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
            for i in range(quantity):
                self.page.get_by_role("tab", name="Equipment").click()
                self.page.get_by_text("Add Equipment").click()
                self.page.locator("input[placeholder='Search by Type, Vendor, Product or Feature'][value='']").fill("S80 (Class B PPS Build)")
                self.page.locator(".src-components-form-combobox-styles--optionsContainer:visible").get_by_text("S80 (Class B PPS Build)").click()
                self.page.locator("input[placeholder='Select an application...'][value='']").click()
                self.page.locator(".src-components-form-combobox-styles--optionsContainer:visible").get_by_text("RAO Stage Only PAX B+").click()
            self.page.get_by_role("banner").get_by_role("button").click() # Back to Products
            logging.info('Products Tab finished')
            # Legal Tab
            self.page.get_by_role("button", name="Legal").click()
            self.page.get_by_role("textbox", name="Client Legal Entity Name").fill(data['Legal Name of Business'])
            self.page.get_by_role("checkbox", name="Copy Location Address").check()
            if data['Legal Type'] == "LLC":
                self.page.get_by_role("textbox", name="Legal Entity Type").fill("LLC")
                self.page.get_by_role("option", name="LLC", exact=True).click()
            elif data['Legal Type'] == "Corporation":
                self.page.get_by_role("textbox", name="Legal Entity Type").fill("S-Corp")
                self.page.get_by_role("option", name="S-Corp", exact=True).click()
            else:
                self.page.get_by_role("textbox", name="Legal Entity Type").fill("Individual or Sole Proprietor")
                self.page.get_by_role("option", name="Individual or Sole Proprietor", exact=True).click()
            
            self.page.get_by_role("textbox", name="Legal Entity Email").fill(data['Email'])
            try:
                self.page.get_by_role("textbox", name="State of Organization").fill(data['State'])
                self.page.get_by_role("option", name=data['State']).click()
            except:
                pass
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
            try:
                self.page.get_by_role("button", name="Open").first.click()
                self.page.get_by_role("option", name=data['Home State']).click()
            except:
                pass
            self.page.get_by_role("textbox", name="Postal Code").fill(data['Home ZIP'])
            self.page.get_by_placeholder("Date of Birth").type(data['Date of Birth'])
            self.page.get_by_role("textbox", name="Email Address").fill(data['Email'])
            self.page.get_by_role("textbox", name="Phone Number").fill(data['Mobile'])
            self.page.get_by_role("textbox", name="SSN").fill(data['Social Security Number'])
            if len(data['Driver License Number']) > 1:
                self.page.get_by_role("textbox", name="Driver's License").fill(data['Driver License Number'])
                try:
                    self.page.get_by_role("button", name="Open").nth(3).click()
                    self.page.get_by_role("option", name=data['State Issued']).click()
                except:
                    pass
            logging.info('Ownership Tab finished')
            self.page.get_by_role("button", name="Save").click()
            time.sleep(3)
        except Exception as e:
            logging.error(f"Error: {e}. Save draft befor exit")
            self.page.get_by_role("button", name="Save").click()
            time.sleep(3)
