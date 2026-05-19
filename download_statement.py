# -*- coding: utf-8 -*-
"""
pyinstaller -F -w --name "PPS_Statement_Downloader" --hidden-import "playwright._impl._api_types" --collect-all playwright --collect-submodules playwright --add-data "Pages;Pages" download_statement.py
"""

import os, re, csv, json, logging, threading, traceback, shutil, smtplib
from pathlib import Path
from typing import Optional, List
from tkinter import Tk, Button, Label, filedialog, messagebox, StringVar, simpledialog
from datetime import datetime
from playwright.sync_api import sync_playwright
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr
from Pages.login_page import LoginPage
from Pages.dashboard_page import DashboardPage

APP_TITLE = "Statement Downloader"
CRED_FILE = Path(os.getenv("APPDATA", Path.home())) / "pps_statement_creds.json"

# ---------- Chrome locator ----------
def find_chrome_exe() -> Optional[str]:
    chrome_in_path = shutil.which("chrome.exe")
    if chrome_in_path and os.path.exists(chrome_in_path):
        return chrome_in_path
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    # Registry lookup
    try:
        import winreg
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            with winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe") as k:
                chrome_path, _ = winreg.QueryValueEx(k, None)
                if chrome_path and os.path.exists(chrome_path):
                    return chrome_path
    except Exception:
        pass
    return None

# ---------- CSV loader (read "MID" column) ----------
def load_mids(csv_path: str):
    mids = {}
    with open(csv_path, newline="", encoding="utf-8-sig") as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            if row.get("is_Active") == "TRUE":
                mid_raw = (row.get("MID") or "").strip()
                # 与原脚本一致：去掉首尾非数字字符
                mid = re.sub(r'^\D+|\D+$', '', mid_raw)
                dba = (row.get("DBA") or "").strip()
                email = (row.get("Email Address") or "").strip()
                if mid and email and dba:
                    mids[mid] = (dba, email)
                else:
                    logging.warning(f"Skipping row with missing MID/Email/DBA: {row}")
            else:
                logging.info(f"Skipping inactive MID row: {row}")
    return mids

# ---------- Plaintext credentials ----------
def load_creds():
    if CRED_FILE.exists():
        try:
            return json.loads(CRED_FILE.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None

def save_creds(username: str, password: str):
    CRED_FILE.write_text(
        json.dumps({"username": username, "password": password}, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

def delete_creds():
    try:
        if CRED_FILE.exists():
            CRED_FILE.unlink()
            return True
        return False
    except Exception:
        return False

# ---------- Email sender ----------
def send_email_with_attachment(recipient: str, pdf_path: str, dba: str, mid: str, month: str) -> str:
    
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 465
    SENDER = "brianl@zbspos.com"
    APP_PASSWORD = "uwdc peoq kymg kncd"  # App Password

    msg = MIMEMultipart()
    msg["From"] = formataddr(("Customer Service", SENDER))
    msg["Reply-To"] = "cs@zbspos.com"
    msg["To"] = recipient
    msg["Subject"] = f"{month} Statement (月结单) for {dba} - {mid}"

    body = """
    Dear valued customer,
    Attached is your monthly statement. Please take a look. 
    The password is your business ZIP code.
    Feel free to contact us if you have any questions or concerns.

    尊敬的客户，
    附上您的本月账单，供您查看。
    账单密码为您商户的邮编(ZIP Code)。
    如有任何问题或需要帮助，请随时联系我们。
    """
    msg.attach(MIMEText(body, "plain"))

    path = Path(pdf_path)
    with open(path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f'attachment; filename="{path.name}"'
    )
    msg.attach(part)

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        try:
            server.login(SENDER, APP_PASSWORD)
            server.send_message(msg)
            status = "SENT_OK"
        except smtplib.SMTPRecipientsRefused as e:
            status = "RECIPIENT_REFUSED"
        except smtplib.SMTPAuthenticationError:
            status = "AUTH_FAILED"
        except smtplib.SMTPException as e:
            status = f"SMTP_ERROR: {e}"

    return status


# ---------- Playwright flow (download statements) ----------
def run_flow(csv_path: str, username: str, password: str, log_cb):
    mids = load_mids(csv_path)
    if not mids:
        raise RuntimeError("No MIDs found in the CSV file (column 'MID').")

    with sync_playwright() as p:
        log_cb("Locating Chrome…")
        chrome_path = find_chrome_exe()
        if not chrome_path:
            raise RuntimeError("Google Chrome not found. Please install Chrome and try again.")

        log_cb(f"Launching Chrome: {chrome_path}")
        browser = p.chromium.launch(headless=False, executable_path=chrome_path)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1440, "height": 800})
        page = context.new_page()
        page.set_default_timeout(20000)

        login_page = LoginPage(page)
        dashboard = DashboardPage(page)
        output_dir = f"{dashboard.last_month_name}_statements"
        os.makedirs(output_dir, exist_ok=True)

        log_cb("Logging in…")
        login_page.login(username, password)
        messagebox.showinfo(
            "Two-Step Authentication",
            "Please click Push on the page to complete two-step authentication, then click OK to continue."
        )

        closed_list = []          # 已关闭的 MID
        failures_rows = []        # 失败的行 (MID, Reason)

        for idx, (mid, (dba, email)) in enumerate(mids.items(), 1):
            try:
                status_prefix = f"[{idx}/{len(mids)}] MID {mid}"
                pdf_path = os.path.join(output_dir, f"{mid}.pdf")

                if os.path.exists(pdf_path):
                    log_cb(f"Skipped (already exists): {mid}")
                    continue

                log_cb(f"Downloading {status_prefix}…")
                dashboard.search_mid(mid)
                status, reason = dashboard.download_pdf(mid,log_cb)
                if status == "downloaded":
                    # 发送邮件
                    e_status = send_email_with_attachment(email, pdf_path, dba, mid, dashboard.last_month_name)
                    if e_status == "SENT_OK":
                        log_cb(f"Downloaded & emailed: {mid} -> {email}")
                    else:
                        failures_rows.append((mid, f"Email failed: {e_status}"))
                        log_cb(f"Email failed for {mid}: {e_status}")
                elif status == "closed":
                    closed_list.append(mid)
                    log_cb(f"Closed merchant: {mid} -> recorded")
                else:
                    # 重试一次
                    log_cb(f"Error on {mid}: {reason} (will retry once)")
                    try:
                        page.reload()
                        status2, reason2 = dashboard.download_pdf(mid,log_cb)
                        if status2 == "downloaded":
                            e_status2 = send_email_with_attachment(email, pdf_path, dba, mid, dashboard.last_month_name)
                            if e_status2 == "SENT_OK":
                                log_cb(f"Downloaded & emailed: {mid} -> {email}")
                            else:
                                failures_rows.append((mid, f"Email failed: {e_status2}"))
                                log_cb(f"Email failed for {mid}: {e_status2}")
                            log_cb(f"Retry success: {mid}")
                        elif status2 == "closed":
                            closed_list.append(mid)
                            log_cb(f"Closed merchant on retry: {mid}")
                        else:
                            failures_rows.append((mid, reason2))
                            log_cb(f"Failed twice: {mid} ({reason2})")
                    except Exception as e2:
                        failures_rows.append((mid, f"Retry exception: {e2}"))
                        log_cb(f"Failed twice: {mid} ({e2})")
            except Exception as e:
                failures_rows.append((mid, f"Loop exception: {e}"))
                log_cb(f"Unhandled error on {mid}: {e}")

        # 关闭浏览器
        context.close()
        browser.close()

        # --- 输出文件 ---
        date_tag = datetime.now().strftime("%Y%m%d")

        # 1) 已关闭的 MID -> txt
        if closed_list:
            closed_path = f"closed_mids_{date_tag}.csv"
            with open(closed_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["MID"])
                for m in closed_list:
                    writer.writerow([m])
            log_cb(f"Closed MIDs saved to {closed_path}")

        # 2) 失败的 MID -> csv (含原因)
        if failures_rows:
            failed_path = f"failed_mids_{date_tag}.csv"
            with open(failed_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["MID", "Reason"])
                writer.writerows(failures_rows)
            log_cb(f"Failed MIDs saved to {failed_path}")

        if failures_rows:
            log_cb(f"Completed with {len(failures_rows)} failures and {len(closed_list)} closed.")
        else:
            log_cb(f"All statements processed. Closed: {len(closed_list)}, Failures: 0")

# ---------- GUI ----------
class App:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.csv_path = None

        self.status = StringVar(value="Select a CSV with a 'MID' column, then click [Run].")
        self.csv_label = StringVar(value="CSV: Not selected")

        self.csv_button = Button(root, text="Select CSV", width=22, command=self.pick_csv)
        self.csv_button.pack(pady=6)
        Label(root, textvariable=self.csv_label, wraplength=520, justify="left").pack(pady=2)

        self.run_button = Button(root, text="Run", width=22, command=self.on_run)
        self.run_button.pack(pady=6)
        self.creds_button = Button(root, text="Forget / Re-enter Credentials", width=22, command=self.on_forget_and_reenter)
        self.creds_button.pack(pady=2)
        self.buttons = [self.csv_button, self.run_button, self.creds_button]

        Label(root, textvariable=self.status, fg="#444", wraplength=560, justify="left").pack(pady=8)

        # file logging
        logging.basicConfig(
            filename=f"statement_gui.log",
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )

    def set_status(self, msg: str):
        self.status.set(msg)
        self.root.update_idletasks()
        logging.info(msg)

    def set_buttons_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for button in self.buttons:
            button.config(state=state)

    def pick_csv(self):
        path = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv")]
        )
        if path:
            self.csv_path = path
            self.csv_label.set(f"CSV: {path}")
            self.set_status("CSV selected. Click [Run] to start.")

    def prompt_and_save_creds(self, require_both=True):
        username = simpledialog.askstring("Login", "Enter username:", parent=self.root)
        if require_both and (not username or not username.strip()):
            return None, None
        password = simpledialog.askstring("Login", "Enter password:", parent=self.root, show="*")
        if require_both and (not password or not password.strip()):
            return None, None
        try:
            save_creds(username.strip(), password or "")
            return username, password
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save credentials:\n{e}")
            return None, None

    def on_forget_and_reenter(self):
        removed = delete_creds()
        if removed:
            messagebox.showinfo("Credentials", "Stored credentials deleted.")
        else:
            messagebox.showinfo("Credentials", "No stored credentials found.")
        self.set_status("Please re-enter credentials.")
        u, p = self.prompt_and_save_creds(require_both=True)
        if u and p:
            messagebox.showinfo("Credentials", "New credentials saved.")
            self.set_status("New credentials saved. You can now click [Run].")
        else:
            self.set_status("Re-entry cancelled or incomplete. Credentials not updated.")

    def on_run(self):
        if not self.csv_path:
            messagebox.showwarning("Notice", "Please select a CSV file first.")
            return

        self.set_buttons_enabled(False)

        # load creds (plaintext)
        creds = load_creds()
        if not creds:
            username, password = self.prompt_and_save_creds(require_both=True)
            if not username or not password:
                messagebox.showwarning("Notice", "Username or password not provided.")
                self.set_buttons_enabled(True)
                return
        else:
            username = creds.get("username") or ""
            password = creds.get("password") or ""
            if not username or not password:
                messagebox.showwarning("Notice", "Stored credentials are incomplete. Please use 'Forget / Re-enter Credentials'.")
                self.set_buttons_enabled(True)
                return

        def worker():
            try:
                self.set_status("Running… (a browser window will open)")
                run_flow(self.csv_path, username, password, self.set_status)
                self.set_status("Task completed ✅")
                messagebox.showinfo("Success", "Flow completed.")
            except Exception as e:
                err = f"Runtime error:\n{e}\n\n{traceback.format_exc()}"
                logging.error(err)
                self.set_status("Run failed")
                messagebox.showerror("Error", err)
            finally:
                self.root.after(0, lambda: self.set_buttons_enabled(True))

        threading.Thread(target=worker, daemon=True).start()

def main():
    root = Tk()
    root.geometry("600x260")
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
