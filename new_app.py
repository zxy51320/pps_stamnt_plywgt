# -*- coding: utf-8 -*-
"""
Tkinter-based GUI wrapper:
- Select a CSV file (uses only the first row, consistent with original logic).
- Click "Run": if no stored credentials, prompt for username/password and save them in plaintext.
- A "Forget / Re-enter Credentials" button deletes the stored credentials and immediately re-prompts.
- Playwright automation runs in a background thread to keep the UI responsive.
WARNING: Credentials are stored IN PLAINTEXT (for demo parity with your requirement).
"""

import os, sys, subprocess, json, csv, logging, threading, traceback, shutil
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import Tk, Button, Label, filedialog, messagebox, StringVar, simpledialog, ttk

from playwright.sync_api import sync_playwright
from Pages.login_page import LoginPage
from Pages.dashboard_page import DashboardPage
from typing import Optional

def find_chrome_exe() -> Optional[str]:
    """
    尝试在常见路径/注册表中查找 Chrome 可执行文件路径。
    返回 None 表示未找到。
    """
    # 先试试 PATH 里有没有 chrome.exe
    chrome_in_path = shutil.which("chrome.exe")
    if chrome_in_path:
        return chrome_in_path

    # 常见安装路径
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p

    # 注册表查找
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe") as k:
            chrome_path, _ = winreg.QueryValueEx(k, None)
            if chrome_path and os.path.exists(chrome_path):
                return chrome_path
    except Exception:
        pass

    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe") as k:
            chrome_path, _ = winreg.QueryValueEx(k, None)
            if chrome_path and os.path.exists(chrome_path):
                return chrome_path
    except Exception:
        pass

    return None

APP_TITLE = "PPS New Application"
CRED_FILE = Path.home() / ".new_app_credentials.json"

# ------------- CSV loader (first row only) -------------
def load_data(csv_path: str) -> dict:
    try:
        with open(csv_path, mode='r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            row = next(reader, None)
            if row is None:
                raise ValueError("CSV has no data (no first row).")

            owner_name = (row.get("Owner Name") or "").strip()
            try:
                name_parts = owner_name.split()
                row['first_name'] = name_parts[0]
                row['last_name'] = name_parts[-1]
            except Exception:
                row['first_name'] = row['last_name'] = owner_name

            for key in ("Home State", "State Issued", "State"):
                val = (row.get(key) or "").strip()
                row[key] = val.split(" (")[0] if val else ""

            __dob = row['Date of Birth'].split("-")
            row['Date of Birth'] = f"{__dob[1]}{__dob[2]}{__dob[0]}"

            yesterday = datetime.now() - timedelta(days=1)
            row['Yesterday'] = yesterday.strftime("%m%d%Y")
            return row
    except Exception as e:
        raise RuntimeError(f"Failed to read CSV: {e}")

# ------------- Plaintext credential helpers -------------
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

# ------------- Playwright main flow -------------
def run_playwright_flow(data: dict, username: str, password: str, log_cb, quantity: int):
    with sync_playwright() as p:
        log_cb("Locating Chrome…")
        chrome_path = find_chrome_exe()
        if not chrome_path:
            raise RuntimeError("Google Chrome not found on this system.")

        log_cb(f"Launching Chrome at: {chrome_path}")
        browser = p.chromium.launch(
            headless=False,
            executable_path=chrome_path
        )
        context = browser.new_context(accept_downloads=True, viewport={"width": 1440, "height": 800})
        page = context.new_page()
        page.set_default_timeout(30000)

        login_page = LoginPage(page)
        dashboard = DashboardPage(page)

        log_cb("Logging in…")
        login_page.login(username, password)

        log_cb("Submitting dashboard flow…")
        dashboard.new_app(data, quantity)

        context.close()
        browser.close()
        log_cb("Done.")

# ------------- GUI -------------
class App:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.csv_path = None

        self.status = StringVar(value="Please select a CSV file, then click [Run].")
        self.csv_label = StringVar(value="CSV: Not selected")

        Button(root, text="Select CSV", command=self.pick_csv, width=22).pack(pady=6)
        Label(root, textvariable=self.csv_label, wraplength=480, justify="left").pack(pady=2)

        Label(root, text="Equipment Quantity").pack(pady=(8, 2))
        self.qty_var = StringVar(value="1")
        self.qty_combo = ttk.Combobox(
            root,
            textvariable=self.qty_var,
            state="readonly",
            values=[str(i) for i in range(0, 6)]
        )
        self.qty_combo.pack(pady=2)

        Button(root, text="Run", command=self.on_run, width=22).pack(pady=6)
        Button(root, text="Forget / Re-enter Credentials", command=self.on_forget_and_reenter, width=22).pack(pady=2)

        Label(root, textvariable=self.status, fg="#444", wraplength=520, justify="left").pack(pady=8)

        # File logging
        logging.basicConfig(
            filename="gui_log.log",
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )

    def set_status(self, msg: str):
        self.status.set(msg)
        self.root.update_idletasks()
        logging.info(msg)
        
    def set_qty_enabled(self, enabled: bool):
        # 运行时禁用，完成后恢复为只读下拉
        state = "readonly" if enabled else "disabled"
        self.qty_combo.configure(state=state)

    def pick_csv(self):
        path = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv")],
        )
        if path:
            self.csv_path = path
            self.csv_label.set(f"CSV: {path}")
            self.set_status("CSV selected. Click [Run] to start.")

    def prompt_and_save_creds(self, require_both=True):
        """
        Prompt for username/password and save them in plaintext.
        If user cancels either field (when require_both=True), return (None, None).
        """
        username = simpledialog.askstring("Login", "Enter username:", parent=self.root)
        if require_both and (username is None or username.strip() == ""):
            return None, None

        password = simpledialog.askstring("Login", "Enter password:", parent=self.root, show="*")
        if require_both and (password is None or password.strip() == ""):
            return None, None

        try:
            save_creds((username or "").strip(), (password or ""))
            return username, password
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save credentials:\n{e}")
            return None, None

    def on_forget_and_reenter(self):
        """
        Delete stored credentials and immediately re-prompt for new ones.
        """
        removed = delete_creds()
        if removed:
            messagebox.showinfo("Credentials", "Stored credentials deleted.")
        else:
            messagebox.showinfo("Credentials", "No stored credentials found. You can enter new ones now.")

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

        # Load data from CSV
        try:
            data = load_data(self.csv_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read CSV:\n{e}")
            return

        # Load or prompt for credentials (PLAINTEXT)
        creds = load_creds()
        if not creds:
            username, password = self.prompt_and_save_creds(require_both=True)
            if not username or not password:
                messagebox.showwarning("Notice", "Username or password not provided.")
                return
        else:
            username = creds.get("username", "")
            password = creds.get("password", "")
            if not username or not password:
                messagebox.showwarning("Notice", "Stored credentials are incomplete. Please use 'Forget / Re-enter Credentials'.")
                return
        try:
            quantity = int(self.qty_var.get())
        except Exception:
            quantity = 1
        
        self.set_qty_enabled(False) # Disable qty combobox during run

        # Background thread for Playwright flow
        def worker():
            try:
                self.set_status("Running… (a browser window will open)")
                run_playwright_flow(data, username, password, self.set_status, quantity)
                self.set_status("Task completed ✅")
                messagebox.showinfo("Success", "Flow completed.")
            except Exception as e:
                err = f"Runtime error:\n{e}\n\n{traceback.format_exc()}"
                logging.error(err)
                self.set_status("Run failed ❌")
                messagebox.showerror("Error", err)
            finally:
                self.root.after(0, lambda: self.set_qty_enabled(True)) # Restore qty combobox state
                
        threading.Thread(target=worker, daemon=True).start()

def main():
    root = Tk()
    root.geometry("560x240")
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
