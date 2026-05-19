# -*- coding: utf-8 -*-
"""
pyinstaller -F -w --name "PPS_New_Application" --hidden-import "playwright._impl._api_types" --collect-all playwright --collect-submodules playwright --add-data "Pages;Pages" new_app.py
"""

import os, json, csv, logging, multiprocessing, queue, threading, traceback, shutil
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import Tk, Button, Label, Toplevel, filedialog, messagebox, StringVar, simpledialog, ttk

from playwright.sync_api import sync_playwright
from Pages.login_page import LoginPage
from Pages.dashboard_page import DashboardPage
from typing import Callable, Optional


class FlowStopped(Exception):
    pass


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
def run_playwright_flow(
    data: dict,
    username: str,
    password: str,
    log_cb,
    quantity: int,
    mfa_prompt_cb: Optional[Callable[[], bool]] = None,
    stop_event: Optional[threading.Event] = None,
    browser_cb: Optional[Callable[[Optional[object]], None]] = None,
):
    def check_stopped():
        if stop_event and stop_event.is_set():
            raise FlowStopped("Run stopped by user.")

    browser = None
    context = None

    with sync_playwright() as p:
        check_stopped()
        log_cb("Locating Chrome…")
        chrome_path = find_chrome_exe()
        if not chrome_path:
            raise RuntimeError("Google Chrome not found on this system.")

        log_cb(f"Launching Chrome at: {chrome_path}")
        check_stopped()
        browser = p.chromium.launch(
            headless=False,
            executable_path=chrome_path
        )
        if browser_cb:
            browser_cb(browser)
        check_stopped()
        context = browser.new_context(accept_downloads=True, viewport={"width": 1440, "height": 800})
        page = context.new_page()
        page.set_default_timeout(30000)

        login_page = LoginPage(page)
        dashboard = DashboardPage(page)

        log_cb("Logging in…")
        login_page.login(username, password)

        check_stopped()

        if mfa_prompt_cb:
            log_cb("Waiting for two-step authentication...")
            if not mfa_prompt_cb():
                raise FlowStopped("Run stopped by user.")

        log_cb("Submitting dashboard flow…")
        check_stopped()
        dashboard.new_app(data, quantity)

        check_stopped()
        context.close()
        browser.close()
        if browser_cb:
            browser_cb(None)
        log_cb("Done.")


def run_playwright_worker(data, username, password, quantity, status_queue, stop_event, mfa_event):
    def log_cb(msg: str):
        status_queue.put(("status", msg))

    def mfa_prompt_cb() -> bool:
        status_queue.put(("mfa", None))
        while not stop_event.is_set():
            if mfa_event.wait(0.2):
                return True
        return False

    try:
        run_playwright_flow(
            data,
            username,
            password,
            log_cb,
            quantity,
            mfa_prompt_cb,
            stop_event,
        )
        status_queue.put(("done", None))
    except FlowStopped:
        status_queue.put(("stopped", None))
    except Exception as e:
        status_queue.put(("error", f"Runtime error:\n{e}\n\n{traceback.format_exc()}"))

# ------------- GUI -------------
class App:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.csv_path = None
        self.worker_process = None
        self.status_queue = None
        self.stop_event = None
        self.mfa_event = None
        self.mfa_prompt_window = None
        self.ui_thread_id = threading.get_ident()

        self.status = StringVar(value="Please select a CSV file, then click [Run].")
        self.csv_label = StringVar(value="CSV: Not selected")

        self.select_csv_button = Button(root, text="Select CSV", command=self.pick_csv, width=22)
        self.select_csv_button.pack(pady=6)
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

        self.run_button = Button(root, text="Run", command=self.on_run, width=22)
        self.run_button.pack(pady=6)
        self.creds_button = Button(root, text="Forget / Re-enter Credentials", command=self.on_forget_and_reenter, width=22)
        self.creds_button.pack(pady=2)

        Label(root, textvariable=self.status, fg="#444", wraplength=520, justify="left").pack(pady=8)

        # File logging
        logging.basicConfig(
            filename="gui_log.log",
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )

    def set_status(self, msg: str):
        logging.info(msg)
        if threading.get_ident() == self.ui_thread_id:
            self.status.set(msg)
            self.root.update_idletasks()
        else:
            self.root.after(0, lambda: self.status.set(msg))
        
    def set_qty_enabled(self, enabled: bool):
        # 运行时禁用，完成后恢复为只读下拉
        state = "readonly" if enabled else "disabled"
        self.qty_combo.configure(state=state)

    def set_running_state(self, running: bool):
        if running:
            self.select_csv_button.configure(state="disabled")
            self.creds_button.configure(state="disabled")
            self.set_qty_enabled(False)
            self.run_button.configure(text="Stop", command=self.on_stop, state="normal")
        else:
            self.select_csv_button.configure(state="normal")
            self.creds_button.configure(state="normal")
            self.set_qty_enabled(True)
            self.run_button.configure(text="Run", command=self.on_run, state="normal")

    def close_mfa_prompt(self):
        if self.mfa_prompt_window:
            try:
                self.mfa_prompt_window.destroy()
            except Exception:
                pass
            self.mfa_prompt_window = None

    def show_two_step_auth_prompt(self):
        if self.mfa_prompt_window:
            return
        if self.stop_event and self.stop_event.is_set():
            return

        dialog = Toplevel(self.root)
        self.mfa_prompt_window = dialog
        dialog.title("Two-Step Authentication")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.attributes("-topmost", True)

        Label(
            dialog,
            text=(
                "Please click Push on the browser page and complete the two-step authentication.\n\n"
                "After authentication is complete, click OK here to continue."
            ),
            wraplength=360,
            justify="left",
            padx=18,
            pady=14,
        ).pack()

        def finish():
            if self.mfa_event:
                self.mfa_event.set()
            self.close_mfa_prompt()

        Button(dialog, text="OK", command=finish, width=16).pack(pady=(0, 14))
        dialog.protocol("WM_DELETE_WINDOW", finish)
        dialog.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{max(x, 0)}+{max(y, 0)}")

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

    def on_stop(self):
        if not self.stop_event:
            return
        self.stop_event.set()
        self.set_status("Stopping...")
        self.run_button.configure(state="disabled")
        self.close_mfa_prompt()
        if self.worker_process and self.worker_process.is_alive():
            self.worker_process.terminate()

    def on_run(self):
        if self.worker_process and self.worker_process.is_alive():
            self.on_stop()
            return

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
        
        stop_event = multiprocessing.Event()
        self.stop_event = stop_event
        self.status_queue = multiprocessing.Queue()
        self.mfa_event = multiprocessing.Event()
        self.set_running_state(True)

        self.set_status("Running... (a browser window will open)")
        self.worker_process = multiprocessing.Process(
            target=run_playwright_worker,
            args=(data, username, password, quantity, self.status_queue, self.stop_event, self.mfa_event),
            daemon=True,
        )
        self.worker_process.start()
        self.root.after(100, self.poll_worker)

    def poll_worker(self):
        finished = False
        error_msg = None

        if self.status_queue:
            while True:
                try:
                    kind, payload = self.status_queue.get_nowait()
                except queue.Empty:
                    break

                if kind == "status":
                    self.set_status(payload)
                elif kind == "mfa":
                    self.show_two_step_auth_prompt()
                elif kind == "done":
                    finished = True
                elif kind == "stopped":
                    self.set_status("Run stopped.")
                    finished = True
                elif kind == "error":
                    error_msg = payload
                    finished = True

        if self.worker_process and self.worker_process.is_alive() and not finished:
            self.root.after(100, self.poll_worker)
            return

        if self.stop_event and self.stop_event.is_set():
            self.set_status("Run stopped.")
        elif error_msg:
            logging.error(error_msg)
            self.set_status("Run failed ❌")
            messagebox.showerror("Error", error_msg)
        elif finished:
            self.set_status("Task completed ✅")
            messagebox.showinfo("Success", "Flow completed.")

        self.cleanup_after_run()

    def cleanup_after_run(self):
        if self.worker_process:
            try:
                self.worker_process.join(timeout=0.2)
            except Exception:
                pass
        self.close_mfa_prompt()
        self.worker_process = None
        self.status_queue = None
        self.stop_event = None
        self.mfa_event = None
        self.set_running_state(False)

def main():
    root = Tk()
    root.geometry("560x240")
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
