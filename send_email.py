import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

# ===== 配置 =====
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

SENDER = "your_email@gmail.com"
APP_PASSWORD = "xxxx xxxx xxxx xxxx"  # App Password
RECEIVER = "target@example.com"

ATTACHMENT_PATH = "report.pdf"  # 任意文件

# ===== 构建邮件 =====
msg = MIMEMultipart()
msg["From"] = SENDER
msg["To"] = RECEIVER
msg["Subject"] = "Monthly Report"

# 正文（纯文本 or HTML）
body = """
Hi,

Please find the attached report.

Best,
Python Bot
"""
msg.attach(MIMEText(body, "plain"))

# ===== 附件 =====
path = Path(ATTACHMENT_PATH)
with open(path, "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())

encoders.encode_base64(part)
part.add_header(
    "Content-Disposition",
    f'attachment; filename="{path.name}"'
)
msg.attach(part)

# ===== 发送 =====
with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
    server.login(SENDER, APP_PASSWORD)
    server.send_message(msg)

print("✅ Email sent with attachment!")
