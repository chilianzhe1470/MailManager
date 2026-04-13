import smtplib
import os
import logging

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

logging.basicConfig(
    filename="logs/mail.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def send_mail(to_email, subject, body, attachments=None,
              from_email=None, auth_code=None):

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))

    if attachments:
        for path in attachments:
            if os.path.exists(path):
                with open(path, "rb") as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())

                encoders.encode_base64(part)
                filename = os.path.basename(path)

                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{filename}"'
                )

                msg.attach(part)

    try:
        server = smtplib.SMTP("smtp.qq.com", 587)
        server.starttls()
        server.login(from_email, auth_code)

        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()

        logging.info(f"发送成功: {to_email}")
        return True

    except Exception as e:
        logging.error(f"发送失败: {to_email}, 错误: {e}")
        return False