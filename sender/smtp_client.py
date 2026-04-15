import smtplib
import os
import logging

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header

os.makedirs("logs", exist_ok=True)

logging.basicConfig(
    filename="logs/mail.log",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8"
)

SMTP_CONFIG = {
    "qq.com": ("smtp.qq.com", 465),
    "gmail.com": ("smtp.gmail.com", 465),
    "163.com": ("smtp.163.com", 465),
    "126.com": ("smtp.126.com", 465),
    "outlook.com": ("smtp-mail.outlook.com", 465),
    "hotmail.com": ("smtp-mail.outlook.com", 465),
    "live.com": ("smtp-mail.outlook.com", 465),
}


def _resolve_smtp_server(from_email):
    if not from_email or "@" not in from_email:
        return None, None, "发件人邮箱格式不正确"

    domain = from_email.split("@", 1)[1].lower()
    host_port = SMTP_CONFIG.get(domain)
    if not host_port:
        return None, None, f"暂不支持自动识别 {domain} 的SMTP服务器"

    return host_port[0], host_port[1], None


def send_mail(to_email, subject, body, attachments=None,
              from_email=None, auth_code=None):
    if not to_email:
        return False, "收件人不能为空"
    if not from_email:
        return False, "发件人不能为空"
    if not auth_code:
        return False, "授权码不能为空"

    host, port, server_error = _resolve_smtp_server(from_email)
    if server_error:
        logging.error(server_error)
        return False, server_error

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = str(Header(subject or "", "utf-8"))

    msg.attach(MIMEText(body, "plain", "utf-8"))

    if attachments:
        for path in attachments:
            if os.path.exists(path):
                with open(path, "rb") as f:
                    file_bytes = f.read()
                filename = os.path.basename(path)
                filename_header = str(Header(filename, "utf-8"))
                part = MIMEApplication(file_bytes, _subtype="octet-stream")
                part.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=("utf-8", "", filename)
                )
                part.add_header(
                    "Content-Type",
                    "application/octet-stream",
                    name=("utf-8", "", filename)
                )
                part.add_header("X-Attachment-Name", filename_header)
                logging.info("准备发送附件 path=%s filename=%s bytes=%s", path, filename, len(file_bytes))
                msg.attach(part)
            else:
                logging.warning(f"附件不存在: {path}")

    try:
        with smtplib.SMTP_SSL(host, port, timeout=15) as server:
            server.ehlo()
            server.login(from_email, auth_code)
            server.sendmail(from_email, to_email, msg.as_string())
        return True, "发送成功"

    except smtplib.SMTPAuthenticationError:
        logging.error("认证失败（授权码错误）")
        return False, "认证失败：请检查邮箱或授权码"

    except smtplib.SMTPConnectError:
        logging.error("连接服务器失败")
        return False, "连接服务器失败"

    except smtplib.SMTPRecipientsRefused:
        logging.error("收件人地址错误")
        return False, "收件人地址无效"

    except smtplib.SMTPException as e:
        logging.error(f"SMTP错误: {e}")
        return False, f"邮件发送失败: {e}"

    except Exception as e:
        logging.error(f"未知错误: {e}")
        return False, f"未知错误: {e}"