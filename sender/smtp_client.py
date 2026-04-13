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
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8"
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
            else:
                logging.warning(f"附件不存在: {path}")

    try:
        server = smtplib.SMTP_SSL("smtp.qq.com", 465, timeout=15)
        server.ehlo()
        server.login(from_email, auth_code)

        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()

        return True, "发送成功"

    except Exception as e:
        return False, str(e)

    # ⭐ 关键异常分类（商业级）
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