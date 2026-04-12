import smtplib
from email.mime.text import MIMEText

def send_mail(to_email):
    msg = MIMEText("Hello from PyCharm")
    msg['Subject'] = "Test"
    msg['From'] = "2032009168@qq.com"
    msg['To'] = "@qq.com"

    server = smtplib.SMTP("smtp.qq.com", 587)
    server.starttls()
    server.login("2032009168@qq.com", "qgukpusxavypeecg")
    server.sendmail(
        msg['From'],
        msg['To'],
        msg.as_string()
    )
    server.quit()