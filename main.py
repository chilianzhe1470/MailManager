from sender.smtp_client import send_mail

ok, msg = send_mail(
    to_email="2032009168@qq.com",
    subject="多附件测试",
    body="这是一个带多个附件的邮件",
    attachments=[
        "attachments/a.pdf",
        "attachments/b.pdf"
    ],
    from_email="your_email@qq.com",
    auth_code="your_auth_code"
)

print(ok, msg)