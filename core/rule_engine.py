import json
import sys
import os
from sender.smtp_client import send_mail

def get_resource_path(relative_path):
    """获取资源文件的绝对路径（支持打包后的exe）"""
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def load_rules(path="config/rules.json"):
    full_path = get_resource_path(path)
    with open(full_path, "r", encoding="utf-8") as f:
        return json.load(f)

def process_rules(from_email, auth_code):
    rules = load_rules()

    results = []

    for rule in rules:
        ok, message = send_mail(
            to_email=rule["email"],
            subject=rule["subject"],
            body=rule["body"],
            attachments=rule.get("attachments", []),
            from_email=from_email,
            auth_code=auth_code
        )

        results.append((rule["email"], ok, message))

    return results