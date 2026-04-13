import json
from sender.smtp_client import send_mail

def load_rules(path="config/rules.json"):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def process_rules(from_email, auth_code):
    rules = load_rules()

    results = []

    for rule in rules:
        success = send_mail(
            to_email=rule["email"],
            subject=rule["subject"],
            body=rule["body"],
            attachments=rule.get("attachments", []),
            from_email=from_email,
            auth_code=auth_code
        )

        results.append((rule["email"], success))

    return results