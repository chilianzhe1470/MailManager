import tkinter as tk
from tkinter import filedialog, messagebox

from sender.smtp_client import send_mail
from core.rule_engine import process_rules

attachments = []


def validate_basic_fields(email, auth):
    if not email.strip():
        messagebox.showerror("失败", "发件人邮箱不能为空")
        return False
    if not auth.strip():
        messagebox.showerror("失败", "授权码不能为空")
        return False
    return True


def select_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        attachments.append(file_path)
        attachment_label.config(text="\n".join(attachments))

def send_single():
    email = email_entry.get().strip()
    auth = auth_entry.get().strip()
    to_email = to_entry.get().strip()
    subject = subject_entry.get().strip()
    body = body_text.get("1.0", tk.END).strip()

    if not validate_basic_fields(email, auth):
        return
    if not to_email:
        messagebox.showerror("失败", "收件人不能为空")
        return

    success, msg = send_mail(
        to_email=to_email,
        subject=subject,
        body=body,
        attachments=attachments,
        from_email=email,
        auth_code=auth
    )

    if success:
        messagebox.showinfo("成功", f"邮件发送成功\n{msg}")
    else:
        messagebox.showerror("失败", f"邮件发送失败\n{msg}")

def send_batch():
    email = email_entry.get().strip()
    auth = auth_entry.get().strip()
    subject = subject_entry.get().strip()
    body = body_text.get("1.0", tk.END).strip()
    batch_lines = batch_to_text.get("1.0", tk.END).splitlines()
    batch_recipients = [line.strip() for line in batch_lines if line.strip()]

    if not validate_basic_fields(email, auth):
        return

    if batch_recipients:
        results = []
        for to_email in batch_recipients:
            ok, message = send_mail(
                to_email=to_email,
                subject=subject,
                body=body,
                attachments=attachments,
                from_email=email,
                auth_code=auth
            )
            results.append((to_email, ok, message))
    else:
        results = process_rules(email, auth)

    msg = ""
    for r in results:
        msg += f"{r[0]} -> {'成功' if r[1] else '失败'} | {r[2]}\n"

    messagebox.showinfo("批量发送结果", msg)

root = tk.Tk()
root.title("邮件管理系统")
root.geometry("500x600")

tk.Label(root, text="邮箱").pack()
email_entry = tk.Entry(root, width=40)
email_entry.pack()

tk.Label(root, text="授权码").pack()
auth_entry = tk.Entry(root, width=40)
auth_entry.pack()

tk.Label(root, text="单个收件人").pack()
to_entry = tk.Entry(root, width=40)
to_entry.pack()

tk.Label(root, text="批量收件人（每行一个）").pack()
batch_to_text = tk.Text(root, height=6)
batch_to_text.pack()

tk.Label(root, text="主题").pack()
subject_entry = tk.Entry(root, width=40)
subject_entry.pack()

tk.Label(root, text="正文").pack()
body_text = tk.Text(root, height=8)
body_text.pack()

tk.Button(root, text="选择附件", command=select_file).pack()
attachment_label = tk.Label(root, text="")
attachment_label.pack()

tk.Button(root, text="发送单封邮件", command=send_single, bg="green", fg="white").pack(pady=5)
tk.Button(root, text="批量自动发送", command=send_batch, bg="blue", fg="white").pack(pady=5)

root.mainloop()