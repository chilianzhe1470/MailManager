import tkinter as tk
from tkinter import filedialog, messagebox

from sender.smtp_client import send_mail
from core.rule_engine import process_rules

attachments = []

def select_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        attachments.append(file_path)
        attachment_label.config(text="\n".join(attachments))

def send_single():
    email = email_entry.get()
    auth = auth_entry.get()
    to_email = to_entry.get()
    subject = subject_entry.get()
    body = body_text.get("1.0", tk.END)

    success ,msg = send_mail(
        to_email=to_email,
        subject=subject,
        body=body,
        attachments=attachments,
        from_email=email,
        auth_code=auth
    )

    if success:
        messagebox.showinfo("成功", "邮件发送成功")
    else:
        messagebox.showerror("失败", "邮件发送失败")

def send_batch():
    email = email_entry.get()
    auth = auth_entry.get()

    results = process_rules(email, auth)

    msg = ""
    for r in results:
        msg += f"{r[0]} -> {'成功' if r[1] else '失败'}\n"

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

tk.Label(root, text="收件人").pack()
to_entry = tk.Entry(root, width=40)
to_entry.pack()

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