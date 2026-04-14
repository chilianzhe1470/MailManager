import tkinter as tk
from tkinter import filedialog, messagebox
import os

from sender.smtp_client import send_mail
from core.rule_engine import process_rules
from receiver.imap_client import fetch_inbox

attachments = []
inbox_mails = []


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


def refresh_inbox():
    email = email_entry.get().strip()
    auth = auth_entry.get().strip()
    if not validate_basic_fields(email, auth):
        return
    limit_text = recv_limit_entry.get().strip()
    if not limit_text:
        limit = 20
    else:
        try:
            limit = int(limit_text)
        except ValueError:
            messagebox.showerror("收件失败", "拉取数量必须是整数")
            return
    if limit <= 0:
        messagebox.showerror("收件失败", "拉取数量必须大于0")
        return
    if limit > 200:
        messagebox.showerror("收件失败", "拉取数量最大为200")
        return

    success, msg, mails = fetch_inbox(email, auth, limit=limit)
    if not success:
        messagebox.showerror("收件失败", msg)
        return

    global inbox_mails
    inbox_mails = mails
    inbox_listbox.delete(0, tk.END)
    download_attachment_btn.config(state=tk.DISABLED)

    for mail in mails:
        attachment_count = len(mail.get("attachments", []))
        attachment_mark = f" | 附件{attachment_count}" if attachment_count else ""
        preview = f"{mail['from']} | {mail['subject']}{attachment_mark}"
        inbox_listbox.insert(tk.END, preview[:120])

    inbox_detail_text.delete("1.0", tk.END)
    inbox_detail_text.insert(tk.END, f"{msg}\n请选择一封邮件查看详情。")


def show_selected_mail(event):
    selection = inbox_listbox.curselection()
    if not selection:
        return
    idx = selection[0]
    if idx >= len(inbox_mails):
        return

    mail = inbox_mails[idx]
    attachment_names = [item["filename"] for item in mail.get("attachments", [])]
    attachment_text = "、".join(attachment_names) if attachment_names else "无"
    content = (
        f"发件人: {mail['from']}\n"
        f"主题: {mail['subject']}\n"
        f"附件: {attachment_text}\n"
        f"{'-' * 50}\n"
        f"{mail['body']}"
    )
    inbox_detail_text.delete("1.0", tk.END)
    inbox_detail_text.insert(tk.END, content)
    if attachment_names:
        download_attachment_btn.config(state=tk.NORMAL)
    else:
        download_attachment_btn.config(state=tk.DISABLED)


def download_selected_attachments():
    selection = inbox_listbox.curselection()
    if not selection:
        messagebox.showerror("下载失败", "请先选择一封邮件")
        return

    idx = selection[0]
    if idx >= len(inbox_mails):
        messagebox.showerror("下载失败", "所选邮件无效")
        return

    mail = inbox_mails[idx]
    items = mail.get("attachments", [])
    if not items:
        messagebox.showinfo("提示", "该邮件没有附件")
        return

    save_dir = filedialog.askdirectory(title="选择附件保存目录")
    if not save_dir:
        return

    saved_count = 0
    for item in items:
        filename = item["filename"]
        content = item["content"]
        target_path = os.path.join(save_dir, filename)

        if os.path.exists(target_path):
            name, ext = os.path.splitext(filename)
            seq = 1
            while os.path.exists(target_path):
                target_path = os.path.join(save_dir, f"{name}_{seq}{ext}")
                seq += 1

        with open(target_path, "wb") as f:
            f.write(content)
        saved_count += 1

    messagebox.showinfo("下载完成", f"已下载 {saved_count} 个附件到：\n{save_dir}")

def show_page(page_name):
    main_frame.pack_forget()
    send_frame.pack_forget()
    receive_frame.pack_forget()

    if page_name == "main":
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
    elif page_name == "send":
        send_frame.pack(fill="both", expand=True, padx=20, pady=20)
    elif page_name == "receive":
        receive_frame.pack(fill="both", expand=True, padx=20, pady=20)


root = tk.Tk()
root.title("邮件管理系统")
root.geometry("950x760")

main_frame = tk.Frame(root)
tk.Label(main_frame, text="邮件管理系统", font=("Microsoft YaHei", 18, "bold")).pack(pady=30)
tk.Button(
    main_frame,
    text="进入发件功能",
    width=20,
    height=2,
    command=lambda: show_page("send"),
    bg="green",
    fg="white"
).pack(pady=10)
tk.Button(
    main_frame,
    text="进入收件功能",
    width=20,
    height=2,
    command=lambda: show_page("receive"),
    bg="blue",
    fg="white"
).pack(pady=10)

send_frame = tk.Frame(root)
tk.Button(send_frame, text="返回主界面", command=lambda: show_page("main")).pack(anchor="w")
tk.Label(send_frame, text="发件功能", font=("Microsoft YaHei", 14, "bold")).pack(pady=8)

tk.Label(send_frame, text="邮箱").pack()
email_entry = tk.Entry(send_frame, width=50)
email_entry.pack()

tk.Label(send_frame, text="授权码").pack()
auth_entry = tk.Entry(send_frame, width=50, show="*")
auth_entry.pack()

tk.Label(send_frame, text="单个收件人").pack()
to_entry = tk.Entry(send_frame, width=50)
to_entry.pack()

tk.Label(send_frame, text="批量收件人（每行一个）").pack()
batch_to_text = tk.Text(send_frame, height=6, width=60)
batch_to_text.pack()

tk.Label(send_frame, text="主题").pack()
subject_entry = tk.Entry(send_frame, width=50)
subject_entry.pack()

tk.Label(send_frame, text="正文").pack()
body_text = tk.Text(send_frame, height=10, width=70)
body_text.pack()

tk.Button(send_frame, text="选择附件", command=select_file).pack(pady=4)
attachment_label = tk.Label(send_frame, text="", wraplength=850, justify="left")
attachment_label.pack()

tk.Button(send_frame, text="发送单封邮件", command=send_single, bg="green", fg="white").pack(pady=5)
tk.Button(send_frame, text="批量自动发送", command=send_batch, bg="blue", fg="white").pack(pady=5)

receive_frame = tk.Frame(root)
tk.Button(receive_frame, text="返回主界面", command=lambda: show_page("main")).pack(anchor="w")
tk.Label(receive_frame, text="收件功能", font=("Microsoft YaHei", 14, "bold")).pack(pady=8)

tk.Label(receive_frame, text="邮箱").pack()
recv_email_entry = tk.Entry(receive_frame, width=50)
recv_email_entry.pack()

tk.Label(receive_frame, text="授权码").pack()
recv_auth_entry = tk.Entry(receive_frame, width=50, show="*")
recv_auth_entry.pack()

tk.Label(receive_frame, text="拉取数量（1-200）").pack()
recv_limit_entry = tk.Entry(receive_frame, width=20)
recv_limit_entry.insert(0, "20")
recv_limit_entry.pack()


def refresh_inbox_by_receive_page():
    email = recv_email_entry.get().strip()
    auth = recv_auth_entry.get().strip()
    if not validate_basic_fields(email, auth):
        return
    email_entry.delete(0, tk.END)
    email_entry.insert(0, email)
    auth_entry.delete(0, tk.END)
    auth_entry.insert(0, auth)
    refresh_inbox()


tk.Button(receive_frame, text="刷新收件箱", command=refresh_inbox_by_receive_page, bg="purple", fg="white").pack(pady=6)

tk.Label(receive_frame, text="收件箱（拉取数量）").pack()
inbox_listbox = tk.Listbox(receive_frame, width=130, height=12)
inbox_listbox.pack()
inbox_listbox.bind("<<ListboxSelect>>", show_selected_mail)
inbox_listbox.bind("<ButtonRelease-1>", show_selected_mail)

download_attachment_btn = tk.Button(
    receive_frame,
    text="下载选中邮件附件",
    command=download_selected_attachments,
    state=tk.DISABLED
)
download_attachment_btn.pack(pady=5)

tk.Label(receive_frame, text="邮件详情").pack()
inbox_detail_text = tk.Text(receive_frame, height=14, width=130)
inbox_detail_text.pack()

show_page("main")
root.mainloop()