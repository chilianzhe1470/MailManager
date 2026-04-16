import logging
import os
import sys
from functools import partial

from core.rule_engine import process_rules
from receiver.imap_client import fetch_inbox
from sender.smtp_client import send_mail

try:
    from PySide6.QtCore import QObject, QRunnable, QThreadPool, Qt, Signal
    from PySide6.QtGui import QColor, QFont, QPalette, QIcon
    from PySide6.QtWidgets import (
        QApplication,
        QFileDialog,
        QFrame,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QPlainTextEdit,
        QSizePolicy,
        QSpinBox,
        QSplitter,
        QStackedWidget,
        QStatusBar,
        QTextEdit,
        QVBoxLayout,
        QWidget,
    )
except ImportError as exc:
    raise SystemExit(
        "PySide6 is not installed. Install it with: .\\.venv\\Scripts\\pip.exe install PySide6"
    ) from exc


logger = logging.getLogger(__name__)

APP_STYLE = """
QWidget {
    background: #f4f7fb;
    color: #10233d;
    font-family: "Microsoft YaHei UI";
    font-size: 14px;
}
QLabel {
    background: transparent;
}
QMainWindow {
    background: #f4f7fb;
}
QWidget#brandBlock {
    background: transparent;
}
QFrame#sidebar {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #10233d, stop:1 #163961);
    border-radius: 24px;
}
QFrame#card {
    background: #ffffff;
    border: 1px solid #dce6f3;
    border-radius: 22px;
}
QFrame#mutedCard {
    background: #edf4ff;
    border: 1px solid #d6e4fb;
    border-radius: 18px;
}
QLabel#brandTitle {
    background: transparent;
    color: white;
    font-size: 28px;
    font-weight: 700;
}
QLabel#brandSubtitle {
    background: transparent;
    color: rgba(255, 255, 255, 0.75);
    font-size: 13px;
}
QLabel#pageTitle {
    background: transparent;
    color: #0f172a;
    font-size: 28px;
    font-weight: 700;
}
QLabel#pageSubtitle {
    background: transparent;
    color: #5f6f86;
    font-size: 13px;
}
QLabel#sectionTitle {
    background: transparent;
    color: #10233d;
    font-size: 18px;
    font-weight: 700;
}
QLabel#mutedLabel {
    background: transparent;
    color: #6b7a90;
    font-size: 13px;
}
QLabel#detailLabel {
    background: transparent;
    color: #41536c;
    font-size: 13px;
}
QLineEdit, QTextEdit, QPlainTextEdit, QListWidget, QSpinBox {
    background: #fbfdff;
    border: 1px solid #cfdae8;
    border-radius: 14px;
    padding: 10px 12px;
    selection-background-color: #dbeafe;
}
QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus, QListWidget:focus, QSpinBox:focus {
    border: 1px solid #2563eb;
}
QListWidget {
    padding: 8px;
}
QPushButton {
    border: none;
    border-radius: 14px;
    padding: 12px 18px;
    font-weight: 600;
    background: #e7eef8;
    color: #153053;
}
QPushButton:hover {
    background: #dbe7f6;
}
QPushButton:disabled {
    background: #ecf1f7;
    color: #8a96a8;
}
QPushButton[role="primary"] {
    background: #2563eb;
    color: white;
}
QPushButton[role="primary"]:hover {
    background: #1e4fbd;
}
QPushButton[role="danger"] {
    background: #f97316;
    color: white;
}
QPushButton[role="danger"]:hover {
    background: #dd6410;
}
QPushButton[nav="true"] {
    text-align: left;
    padding: 14px 16px;
    color: rgba(255, 255, 255, 0.85);
    background: transparent;
    border: 1px solid transparent;
}
QPushButton[nav="true"]:hover {
    background: rgba(255, 255, 255, 0.12);
}
QPushButton[nav="true"][active="true"] {
    background: rgba(255, 255, 255, 0.18);
    border: 1px solid rgba(255, 255, 255, 0.12);
    color: white;
}
QStatusBar {
    background: transparent;
    color: #52637c;
}
"""


def make_card(object_name="card"):
    frame = QFrame()
    frame.setObjectName(object_name)
    return frame


def make_button(text, role=None):
    button = QPushButton(text)
    if role:
        button.setProperty("role", role)
        button.style().unpolish(button)
        button.style().polish(button)
    return button


def get_app_icon():
    if getattr(sys, "frozen", False):
        return QIcon(sys.executable)

    icon_path = os.path.abspath("icon.ico")
    if os.path.exists(icon_path):
        return QIcon(icon_path)
    return QIcon()


class WorkerSignals(QObject):
    result = Signal(object)
    error = Signal(str)
    finished = Signal()


class TaskWorker(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs)
        except Exception as exc:
            logger.exception("Background task failed: %s", exc)
            self.signals.error.emit(str(exc))
        else:
            self.signals.result.emit(result)
        finally:
            self.signals.finished.emit()


class MailManagerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.thread_pool = QThreadPool.globalInstance()
        self.active_workers = set()
        self.attachments = []
        self.inbox_mails = []
        self.nav_buttons = {}
        self.setWindowTitle("MailManager")
        self.setWindowIcon(get_app_icon())
        self.resize(1360, 860)
        self.setMinimumSize(1180, 760)
        self._build_ui()
        self.show_send_page()
        self._set_status("就绪")

    def _build_ui(self):
        root = QWidget()
        root_layout = QHBoxLayout(root)
        root_layout.setContentsMargins(20, 20, 20, 20)
        root_layout.setSpacing(18)

        sidebar = make_card("sidebar")
        sidebar.setFixedWidth(270)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(24, 24, 24, 24)
        sidebar_layout.setSpacing(16)

        brand_block = QWidget()
        brand_block.setObjectName("brandBlock")
        brand_layout = QVBoxLayout(brand_block)
        brand_layout.setContentsMargins(0, 0, 0, 0)
        brand_layout.setSpacing(4)

        brand_title = QLabel("MailManager")
        brand_title.setObjectName("brandTitle")
        brand_subtitle = QLabel("邮件工作台")
        brand_subtitle.setObjectName("brandSubtitle")
        brand_subtitle.setWordWrap(True)

        brand_layout.addWidget(brand_title)
        brand_layout.addWidget(brand_subtitle)

        sidebar_layout.addWidget(brand_block)
        sidebar_layout.addSpacing(10)

        self.send_nav = self._make_nav_button("发邮件", self.show_send_page)
        self.receive_nav = self._make_nav_button("收件箱", self.show_receive_page)
        sidebar_layout.addWidget(self.send_nav)
        sidebar_layout.addWidget(self.receive_nav)
        sidebar_layout.addStretch(1)

        tip_card = make_card("mutedCard")
        tip_layout = QVBoxLayout(tip_card)
        tip_layout.setContentsMargins(16, 16, 16, 16)
        tip_layout.setSpacing(6)
        tip_title = QLabel("使用提示")
        tip_title.setObjectName("sectionTitle")
        tip_text = QLabel(
            "请准备好邮箱授权码。发送和收取任务会在后台执行，"
            "这样窗口在网络请求期间也能保持响应。"
        )
        tip_text.setObjectName("mutedLabel")
        tip_text.setWordWrap(True)
        tip_layout.addWidget(tip_title)
        tip_layout.addWidget(tip_text)
        sidebar_layout.addWidget(tip_card)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 8, 0, 8)
        content_layout.setSpacing(18)

        header = QWidget()
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(6, 0, 6, 0)
        header_layout.setSpacing(2)
        self.page_title = QLabel("MailManager")
        self.page_title.setObjectName("pageTitle")
        self.page_subtitle = QLabel("更清爽的桌面邮件工作流。")
        self.page_subtitle.setObjectName("pageSubtitle")
        header_layout.addWidget(self.page_title)
        header_layout.addWidget(self.page_subtitle)

        self.stack = QStackedWidget()
        self.send_page = self._build_send_page()
        self.receive_page = self._build_receive_page()
        self.stack.addWidget(self.send_page)
        self.stack.addWidget(self.receive_page)

        content_layout.addWidget(header)
        content_layout.addWidget(self.stack, 1)

        root_layout.addWidget(sidebar)
        root_layout.addWidget(content, 1)

        self.setCentralWidget(root)
        self.setStatusBar(QStatusBar())

    def _build_send_page(self):
        page = QWidget()
        layout = QHBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(18)

        left_card = make_card()
        left_layout = QVBoxLayout(left_card)
        left_layout.setContentsMargins(24, 24, 24, 24)
        left_layout.setSpacing(14)
        left_layout.addWidget(self._section_header("账号与收件人", "填写发件邮箱信息，以及单发或群发收件人。"))

        self.send_email_input = self._make_line_edit("发件邮箱地址")
        self.send_auth_input = self._make_line_edit("邮箱授权码", echo_mode=QLineEdit.Password)
        self.single_to_input = self._make_line_edit("单个收件人邮箱")
        self.batch_to_input = QPlainTextEdit()
        self.batch_to_input.setPlaceholderText("群发时每行填写一个收件人邮箱")
        self.batch_to_input.setFixedHeight(180)

        left_layout.addWidget(self._labeled_widget("发件邮箱", self.send_email_input))
        left_layout.addWidget(self._labeled_widget("授权码", self.send_auth_input))
        left_layout.addWidget(self._labeled_widget("单个收件人", self.single_to_input))
        left_layout.addWidget(self._labeled_widget("群发收件人", self.batch_to_input))
        left_layout.addStretch(1)

        right_card = make_card()
        right_layout = QVBoxLayout(right_card)
        right_layout.setContentsMargins(24, 24, 24, 24)
        right_layout.setSpacing(14)
        right_layout.addWidget(self._section_header("邮件编辑区", "在这里填写主题、正文，并管理附件。"))

        self.subject_input = self._make_line_edit("邮件主题")
        self.body_input = QTextEdit()
        self.body_input.setPlaceholderText("在这里填写邮件正文")
        self.body_input.setMinimumHeight(260)
        self.attachment_list = QListWidget()
        self.attachment_list.setMinimumHeight(180)

        attachment_actions = QHBoxLayout()
        self.add_attachment_button = make_button("添加附件")
        self.clear_attachments_button = make_button("清空附件")
        self.add_attachment_button.clicked.connect(self.select_files)
        self.clear_attachments_button.clicked.connect(self.clear_attachments)
        attachment_actions.addWidget(self.add_attachment_button)
        attachment_actions.addWidget(self.clear_attachments_button)
        attachment_actions.addStretch(1)

        send_actions = QHBoxLayout()
        self.send_single_button = make_button("发送单封邮件", "primary")
        self.send_batch_button = make_button("群发 / 使用规则")
        self.send_single_button.clicked.connect(self.send_single)
        self.send_batch_button.clicked.connect(self.send_batch)
        send_actions.addWidget(self.send_single_button)
        send_actions.addWidget(self.send_batch_button)
        send_actions.addStretch(1)

        right_layout.addWidget(self._labeled_widget("主题", self.subject_input))
        right_layout.addWidget(self._labeled_widget("正文", self.body_input))
        right_layout.addWidget(self._labeled_widget("附件", self.attachment_list), 1)
        right_layout.addLayout(attachment_actions)
        right_layout.addLayout(send_actions)

        layout.addWidget(left_card, 4)
        layout.addWidget(right_card, 6)
        return page

    def _build_receive_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(18)

        top_card = make_card()
        top_layout = QHBoxLayout(top_card)
        top_layout.setContentsMargins(24, 20, 24, 20)
        top_layout.setSpacing(12)

        self.recv_email_input = self._make_line_edit("邮箱地址")
        self.recv_auth_input = self._make_line_edit("邮箱授权码", echo_mode=QLineEdit.Password)
        self.recv_limit_input = QSpinBox()
        self.recv_limit_input.setRange(1, 200)
        self.recv_limit_input.setValue(20)
        self.recv_limit_input.setButtonSymbols(QSpinBox.NoButtons)
        self.recv_limit_input.setFixedWidth(100)
        self.refresh_button = make_button("刷新收件箱", "primary")
        self.refresh_button.clicked.connect(self.refresh_inbox)

        top_layout.addWidget(self._labeled_widget("邮箱", self.recv_email_input))
        top_layout.addWidget(self._labeled_widget("授权码", self.recv_auth_input))
        top_layout.addWidget(self._labeled_widget("拉取数量", self.recv_limit_input))
        top_layout.addWidget(self.refresh_button, 0, Qt.AlignBottom)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)

        list_card = make_card()
        list_layout = QVBoxLayout(list_card)
        list_layout.setContentsMargins(20, 20, 20, 20)
        list_layout.setSpacing(12)
        list_layout.addWidget(self._section_header("邮件列表", "最新邮件显示在最前，点击后可查看详细内容。"))
        self.mail_list = QListWidget()
        self.mail_list.currentRowChanged.connect(self.show_selected_mail)
        list_layout.addWidget(self.mail_list, 1)

        detail_card = make_card()
        detail_layout = QVBoxLayout(detail_card)
        detail_layout.setContentsMargins(20, 20, 20, 20)
        detail_layout.setSpacing(12)
        detail_layout.addWidget(self._section_header("邮件详情", "可在下载前预览正文内容和附件名称。"))

        self.detail_from = self._detail_label("发件人", "-")
        self.detail_subject = self._detail_label("主题", "-")
        self.detail_date = self._detail_label("时间", "-")
        self.detail_body = QTextEdit()
        self.detail_body.setReadOnly(True)
        self.detail_body.setPlaceholderText("请选择一封邮件预览内容")
        self.detail_body.setMinimumHeight(220)
        self.detail_attachments = QListWidget()
        self.detail_attachments.setMinimumHeight(140)
        self.detail_attachments.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.download_attachment_button = make_button("下载当前邮件附件")
        self.download_attachment_button.setProperty("role", "danger")
        self.download_attachment_button.style().unpolish(self.download_attachment_button)
        self.download_attachment_button.style().polish(self.download_attachment_button)
        self.download_attachment_button.clicked.connect(self.download_selected_attachments)
        self.download_attachment_button.setEnabled(False)

        detail_layout.addWidget(self.detail_from)
        detail_layout.addWidget(self.detail_subject)
        detail_layout.addWidget(self.detail_date)
        detail_layout.addWidget(self._labeled_widget("正文预览", self.detail_body), 1)
        detail_layout.addWidget(self._labeled_widget("附件", self.detail_attachments))
        detail_layout.addWidget(self.download_attachment_button, 0, Qt.AlignLeft)

        splitter.addWidget(list_card)
        splitter.addWidget(detail_card)
        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 6)

        layout.addWidget(top_card)
        layout.addWidget(splitter, 1)
        return page

    def _make_nav_button(self, text, callback):
        button = QPushButton(text)
        button.setProperty("nav", "true")
        button.clicked.connect(callback)
        self.nav_buttons[text] = button
        return button

    def _section_header(self, title, text):
        wrapper = QWidget()
        layout = QVBoxLayout(wrapper)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        title_label = QLabel(title)
        title_label.setObjectName("sectionTitle")
        text_label = QLabel(text)
        text_label.setObjectName("mutedLabel")
        text_label.setWordWrap(True)
        layout.addWidget(title_label)
        layout.addWidget(text_label)
        return wrapper

    def _make_line_edit(self, placeholder, echo_mode=QLineEdit.Normal):
        line_edit = QLineEdit()
        line_edit.setPlaceholderText(placeholder)
        line_edit.setEchoMode(echo_mode)
        return line_edit

    def _labeled_widget(self, label_text, widget):
        wrapper = QWidget()
        layout = QVBoxLayout(wrapper)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        label = QLabel(label_text)
        label.setObjectName("detailLabel")
        layout.addWidget(label)
        layout.addWidget(widget)
        return wrapper

    def _detail_label(self, title, value):
        label = QLabel(f"{title}: {value}")
        label.setObjectName("detailLabel")
        label.setWordWrap(True)
        return label

    def _set_active_nav(self, active_button):
        for button in self.nav_buttons.values():
            button.setProperty("active", "true" if button is active_button else "false")
            button.style().unpolish(button)
            button.style().polish(button)

    def _set_page_header(self, title, subtitle):
        self.page_title.setText(title)
        self.page_subtitle.setText(subtitle)

    def _set_status(self, message):
        self.statusBar().showMessage(message)

    def show_send_page(self):
        self.stack.setCurrentWidget(self.send_page)
        self._set_active_nav(self.send_nav)
        self._set_page_header(
            "发邮件",
            "支持单发、群发和规则发送，执行时界面仍可保持响应。",
        )

    def show_receive_page(self):
        self.stack.setCurrentWidget(self.receive_page)
        self._set_active_nav(self.receive_nav)
        self._set_page_header(
            "收件箱",
            "拉取最近邮件、查看详情，并下载当前邮件的附件。",
        )

    def _show_error(self, message, title="错误"):
        QMessageBox.critical(self, title, message)

    def _show_info(self, message, title="提示"):
        QMessageBox.information(self, title, message)

    def _run_task(self, fn, on_success, busy_text, lock_widgets=None):
        widgets = lock_widgets or []
        for widget in widgets:
            widget.setEnabled(False)
        self._set_status(busy_text)

        worker = TaskWorker(fn)
        worker.setAutoDelete(False)
        self.active_workers.add(worker)

        def finalize():
            self._unlock_widgets(widgets)
            self._set_status("就绪")
            self.active_workers.discard(worker)

        worker.signals.result.connect(on_success)
        worker.signals.error.connect(lambda message: self._show_error(message, "任务执行失败"))
        worker.signals.finished.connect(finalize)
        self.thread_pool.start(worker)

    def _unlock_widgets(self, widgets):
        for widget in widgets:
            widget.setEnabled(True)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择附件")
        if not files:
            return
        existing = set(self.attachments)
        for file_path in files:
            if file_path not in existing:
                self.attachments.append(file_path)
                existing.add(file_path)
        self._refresh_attachment_list()
        self._set_status(f"已加载 {len(self.attachments)} 个附件")

    def clear_attachments(self):
        self.attachments.clear()
        self._refresh_attachment_list()
        self._set_status("已清空附件")

    def _refresh_attachment_list(self):
        self.attachment_list.clear()
        for path in self.attachments:
            item = QListWidgetItem(os.path.basename(path))
            item.setToolTip(path)
            self.attachment_list.addItem(item)

    def _validate_send_inputs(self):
        email = self.send_email_input.text().strip()
        auth = self.send_auth_input.text().strip()
        if not email:
            self._show_error("请输入发件邮箱地址。")
            return None
        if not auth:
            self._show_error("请输入邮箱授权码。")
            return None
        return email, auth

    def _validate_receive_inputs(self):
        email = self.recv_email_input.text().strip()
        auth = self.recv_auth_input.text().strip()
        if not email:
            self._show_error("请输入邮箱地址。")
            return None
        if not auth:
            self._show_error("请输入邮箱授权码。")
            return None
        return email, auth, int(self.recv_limit_input.value())

    def send_single(self):
        validated = self._validate_send_inputs()
        if not validated:
            return

        to_email = self.single_to_input.text().strip()
        if not to_email:
            self._show_error("请输入单个收件人邮箱地址。")
            return

        email, auth = validated
        subject = self.subject_input.text().strip()
        body = self.body_input.toPlainText().strip()
        attachments = list(self.attachments)

        self._run_task(
            lambda: send_mail(
                to_email=to_email,
                subject=subject,
                body=body,
                attachments=attachments,
                from_email=email,
                auth_code=auth,
            ),
            self._handle_send_single_result,
            "正在发送单封邮件...",
            lock_widgets=[self.send_single_button, self.send_batch_button],
        )

    def _handle_send_single_result(self, result):
        success, message = result
        if success:
            self._show_info(message, "发送成功")
        else:
            self._show_error(message, "发送失败")

    def send_batch(self):
        validated = self._validate_send_inputs()
        if not validated:
            return

        email, auth = validated
        subject = self.subject_input.text().strip()
        body = self.body_input.toPlainText().strip()
        attachments = list(self.attachments)
        recipients = [
            line.strip()
            for line in self.batch_to_input.toPlainText().splitlines()
            if line.strip()
        ]

        def batch_job():
            if recipients:
                results = []
                for to_email in recipients:
                    ok, message = send_mail(
                        to_email=to_email,
                        subject=subject,
                        body=body,
                        attachments=attachments,
                        from_email=email,
                        auth_code=auth,
                    )
                    results.append((to_email, ok, message))
                return results
            return process_rules(email, auth)

        self._run_task(
            batch_job,
            self._handle_send_batch_result,
            "正在执行群发任务...",
            lock_widgets=[self.send_single_button, self.send_batch_button],
        )

    def _handle_send_batch_result(self, results):
        summary_lines = []
        success_count = 0
        for to_email, success, message in results:
            if success:
                success_count += 1
            summary_lines.append(
                f"{to_email} -> {'成功' if success else '失败'} | {message}"
            )

        summary = "\n".join(summary_lines) if summary_lines else "没有发送任何邮件。"
        title = f"群发完成（{success_count}/{len(results)}）"
        self._show_info(summary, title)

    def refresh_inbox(self):
        validated = self._validate_receive_inputs()
        if not validated:
            return

        email, auth, limit = validated

        self._run_task(
            lambda: fetch_inbox(email, auth, limit=limit),
            self._handle_refresh_result,
            "正在刷新收件箱...",
            lock_widgets=[self.refresh_button],
        )

    def _handle_refresh_result(self, result):
        success, message, mails = result
        if not success:
            self._show_error(message, "刷新收件箱失败")
            return

        self.inbox_mails = mails
        self.mail_list.clear()
        self.detail_attachments.clear()
        self.detail_body.clear()
        self.download_attachment_button.setEnabled(False)
        self.detail_from.setText("发件人: -")
        self.detail_subject.setText("主题: -")
        self.detail_date.setText("时间: -")

        for mail in mails:
            attachment_count = len(mail.get("attachments", []))
            summary = (
                f"{mail['from']}\n"
                f"{mail['subject']}\n"
                f"{mail['date']}  |  附件 {attachment_count}"
            )
            item = QListWidgetItem(summary)
            item.setToolTip(mail["subject"])
            self.mail_list.addItem(item)

        if mails:
            self.mail_list.setCurrentRow(0)
        self._set_status(message)

    def show_selected_mail(self, row):
        if row < 0 or row >= len(self.inbox_mails):
            self.detail_from.setText("发件人: -")
            self.detail_subject.setText("主题: -")
            self.detail_date.setText("时间: -")
            self.detail_body.clear()
            self.detail_attachments.clear()
            self.download_attachment_button.setEnabled(False)
            return

        mail = self.inbox_mails[row]
        attachments = mail.get("attachments", [])
        self.detail_from.setText(f"发件人: {mail.get('from', '-')}")
        self.detail_subject.setText(f"主题: {mail.get('subject', '-')}")
        self.detail_date.setText(f"时间: {mail.get('date', '-')}")
        self.detail_body.setPlainText(mail.get("body", ""))
        self.detail_attachments.clear()
        for item in attachments:
            name = item.get("filename", "attachment.bin")
            attachment_item = QListWidgetItem(name)
            attachment_item.setToolTip(name)
            self.detail_attachments.addItem(attachment_item)
        self.download_attachment_button.setEnabled(bool(attachments))

    def download_selected_attachments(self):
        row = self.mail_list.currentRow()
        if row < 0 or row >= len(self.inbox_mails):
            self._show_error("请先选择一封邮件。")
            return

        mail = self.inbox_mails[row]
        items = mail.get("attachments", [])
        if not items:
            self._show_info("当前邮件没有附件。")
            return

        save_dir = QFileDialog.getExistingDirectory(self, "选择附件保存目录")
        if not save_dir:
            return

        saved_count = 0
        failed_items = []
        for item in items:
            filename = item.get("filename", "attachment.bin")
            content = item.get("content", b"")
            target_path = self._resolve_unique_path(save_dir, filename)

            try:
                if not isinstance(content, (bytes, bytearray)):
                    content = str(content).encode("utf-8", errors="replace")

                with open(target_path, "wb") as file_obj:
                    file_obj.write(content)
                saved_count += 1
            except Exception as exc:
                failed_items.append((filename, str(exc)))
                logger.exception("Failed to save attachment %s: %s", filename, exc)

        if failed_items:
            details = "\n".join(f"- {name}: {error}" for name, error in failed_items[:5])
            self._show_error(
                f"已保存 {saved_count} 个附件，但仍有 {len(failed_items)} 个保存失败。\n\n{details}",
                "附件下载部分失败",
            )
            return

        self._show_info(
            f"已将 {saved_count} 个附件下载到：\n{save_dir}",
            "附件下载完成",
        )

    def _resolve_unique_path(self, save_dir, filename):
        target_path = os.path.join(save_dir, filename)
        if not os.path.exists(target_path):
            return target_path

        name, ext = os.path.splitext(filename)
        index = 1
        while True:
            candidate = os.path.join(save_dir, f"{name}_{index}{ext}")
            if not os.path.exists(candidate):
                return candidate
            index += 1

    def closeEvent(self, event):
        logging.shutdown()
        self._clear_log_file()
        super().closeEvent(event)

    def _clear_log_file(self):
        log_path = os.path.join("logs", "mail.log")
        try:
            os.makedirs("logs", exist_ok=True)
            with open(log_path, "w", encoding="utf-8"):
                pass
        except Exception as exc:
            print(f"Failed to clear log file: {exc}")


def build_application():
    app = QApplication(sys.argv)
    app.setStyleSheet(APP_STYLE)
    app.setFont(QFont("Microsoft YaHei UI", 10))
    app.setWindowIcon(get_app_icon())

    palette = app.palette()
    palette.setColor(QPalette.Window, QColor("#f4f7fb"))
    app.setPalette(palette)

    window = MailManagerWindow()
    # Keep a strong reference to the main window so it does not get
    # garbage-collected right after startup in packaged builds.
    app.main_window = window
    screen = app.primaryScreen()
    if screen is not None:
        available = screen.availableGeometry()
        x = available.x() + max(0, (available.width() - window.width()) // 2)
        y = available.y() + max(0, (available.height() - window.height()) // 2)
        window.move(x, y)
    window.show()
    return app


def main():
    app = build_application()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
