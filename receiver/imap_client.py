import imaplib
from email import message_from_bytes
from email.header import decode_header
from email.utils import parsedate_to_datetime
import re
from datetime import datetime
import logging
import os


IMAP_CONFIG = {
    "qq.com": ("imap.qq.com", 993),
    "gmail.com": ("imap.gmail.com", 993),
    "163.com": ("imap.163.com", 993),
    "126.com": ("imap.126.com", 993),
    "outlook.com": ("outlook.office365.com", 993),
    "hotmail.com": ("outlook.office365.com", 993),
    "live.com": ("outlook.office365.com", 993),
}

os.makedirs("logs", exist_ok=True)
logger = logging.getLogger(__name__)


def _decode_header_value(raw_value):
    if not raw_value:
        return ""
    parts = decode_header(raw_value)
    decoded = []
    for text, charset in parts:
        if isinstance(text, bytes):
            decoded.append(text.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(text)
    return "".join(decoded)


def _resolve_imap_server(from_email):
    if not from_email or "@" not in from_email:
        return None, None, "发件人邮箱格式不正确"
    domain = from_email.split("@", 1)[1].lower()
    host_port = IMAP_CONFIG.get(domain)
    if not host_port:
        return None, None, f"暂不支持自动识别 {domain} 的IMAP服务器"
    return host_port[0], host_port[1], None


def _extract_body(mail_obj):
    if mail_obj.is_multipart():
        for part in mail_obj.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))
            if "attachment" in content_disposition.lower():
                continue
            if content_type == "text/plain":
                payload = part.get_payload(decode=True)
                if payload is None:
                    continue
                charset = part.get_content_charset() or "utf-8"
                return payload.decode(charset, errors="replace").strip()
    else:
        payload = mail_obj.get_payload(decode=True)
        if payload:
            charset = mail_obj.get_content_charset() or "utf-8"
            return payload.decode(charset, errors="replace").strip()
    return ""


def _safe_filename(name):
    cleaned = re.sub(r'[\\/:*?"<>|]', "_", name).strip()
    return cleaned or "attachment.bin"


def _extract_attachments(mail_obj, mail_uid=None):
    attachments = []
    for idx, part in enumerate(mail_obj.walk()):
        if part.is_multipart():
            continue
        content_disposition = str(part.get("Content-Disposition", ""))
        filename = part.get_filename() or part.get_param("name")
        content_type = part.get_content_type()
        disposition_lower = content_disposition.lower()
        has_file_hint = bool(filename)
        is_attachment_like = (
            "attachment" in disposition_lower
            or ("inline" in disposition_lower and has_file_hint)
            or has_file_hint
        )
        if not is_attachment_like:
            logger.info(
                "附件识别跳过 uid=%s part=%s content_type=%s disposition=%s filename=%s",
                mail_uid,
                idx,
                content_type,
                content_disposition,
                filename,
            )
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            raw_payload = part.get_payload()
            if isinstance(raw_payload, bytes):
                payload = raw_payload
            elif isinstance(raw_payload, str):
                payload = raw_payload.encode("utf-8", errors="replace")
        if payload is None:
            logger.warning(
                "附件内容为空 uid=%s part=%s content_type=%s disposition=%s filename=%s",
                mail_uid,
                idx,
                content_type,
                content_disposition,
                filename,
            )
            continue
        decoded_name = _decode_header_value(filename) if filename else "attachment.bin"
        safe_name = _safe_filename(decoded_name)
        if len(payload) == 0:
            logger.info(
                "附件为空文件 uid=%s part=%s filename=%s content_type=%s",
                mail_uid,
                idx,
                safe_name,
                content_type,
            )
        attachments.append(
            {
                "filename": safe_name,
                "content": payload,
            }
        )
        logger.info(
            "附件识别成功 uid=%s part=%s filename=%s bytes=%s content_type=%s",
            mail_uid,
            idx,
            safe_name,
            len(payload),
            content_type,
        )
    return attachments


def _extract_internal_date(fetch_blocks):
    for block in fetch_blocks:
        if isinstance(block, tuple) and len(block) >= 1:
            meta = block[0]
            if isinstance(meta, bytes):
                text = meta.decode("utf-8", errors="replace")
                m = re.search(r'INTERNALDATE "([^"]+)"', text)
                if m:
                    raw = m.group(1)
                    try:
                        dt = datetime.strptime(raw, "%d-%b-%Y %H:%M:%S %z")
                        return dt.strftime("%Y-%m-%d %H:%M:%S")
                    except Exception:
                        return raw
    return ""


def _extract_raw_message(fetch_blocks):
    for block in fetch_blocks:
        if isinstance(block, tuple) and len(block) >= 2 and isinstance(block[1], bytes):
            return block[1]
    return b""


def fetch_inbox(from_email, auth_code, limit=20):
    host, port, err = _resolve_imap_server(from_email)
    if err:
        return False, err, []
    if not auth_code:
        return False, "授权码不能为空", []

    try:
        logger.info(
            "开始收件 from=%s host=%s port=%s limit=%s",
            from_email,
            host,
            port,
            limit,
        )
        with imaplib.IMAP4_SSL(host, port, timeout=20) as imap:
            imap.login(from_email, auth_code)
            status, _ = imap.select("INBOX")
            if status != "OK":
                logger.error("打开INBOX失败 status=%s", status)
                return False, "无法打开收件箱", []

            status, data = imap.search(None, "ALL")
            if status != "OK":
                logger.error("搜索邮件失败 status=%s data=%s", status, data)
                return False, "搜索邮件失败", []

            uids = data[0].split()
            logger.info("搜索到邮件数量 total=%s", len(uids))
            if not uids:
                return True, "收件箱为空", []

            recent_uids = uids[-limit:]
            recent_uids.reverse()

            mails = []
            for uid in recent_uids:
                uid_text = uid.decode("utf-8", errors="replace")
                status, msg_data = imap.fetch(uid, "(RFC822 INTERNALDATE)")
                if status != "OK" or not msg_data or msg_data[0] is None:
                    logger.warning("拉取邮件失败 uid=%s status=%s", uid_text, status)
                    continue
                raw = _extract_raw_message(msg_data)
                if not raw:
                    logger.warning("邮件原文为空 uid=%s", uid_text)
                    continue
                mail_obj = message_from_bytes(raw)

                subject = _decode_header_value(mail_obj.get("Subject"))
                sender = _decode_header_value(mail_obj.get("From"))
                date_raw = mail_obj.get("Date", "")
                date_text = date_raw
                try:
                    if date_raw:
                        date_text = parsedate_to_datetime(date_raw).strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    pass
                if not date_text:
                    date_text = _extract_internal_date(msg_data)
                if not date_text:
                    date_text = "(未知时间)"

                body = _extract_body(mail_obj)
                mail_attachments = _extract_attachments(mail_obj, mail_uid=uid_text)
                logger.info(
                    "邮件解析完成 uid=%s from=%s subject=%s attachments=%s body_len=%s",
                    uid_text,
                    sender or "(未知发件人)",
                    subject or "(无主题)",
                    len(mail_attachments),
                    len(body or ""),
                )
                mails.append(
                    {
                        "subject": subject or "(无主题)",
                        "from": sender or "(未知发件人)",
                        "date": date_text,
                        "body": body or "(正文为空或仅HTML内容)",
                        "attachments": mail_attachments,
                    }
                )

            logger.info("收件完成 successfully_parsed=%s", len(mails))
            return True, f"成功拉取 {len(mails)} 封邮件", mails
    except imaplib.IMAP4.error as e:
        logger.error("IMAP认证或协议错误: %s", e)
        return False, f"IMAP认证或协议错误: {e}", []
    except Exception as e:
        logger.exception("收取失败: %s", e)
        return False, f"收取失败: {e}", []
