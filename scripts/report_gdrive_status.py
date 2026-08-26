"""CLI: 列出各監看試算表最新檔與台北時間；可寫入 Actions summary，有 SMTP 機密才寄信。"""
import os
import smtplib
import sys
from email.message import EmailMessage

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from utils import collect_tracked_file_status, format_status_markdown  # noqa: E402


def _maybe_email(body):
    to_addr = (os.environ.get("STATUS_MAIL_TO") or "").strip()
    host = (os.environ.get("SMTP_HOST") or "").strip()
    if not to_addr or not host:
        print("未設定 STATUS_MAIL_TO / SMTP_HOST，略過寄信。")
        return False
    port = int(os.environ.get("SMTP_PORT") or "587")
    user = (os.environ.get("SMTP_USER") or "").strip()
    password = os.environ.get("SMTP_PASSWORD") or ""
    from_addr = (os.environ.get("SMTP_FROM") or user or to_addr).strip()
    msg = EmailMessage()
    msg["Subject"] = "雲端試算表最新修改時間"
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg.set_content(body)
    with smtplib.SMTP(host, port, timeout=30) as smtp:
        smtp.starttls()
        if user:
            smtp.login(user, password)
        smtp.send_message(msg)
    print(f"已寄出狀態信至 {to_addr}")
    return True


def main():
    rows = collect_tracked_file_status()
    body = format_status_markdown(rows)
    print(body)
    summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary_path:
        with open(summary_path, "a", encoding="utf-8") as f:
            f.write(body)
            f.write("\n")
    _maybe_email(body)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
