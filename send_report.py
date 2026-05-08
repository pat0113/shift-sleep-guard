#!/usr/bin/env python3
"""
send_report.py — email the weekly shift sleep report via Maton Gmail proxy.
Reads MATON_API_KEY from environment or from ~/.openclaw/maton_key.json.
"""

import base64
import json
import os
import sys
import sys
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import requests


MATON_KEY_PATH = Path.home() / ".openclaw" / "maton_key.json"
GATEWAY_BASE   = "https://gateway.maton.ai/google-mail/gmail/v1/users/me"


def get_api_key() -> str:
    key = os.environ.get("MATON_API_KEY")
    if key:
        return key
    if MATON_KEY_PATH.exists():
        data = json.loads(MATON_KEY_PATH.read_text())
        return data.get("MATON_API_KEY", data.get("api_key", ""))
    raise RuntimeError("MATON_API_KEY not found in env or ~/_openclaw/maton_key.json")


def raw_base64url(msg: MIMEMultipart) -> str:
    raw = msg.as_bytes()
    return base64.urlsafe_b64encode(raw).decode().rstrip("=")


def send_email(to: str, subject: str, body_html: str, attachment_path: Path = None):
    key = get_api_key()
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type":  "application/json",
    }

    msg = MIMEMultipart("mixed")
    msg["To"]      = to
    msg["Subject"] = subject
    msg["From"]    = "Shift Sleep Guard <noreply@visayahealth.org>"

    body = MIMEText(body_html, "html", "utf-8")
    msg.attach(body)

    if attachment_path and attachment_path.exists():
        with attachment_path.open("rb") as f:
            part = MIMEApplication(f.read(), Name=attachment_path.name)
        part["Content-Disposition"] = f'attachment; filename="{attachment_path.name}"'
        msg.attach(part)

    payload = {"raw": raw_base64url(msg)}
    resp = requests.post(
        f"{GATEWAY_BASE}/messages/send",
        headers=headers, json=payload, timeout=30
    )
    resp.raise_for_status()
    print(f"✅ Email sent to {to}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--to",           required=True, help="Recipient email")
    parser.add_argument("--subject",       default="📋 Weekly Shift Sleep Report — Visaya Health Group")
    parser.add_argument("--report-file",   required=True, help="Path to .xlsx report")
    args = parser.parse_args()

    report_path = Path(args.report_file)
    if not report_path.exists():
        print(f"❌ Report file not found: {report_path}")
        sys.exit(1)

    html_body = """
    <html><body style="font-family: Arial, sans-serif; max-width: 600px; margin: auto;">
      <h2 style="color: #1F3864;">📋 Weekly Shift Sleep Report</h2>
      <p><strong>Visaya Health Group — Nursing Division</strong></p>
      <p>Attached is the automated weekly rest-gap and sleep risk report for this period.
         Please review flagged violations and take corrective action before the upcoming shift cycle.</p>
      <hr/>
      <p style="font-size: 0.85em; color: #666;">
        ⚠️ This is an automated message from <strong>Shift Sleep Guard</strong>.
        Do not reply to this address.<br/>
        For issues contact your unit coordinator or the scheduling team.
      </p>
    </body></html>
    """

    try:
        send_email(args.to, args.subject, html_body, report_path)
    except Exception as exc:
        print(f"❌ Failed to send email: {exc}")
        sys.exit(1)