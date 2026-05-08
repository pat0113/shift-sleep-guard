#!/usr/bin/env python3
"""
run_weekly_report.py — run the full pipeline: analyze + email.
Intended to be called by cron every Sunday at 23:00 São Paulo time.
Reads config from config.ini.
"""

import configparser
import subprocess
import sys
from pathlib import Path

import requests


MATON_KEY_PATH = Path.home() / ".openclaw" / "maton_key.json"
GATEWAY_BASE   = "https://gateway.maton.ai/google-mail/gmail/v1/users/me"


def get_maton_key():
    import json, os
    k = os.environ.get("MATON_API_KEY")
    if k:
        return k
    if MATON_KEY_PATH.exists():
        d = json.loads(MATON_KEY_PATH.read_text())
        return d.get("MATON_API_KEY", d.get("api_key", ""))
    raise RuntimeError("MATON_API_KEY not set")


def send_email(to: str, subject: str, html_body: str, xlsx_path: Path):
    """Send report via Maton Gmail proxy with xlsx attachment."""
    import json, base64
    from email.mime.multipart import MIMEMultipart
    from email.mime.text      import MIMEText
    from email.mime.application import MIMEApplication

    key    = get_maton_key()
    header = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}

    msg = MIMEMultipart("mixed")
    msg["To"]      = to
    msg["Subject"] = subject
    msg["From"]    = "Shift Sleep Guard <noreply@visayahealth.org>"
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    if xlsx_path.exists():
        with xlsx_path.open("rb") as f:
            part = MIMEApplication(f.read(), Name=xlsx_path.name)
        part["Content-Disposition"] = f'attachment; filename="{xlsx_path.name}"'
        msg.attach(part)

    raw_b64 = base64.urlsafe_b64encode(msg.as_bytes()).decode().rstrip("=")
    resp = requests.post(
        f"{GATEWAY_BASE}/messages/send",
        headers=header, json={"raw": raw_b64}, timeout=30
    )
    resp.raise_for_status()
    print(f"✅ Report emailed to {to}")


if __name__ == "__main__":
    cfg = configparser.ConfigParser()
    cfg.read("config.ini")

    csv_path = cfg["paths"]["csv_path"]
    out_xlsx = cfg["paths"]["output_xlsx"]
    to_addr  = cfg["email"]["to_address"]

    print("─" * 50)
    print("  Shift Sleep Guard — Weekly Pipeline")
    print("─" * 50)

    # 1. Run analysis
    result = subprocess.run(
        [sys.executable, "shift_sleep_guard.py",
         "--csv", csv_path, "--out", out_xlsx],
        check=False
    )
    if result.returncode:
        print("❌ Analysis failed")
        sys.exit(1)

    # 2. Count violations for email summary
    import pandas as pd
    df = pd.read_excel(out_xlsx, sheet_name="Violations")
    total   = len(df)
    viol    = int((df["Rest Status"].str.startswith("🔴")).sum()) if total else 0
    mod     = int((df["Rest Status"].str.startswith("🟡")).sum()) if total else 0

    subject = f"📋 Shift Sleep Report — {viol} HIGH risk | {mod} MODERATE"

    html = f"""
    <html><body style="font-family:Arial,sans-serif;max-width:620px;margin:auto">
      <h2 style="color:#1F3864;">📋 Weekly Shift Sleep Report</h2>
      <p><strong>Visaya Health Group — Nursing Division</strong></p>
      <p>This automated report covers the most recent scheduling period.</p>
      <table style="border-collapse:collapse;width:100%">
        <tr>
          <th style="border:1px solid #ccc;padding:8px;text-align:left">Metric</th>
          <th style="border:1px solid #ccc;padding:8px;text-align:center">Count</th>
        </tr>
        <tr style="background:#FFF0F0">
          <td style="border:1px solid #ccc;padding:8px">🔴 HIGH risk violations</td>
          <td style="border:1px solid #ccc;padding:8px;text-align:center;font-weight:bold">{viol}</td>
        </tr>
        <tr style="background:#FFFBE6">
          <td style="border:1px solid #ccc;padding:8px">🟡 MODERATE risk violations</td>
          <td style="border:1px solid #ccc;padding:8px;text-align:center;font-weight:bold">{mod}</td>
        </tr>
        <tr>
          <td style="border:1px solid #ccc;padding:8px">Total gap records reviewed</td>
          <td style="border:1px solid #ccc;padding:8px;text-align:center">{total}</td>
        </tr>
      </table>
      <p style="margin-top:16px">The full spreadsheet with per-nurse scores and conditional formatting is attached.
         Please review highlighted rows before the upcoming shift cycle.</p>
      <hr/>
      <p style="font-size:0.8em;color:#666">⚠️ Automated message from Shift Sleep Guard.
         Do not reply. Contact your unit coordinator for scheduling corrections.</p>
    </body></html>
    """

    send_email(to_addr, subject, html, Path(out_xlsx))
    print("✅ Pipeline complete")