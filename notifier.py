"""
notifier.py
-------------------------------------------------------------------------
Self-contained email alerts for the standalone HSN/SAC tool. No
dependency on any other project - reads its own config.json and
subscribers.json.

Free: uses your own SMTP account (e.g. a Gmail App Password), sent via
Python's built-in smtplib. No third-party service, no per-message cost.
-------------------------------------------------------------------------
"""

import json
import os
import re
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

BASE_DIR = Path(os.environ.get("HSN_DATA_DIR", Path(__file__).parent))
CONFIG_FILE = BASE_DIR / "config.json"
SUBSCRIBERS_FILE = BASE_DIR / "subscribers.json"


def load_config():
    """Reads SMTP settings from config.json if present (normal local/Flask
    use). If config.json doesn't exist, falls back to environment
    variables SMTP_HOST/SMTP_PORT/SMTP_USER/SMTP_PASSWORD - this is what
    lets a GitHub Action supply credentials via GitHub Secrets without
    ever needing a committed config file."""
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            pass

    smtp_host = os.environ.get("SMTP_HOST")
    smtp_user = os.environ.get("SMTP_USER")
    smtp_password = os.environ.get("SMTP_PASSWORD")
    if smtp_host and smtp_user and smtp_password:
        return {
            "smtp_host": smtp_host,
            "smtp_port": int(os.environ.get("SMTP_PORT", 587)),
            "smtp_user": smtp_user,
            "smtp_password": smtp_password,
        }
    return None


def load_subscribers():
    if not SUBSCRIBERS_FILE.exists():
        return {"emails": []}
    try:
        return json.loads(SUBSCRIBERS_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"emails": []}


def save_subscribers(data):
    SUBSCRIBERS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


def add_email_subscriber(email):
    data = load_subscribers()
    if email not in data["emails"]:
        data["emails"].append(email)
        save_subscribers(data)
    return data


def remove_email_subscriber(email):
    data = load_subscribers()
    data["emails"] = [e for e in data["emails"] if e != email]
    save_subscribers(data)
    return data


def send_hsn_email_alert(summary):
    """Emails every subscribed address about a new HSN/SAC delta.
    Attaches each sheet's highlighted Excel and Added+Modified XML if
    they exist on disk. Silently no-ops if SMTP isn't configured, there
    are no subscribers, or nothing changed anywhere - safe to call after
    every check regardless of setup state."""
    sheets = summary.get("sheets", {})
    changed_sheets = {
        name: s for name, s in sheets.items()
        if not s.get("is_baseline") and sum(s.get("counts", {}).values()) > 0
    }
    if not changed_sheets:
        return {"sent": False, "reason": "no changes since last check on any sheet"}

    config = load_config()
    if not config or not config.get("smtp_host"):
        return {"sent": False, "reason": "email not configured (see config.example.json)"}

    subscribers = load_subscribers()
    recipients = subscribers.get("emails", [])
    if not recipients:
        env_recipients = os.environ.get("ALERT_EMAILS", "")
        recipients = [e.strip() for e in env_recipients.split(",") if e.strip()]
    if not recipients:
        return {"sent": False, "reason": "no subscribers"}

    total_changes = 0
    body_lines = ["HSN/SAC delta found since the last check:", ""]
    for name, s in changed_sheets.items():
        c = s["counts"]
        sheet_total = c["added"] + c["deleted"] + c["modified"]
        total_changes += sheet_total
        body_lines.append(f"[{name}] Added: {c['added']}, Deleted: {c['deleted']}, Modified: {c['modified']}")
        if s.get("counts_vs_reference"):
            r = s["counts_vs_reference"]
            body_lines.append(f"  vs reference file: Added {r['added']}, Deleted {r['deleted']}, Modified {r['modified']}")
        if s.get("validation_issue_count", 0) > 0:
            body_lines.append(f"  ⚠ {s['validation_issue_count']} field(s) with invisible characters/whitespace - see attached CSV")
        body_lines.append("")
    body_lines.append("— Sent automatically by your local HSN/SAC Delta Validator.")

    msg = MIMEMultipart()
    msg["From"] = config["smtp_user"]
    msg["Subject"] = f"HSN/SAC Dashboard: {total_changes} change(s) found across {len(changed_sheets)} sheet(s)"
    msg.attach(MIMEText("\n".join(body_lines), "plain"))

    for name in changed_sheets:
        safe = re.sub(r"\W+", "_", name).strip("_") or "sheet"
        for filename in (f"hsn_delta__{safe}.xlsx", f"hsn_delta__{safe}.xml", f"hsn_validation_issues__{safe}.csv"):
            filepath = BASE_DIR / filename
            if filepath.exists():
                part = MIMEBase("application", "octet-stream")
                part.set_payload(filepath.read_bytes())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
                msg.attach(part)

    try:
        with smtplib.SMTP(config["smtp_host"], config.get("smtp_port", 587)) as server:
            server.starttls()
            server.login(config["smtp_user"], config["smtp_password"])
            for recipient in recipients:
                msg["To"] = recipient
                server.sendmail(config["smtp_user"], recipient, msg.as_string())
        return {"sent": True, "recipients": len(recipients)}
    except Exception as exc:  # noqa: BLE001 — email failures shouldn't break the check
        return {"sent": False, "reason": str(exc)}
