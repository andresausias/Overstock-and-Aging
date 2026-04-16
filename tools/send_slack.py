"""
send_slack.py
-------------
Posts a KPI summary message to a Slack channel with Drive links to all outputs.
Reads aging KPIs from the latest filtered CSV and run_params.json.

Usage:
    python tools/send_slack.py

Required .env keys:
    SLACK_BOT_TOKEN     xoxb-... bot token
    SLACK_CHANNEL_ID    Channel to post to (e.g. C01234567)

Required Drive folder IDs in .env (for generating links):
    GOOGLE_DRIVE_OUTPUT_FINALIZED_ID
    GOOGLE_DRIVE_OUTPUT_SLIDES_ID
"""

import os
import re
import sys
from pathlib import Path
from datetime import datetime
import json

from dotenv import load_dotenv

try:
    from slack_sdk import WebClient
    from slack_sdk.errors import SlackApiError
except ImportError:
    print("slack_sdk not installed. Run: pip install slack_sdk")
    sys.exit(1)

load_dotenv()

ROOT       = Path(__file__).parent.parent
FILTERED_DIR = ROOT / ".tmp" / "filtered"
CONFIG_DIR   = Path(__file__).parent / "config"
REPORTS_DIR  = ROOT / ".tmp" / "reports"
SLIDES_DIR   = ROOT / ".tmp" / "slides"


def extract_aging_kpis() -> dict:
    """Extract aged KPIs from latest filtered CSV."""
    import pandas as pd
    files = sorted([
        f for f in FILTERED_DIR.glob("*_filtered.csv")
        if re.match(r"^\d{6}_Aging", f.name, re.IGNORECASE)
    ])
    if not files:
        return {"aged_value": 0, "aged_units": 0, "aged_pct": 0}

    df = pd.read_csv(files[-1])
    df["Total Amount $"] = pd.to_numeric(df["Total Amount $"], errors="coerce").fillna(0)
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    aged = df[df["Range TOTAL"].astype(str).str.strip() == "Over 365"]
    val   = aged["Total Amount $"].sum()
    units = int(aged["Qty"].sum())

    params = json.loads((CONFIG_DIR / "run_params.json").read_text())
    denom  = params.get("total_inventory_denominator", 50_935_732)
    return {
        "aged_value": val,
        "aged_units": units,
        "aged_pct": val / denom * 100 if denom else 0,
    }


def build_drive_link(folder_id: str, label: str) -> str:
    return f"<https://drive.google.com/drive/folders/{folder_id}|{label}>"


def build_message(aging: dict, params: dict, finalized_id: str, slides_id: str) -> list:
    """Build Slack Block Kit message payload."""
    date_str = params.get("report_date", datetime.now().strftime("%b %d, %Y"))
    wk       = params.get("week_current", "?")
    ov_val   = params.get("overstock_valuation", 0)
    ov_units = params.get("overstock_units", 0)
    ov_pct   = ov_val / params.get("total_inventory_denominator", 50_935_732) * 100

    reports_link = build_drive_link(finalized_id, "📊 View Reports") if finalized_id else "_(no link)_"
    slides_link  = build_drive_link(slides_id, "📑 View Slides") if slides_id else "_(no link)_"

    # Latest report files
    html_reports = sorted(REPORTS_DIR.glob("*.html")) if REPORTS_DIR.exists() else []
    pptx_files   = sorted(SLIDES_DIR.glob("*.pptx")) if SLIDES_DIR.exists() else []
    latest_report = html_reports[-1].name if html_reports else "—"
    latest_slide  = pptx_files[-1].name if pptx_files else "—"

    exec_bullets = params.get("exec_summary_bullets", [])

    blocks = [
        {
            "type": "header",
            "text": {"type": "plain_text", "text": f"📦 Weekly Inventory Update — {date_str} (WK{wk})"}
        },
        {"type": "divider"},
        {
            "type": "section",
            "fields": [
                {"type": "mrkdwn", "text": f"*🏷️ Overstock*\n${ov_val/1e6:.1f}M | {ov_pct:.1f}% of Total | {ov_units:,} units"},
                {"type": "mrkdwn", "text": f"*⏳ Aging (Over 365)*\n${aging['aged_value']/1e6:.2f}M | {aging['aged_pct']:.1f}% of Total | {aging['aged_units']:,} units"},
            ]
        },
        {"type": "divider"},
    ]

    if exec_bullets:
        bullet_text = "\n".join(f"• {b}" for b in exec_bullets)
        blocks.append({
            "type": "section",
            "text": {"type": "mrkdwn", "text": f"*Executive Summary:*\n{bullet_text}"}
        })
        blocks.append({"type": "divider"})

    blocks.append({
        "type": "section",
        "fields": [
            {"type": "mrkdwn", "text": f"*Latest Report:*\n`{latest_report}`"},
            {"type": "mrkdwn", "text": f"*Latest Slides:*\n`{latest_slide}`"},
        ]
    })
    blocks.append({
        "type": "actions",
        "elements": [
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "View Reports"},
                "url": f"https://drive.google.com/drive/folders/{finalized_id}" if finalized_id else "#",
                "style": "primary"
            },
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "View Slides"},
                "url": f"https://drive.google.com/drive/folders/{slides_id}" if slides_id else "#",
            }
        ]
    })
    blocks.append({
        "type": "context",
        "elements": [
            {"type": "mrkdwn", "text": f"Generated by Overstock & Aging Automation · {datetime.now().strftime('%Y-%m-%d %H:%M')}"}
        ]
    })
    return blocks


def main():
    token      = os.getenv("SLACK_BOT_TOKEN")
    channel_id = os.getenv("SLACK_CHANNEL_ID")
    finalized_id = os.getenv("GOOGLE_DRIVE_OUTPUT_FINALIZED_ID", "")
    slides_id    = os.getenv("GOOGLE_DRIVE_OUTPUT_SLIDES_ID", "")

    if not token or not channel_id:
        print("ERROR: SLACK_BOT_TOKEN and SLACK_CHANNEL_ID must be set in .env")
        sys.exit(1)

    try:
        import pandas as pd
    except ImportError:
        print("pandas not installed. Run: pip install pandas")
        sys.exit(1)

    print("Extracting aging KPIs...")
    aging = extract_aging_kpis()

    params = json.loads((CONFIG_DIR / "run_params.json").read_text())

    print("Building Slack message...")
    blocks = build_message(aging, params, finalized_id, slides_id)

    print(f"Sending to channel {channel_id}...")
    client = WebClient(token=token)
    try:
        response = client.chat_postMessage(
            channel=channel_id,
            text=f"Weekly Inventory Update — {params.get('report_date', '')} (WK{params.get('week_current', '?')})",
            blocks=blocks,
        )
        print(f"✓ Message sent. Timestamp: {response['ts']}")
    except SlackApiError as e:
        print(f"Slack API error: {e.response['error']}")
        sys.exit(1)


if __name__ == "__main__":
    main()
