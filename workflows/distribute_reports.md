# Workflow: Distribute Reports

## Objective
Upload all outputs to Google Drive and send a Slack summary notification.

## Prerequisites
- All HTML reports generated (`.tmp/reports/`)
- PPTX slides generated (`.tmp/slides/`)
- `.env` populated with Drive folder IDs and Slack credentials

---

## Step 1: Upload to Google Drive

```bash
python tools/drive_upload.py
```

**Routing:**
| Local Path | Drive Folder |
|------------|-------------|
| `.tmp/reports/*.html` | `GOOGLE_DRIVE_OUTPUT_FINALIZED_ID` |
| `.tmp/raw/*.xlsx, *.csv` | `GOOGLE_DRIVE_OUTPUT_RAW_ID` |
| `.tmp/slides/*.pptx` | `GOOGLE_DRIVE_OUTPUT_SLIDES_ID` |

Files with the same name are updated (not duplicated) on re-upload.

---

## Step 2: Send Slack Notification

```bash
python tools/send_slack.py
```

**What it posts:**
- Header: "Weekly Inventory Update — [Date] (WK[n])"
- Overstock pill: valuation, % of total, units (from `run_params.json`)
- Aging pill: valuation, % of total, units (auto-extracted from filtered CSV)
- Executive summary bullets (from `run_params.json`)
- Buttons linking to Drive folders

---

## Setting Up Slack

### Create a Slack App
1. Go to https://api.slack.com/apps → **Create New App → From scratch**
2. Name: `Inventory Automation`, Workspace: your workspace
3. Go to **OAuth & Permissions**
4. Add Bot Token Scopes:
   - `chat:write`
   - `chat:write.public` (to post without being in the channel)
5. Install to workspace
6. Copy **Bot User OAuth Token** (`xoxb-...`)

### Add to .env
```
SLACK_BOT_TOKEN=xoxb-your-token-here
SLACK_CHANNEL_ID=C01234567
```

Get `SLACK_CHANNEL_ID` by right-clicking the channel in Slack → **Copy Link** → the ID is the last part of the URL.

---

## Notes
- Slack messages use Block Kit (rich formatting with buttons and fields)
- If you want to post to a private channel, invite the bot: `/invite @Inventory Automation` in the channel
- The Slack message does NOT attach files — it links to Drive folders. This keeps the message clean and files versioned in Drive.
