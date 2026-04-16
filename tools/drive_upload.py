"""
drive_upload.py
---------------
Uploads all generated outputs to designated Google Drive folders.

Folder routing:
  .tmp/reports/*.html  →  GOOGLE_DRIVE_OUTPUT_FINALIZED_ID
  .tmp/raw/*           →  GOOGLE_DRIVE_OUTPUT_RAW_ID
  .tmp/slides/*.pptx   →  GOOGLE_DRIVE_OUTPUT_SLIDES_ID

Usage:
    python tools/drive_upload.py

Required .env keys:
    GOOGLE_DRIVE_OUTPUT_FINALIZED_ID
    GOOGLE_DRIVE_OUTPUT_RAW_ID
    GOOGLE_DRIVE_OUTPUT_SLIDES_ID
"""

import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from googleapiclient.http import MediaFileUpload

sys.path.insert(0, str(Path(__file__).parent))
from auth_google import get_drive_service

load_dotenv()

ROOT = Path(__file__).parent.parent

MIME_MAP = {
    ".html":  "text/html",
    ".csv":   "text/csv",
    ".xlsx":  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pptx":  "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".pdf":   "application/pdf",
    ".json":  "application/json",
}


def upload_file(service, local_path: Path, folder_id: str) -> str:
    """Upload a file to Drive, replacing existing if same name. Returns file ID."""
    mime_type = MIME_MAP.get(local_path.suffix.lower(), "application/octet-stream")
    file_metadata = {"name": local_path.name, "parents": [folder_id]}

    # Check if file already exists in folder (to update instead of duplicate)
    existing = service.files().list(
        q=f"name = '{local_path.name}' and '{folder_id}' in parents and trashed = false",
        fields="files(id, name)",
    ).execute().get("files", [])

    media = MediaFileUpload(str(local_path), mimetype=mime_type, resumable=True)

    if existing:
        # Update existing file
        file_id = existing[0]["id"]
        service.files().update(fileId=file_id, media_body=media).execute()
        return file_id
    else:
        # Create new file
        f = service.files().create(
            body=file_metadata, media_body=media, fields="id"
        ).execute()
        return f["id"]


def upload_directory(service, dir_path: Path, folder_id: str, extensions: list[str]) -> int:
    """Upload all matching files from a directory. Returns upload count."""
    if not dir_path.exists():
        return 0
    count = 0
    for f in sorted(dir_path.iterdir()):
        if f.is_file() and f.suffix.lower() in extensions:
            print(f"  Uploading: {f.name} → Drive folder {folder_id[:12]}...")
            file_id = upload_file(service, f, folder_id)
            print(f"    ✓ File ID: {file_id}")
            count += 1
    return count


def main():
    finalized_id = os.getenv("GOOGLE_DRIVE_OUTPUT_FINALIZED_ID")
    raw_id       = os.getenv("GOOGLE_DRIVE_OUTPUT_RAW_ID")
    slides_id    = os.getenv("GOOGLE_DRIVE_OUTPUT_SLIDES_ID")

    missing = [k for k, v in {
        "GOOGLE_DRIVE_OUTPUT_FINALIZED_ID": finalized_id,
        "GOOGLE_DRIVE_OUTPUT_RAW_ID": raw_id,
        "GOOGLE_DRIVE_OUTPUT_SLIDES_ID": slides_id,
    }.items() if not v]
    if missing:
        print(f"ERROR: Missing .env keys: {', '.join(missing)}")
        sys.exit(1)

    print("Connecting to Google Drive...")
    service = get_drive_service()

    total = 0

    print("\n── Uploading HTML reports → finalized/ ──")
    total += upload_directory(service, ROOT / ".tmp" / "reports", finalized_id, [".html"])

    print("\n── Uploading raw files → raw/ ──")
    for subdir in ["aging", "exclusions", "shipments"]:
        total += upload_directory(service, ROOT / ".tmp" / "raw" / subdir, raw_id, [".xlsx", ".xls", ".csv"])

    print("\n── Uploading slides → slides/ ──")
    total += upload_directory(service, ROOT / ".tmp" / "slides", slides_id, [".pptx", ".pdf"])

    print(f"\nDone. {total} file(s) uploaded to Drive.")


if __name__ == "__main__":
    main()
