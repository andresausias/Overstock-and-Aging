"""
drive_download.py
-----------------
Downloads files from three separate Google Drive input folders into
dedicated local subdirectories.

Folder mapping:
  GOOGLE_DRIVE_INPUT_AGING_ID      → .tmp/raw/aging/
  GOOGLE_DRIVE_INPUT_EXCLUSIONS_ID → .tmp/raw/exclusions/
  GOOGLE_DRIVE_INPUT_SHIPMENTS_ID  → .tmp/raw/shipments/

Usage:
    python tools/drive_download.py

Required .env keys:
    GOOGLE_DRIVE_INPUT_AGING_ID
    GOOGLE_DRIVE_INPUT_EXCLUSIONS_ID
    GOOGLE_DRIVE_INPUT_SHIPMENTS_ID
"""

import os
import io
import sys
from pathlib import Path
from dotenv import load_dotenv
from googleapiclient.http import MediaIoBaseDownload

sys.path.insert(0, str(Path(__file__).parent))
from auth_google import get_drive_service

load_dotenv()

ROOT = Path(__file__).parent.parent

SUPPORTED_MIME_TYPES = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",  # .xlsx
    "application/vnd.ms-excel",                                            # .xls
    "text/csv",                                                             # .csv
}

# Map: env key → local input subdirectory
FOLDER_MAP = {
    "GOOGLE_DRIVE_INPUT_AGING_ID":      ROOT / "input" / "aging",
    "GOOGLE_DRIVE_INPUT_EXCLUSIONS_ID": ROOT / "input" / "reference",
    "GOOGLE_DRIVE_INPUT_SHIPMENTS_ID":  ROOT / "input" / "shipments",
}


def list_files(service, folder_id: str) -> list[dict]:
    """Return all Excel/CSV files in the given Drive folder."""
    query = (
        f"'{folder_id}' in parents and trashed = false and ("
        + " or ".join(f"mimeType = '{m}'" for m in SUPPORTED_MIME_TYPES)
        + ")"
    )
    results = []
    page_token = None
    while True:
        response = service.files().list(
            q=query,
            spaces="drive",
            fields="nextPageToken, files(id, name, mimeType, modifiedTime)",
            pageToken=page_token,
        ).execute()
        results.extend(response.get("files", []))
        page_token = response.get("nextPageToken")
        if not page_token:
            break
    return results


def download_file(service, file_id: str, file_name: str, dest_dir: Path) -> Path:
    """Download a single file from Drive and return its local path."""
    dest_path = dest_dir / file_name
    request = service.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    dest_path.write_bytes(buffer.getvalue())
    return dest_path


def main():
    # Validate all three keys are set
    missing = [key for key in FOLDER_MAP if not os.getenv(key)]
    if missing:
        print(f"ERROR: Missing .env keys: {', '.join(missing)}")
        sys.exit(1)

    print("Connecting to Google Drive...")
    service = get_drive_service()

    total = 0
    for env_key, dest_dir in FOLDER_MAP.items():
        folder_id = os.getenv(env_key)
        dest_dir.mkdir(parents=True, exist_ok=True)
        label = env_key.replace("GOOGLE_DRIVE_INPUT_", "").replace("_ID", "").title()

        print(f"\n── {label} folder → {dest_dir.relative_to(ROOT)} ──")
        files = list_files(service, folder_id)

        if not files:
            print(f"  (no files found)")
            continue

        for f in files:
            print(f"  → {f['name']}")
            local_path = download_file(service, f["id"], f["name"], dest_dir)
            print(f"     Saved: {local_path.relative_to(ROOT)}")
            total += 1

    print(f"\nDone. {total} file(s) downloaded across all input folders.")


if __name__ == "__main__":
    main()
