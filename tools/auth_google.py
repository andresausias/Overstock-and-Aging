"""
auth_google.py
--------------
Shared Google OAuth helper. Returns authorized service clients for Drive and Slides APIs.

Usage:
    from auth_google import get_drive_service, get_slides_service

Credentials:
    - Place credentials.json (downloaded from GCP) in the project root.
    - On first run, a browser window opens for OAuth consent.
    - token.json is saved automatically for subsequent runs.
"""

import os
from pathlib import Path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Scopes required by this project
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
]

# Paths relative to project root (one level above tools/)
ROOT = Path(__file__).parent.parent
CREDENTIALS_PATH = ROOT / "credentials.json"
TOKEN_PATH = ROOT / "token.json"


def _get_credentials() -> Credentials:
    """Load or refresh Google OAuth credentials."""
    creds = None

    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_PATH.exists():
                raise FileNotFoundError(
                    f"credentials.json not found at {CREDENTIALS_PATH}.\n"
                    "Download it from GCP Console → APIs & Services → Credentials.\n"
                    "See workflows/setup_google_auth.md for setup instructions."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_PATH), SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open(TOKEN_PATH, "w") as token_file:
            token_file.write(creds.to_json())

    return creds


def get_drive_service():
    """Return an authorized Google Drive v3 service client."""
    creds = _get_credentials()
    return build("drive", "v3", credentials=creds)


def get_slides_service():
    """Return an authorized Google Slides v1 service client."""
    creds = _get_credentials()
    return build("slides", "v1", credentials=creds)
