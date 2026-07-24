# -*- coding: utf-8 -*-
"""job-hunt-il shared Google auth.

Every script that touches the tracker (Google Sheets) or converts a document
to PDF (Google Drive) goes through here for credentials. This uses an OAuth
"installed app" flow against the USER'S OWN Google account (never a service
account) so every Sheet/Doc the pipeline creates lands in the user's own
Drive, exactly as if they'd made it by hand — no separate service-account
Drive that has to be explicitly shared.

One-time setup (~5 minutes): docs/GOOGLE_SETUP.md walks through creating an
OAuth client ID in Google Cloud Console and downloading credentials.json into
the workdir. First run after that opens a browser for consent; a cached
token (token.json, also in the workdir) is reused silently after that until
it's revoked or the scopes change.

Nothing here ever touches ANTHROPIC_API_KEY — that stays in jh_config.py.
"""
import os

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# spreadsheets: read/write the tracker.
# drive.file: create/read/export only files THIS app creates (or the user
# explicitly opens with it) — deliberately narrower than full Drive access.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]

_creds_cache = None
_sheets_service = None
_drive_service = None


def _resolve(cfg, key, default):
    val = cfg.get(key) or default
    if not os.path.isabs(val):
        val = os.path.join(cfg["workdir"], val)
    return val


def get_credentials(cfg: dict):
    """Load cached credentials, refresh them, or run the consent flow.

    Caches in-process (module-level) so repeated calls within one script run
    don't re-touch disk or re-prompt.
    """
    global _creds_cache
    if _creds_cache and _creds_cache.valid:
        return _creds_cache

    token_path = _resolve(cfg, "google_token", "token.json")
    creds_path = _resolve(cfg, "google_credentials", "credentials.json")

    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(creds_path):
                raise FileNotFoundError(
                    f"Google OAuth client file not found at {creds_path}.\n"
                    "See docs/GOOGLE_SETUP.md to create one (one-time, ~5 min): "
                    "Google Cloud Console -> OAuth client ID (Desktop app) -> "
                    "download the JSON -> save it at that path (or point "
                    "config.json 'google_credentials' at it)."
                )
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            print("Opening a browser to sign in with your Google account "
                  "(one-time; the pipeline only ever touches files it creates)...")
            creds = flow.run_local_server(port=0)
        with open(token_path, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    _creds_cache = creds
    return creds


def get_sheets_service(cfg: dict):
    global _sheets_service
    if _sheets_service is None:
        _sheets_service = build("sheets", "v4", credentials=get_credentials(cfg),
                                 cache_discovery=False)
    return _sheets_service


def get_drive_service(cfg: dict):
    global _drive_service
    if _drive_service is None:
        _drive_service = build("drive", "v3", credentials=get_credentials(cfg),
                                cache_discovery=False)
    return _drive_service
