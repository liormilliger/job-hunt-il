# -*- coding: utf-8 -*-
"""docx -> PDF via Google Drive. No Word, no LibreOffice, no COM/AppleScript
automation — this is what replaced docx2pdf so the pipeline runs the same on
Windows, macOS, and Linux (including a headless Docker container).

How it works: upload the .docx with mimeType='application/vnd.google-apps.document',
which makes Drive auto-convert it into a Google Doc on upload; export that Doc
as application/pdf; delete the temporary Google Doc. The conversion happens on
Google's servers, so this needs network access but no local office suite.

Known trade-off: Google Drive's docx->Doc conversion is not pixel-identical to
Microsoft Word's own rendering. Plain CVs/cover letters (the English
render_docx.js output) come through fine. Hebrew/RTL output specifically was
tuned against real Word — if you notice any RTL glitches after switching to
this path, see docs/GOOGLE_SETUP.md's note on verifying one Hebrew CV.

Public API matches docx2pdf's convert() signature so callers don't change:
    convert(docx_path, pdf_path)
"""
import io
import os

from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

import google_auth

_DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
_GDOC_MIME = "application/vnd.google-apps.document"
_PDF_MIME = "application/pdf"


def convert(docx_path: str, pdf_path: str, cfg: dict = None):
    if cfg is None:
        import jh_config
        cfg = jh_config.load()

    service = google_auth.get_drive_service(cfg)

    file_metadata = {
        "name": os.path.basename(docx_path),
        "mimeType": _GDOC_MIME,
    }
    folder_id = cfg.get("google_drive_folder_id", "")
    if folder_id:
        file_metadata["parents"] = [folder_id]

    media = MediaFileUpload(docx_path, mimetype=_DOCX_MIME, resumable=False)
    uploaded = service.files().create(
        body=file_metadata, media_body=media, fields="id"
    ).execute()
    doc_id = uploaded["id"]

    try:
        request = service.files().export_media(fileId=doc_id, mimeType=_PDF_MIME)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        with open(pdf_path, "wb") as f:
            f.write(buf.getvalue())
    finally:
        # Best-effort cleanup: don't let a delete failure hide a successful
        # PDF export, but don't leave junk Docs behind either.
        try:
            service.files().delete(fileId=doc_id).execute()
        except Exception as e:
            print(f"    NOTE: could not delete temporary Google Doc {doc_id}: {e}")
