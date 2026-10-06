"""Private, read-only Google Drive sources. No faculty files or keys are written to disk."""
from __future__ import annotations

from hashlib import md5
import json
import re

from google.auth.transport.requests import AuthorizedSession
from google.oauth2 import service_account

from data_pipeline import REQUIRED

API = "https://www.googleapis.com/drive/v3/files"
SCOPE = "https://www.googleapis.com/auth/drive.readonly"
MAX_FILE_BYTES = 200 * 1024 * 1024
GOOGLE_DOCUMENT = "application/vnd.google-apps."


class DriveSourceError(ValueError):
    """An actionable message safe to display without revealing credentials."""


def validate_config(config):
    folder_id = str(config.get("folder_id", "")).strip()
    if not re.fullmatch(r"[A-Za-z0-9_-]+", folder_id):
        raise DriveSourceError("Set google_drive.folder_id to the folder ID, not the full Drive link, in Streamlit Secrets.")
    return folder_id


def authorized_session(config):
    try:
        info = json.loads(config.get("service_account_json", ""))
        if info.get("type") != "service_account":
            raise ValueError("Wrong credential type")
        # Use Google's fixed token destination, never one supplied in an uploaded key.
        info["token_uri"] = "https://oauth2.googleapis.com/token"
        info["universe_domain"] = "googleapis.com"
        credentials = service_account.Credentials.from_service_account_info(info, scopes=[SCOPE])
        return AuthorizedSession(credentials)
    except Exception:
        raise DriveSourceError("The Google Drive service-account key is missing or invalid. Check google_drive.service_account_json in Streamlit Secrets.") from None


def _request(session, url, **kwargs):
    try:
        response = session.get(url, timeout=(10, 90), **kwargs)
    except Exception:
        raise DriveSourceError("Could not connect to Google Drive. Check the app key, network access and Drive API, then retry.") from None
    if response.status_code != 200:
        response.close()
        messages = {
            401: "Google Drive authentication failed. Check that the app's service-account key is still valid.",
            403: "Google Drive denied access or a quota was reached. Enable the Drive API and share the source folder with the service-account email as Viewer. Retry after a quota limit clears.",
            404: "Google Drive folder or file was not found or is not shared with the service-account email. Check the folder ID and Viewer access.",
            429: "Google Drive is rate limiting requests. Wait a few minutes and refresh again.",
        }
        raise DriveSourceError(messages.get(response.status_code, "Google Drive could not complete the read. Retry later."))
    return response


def _list_files(session, folder_id):
    selected = {}
    page_token = None
    seen_pages = set()
    while True:
        params = {
            "q": f"'{folder_id}' in parents and trashed = false",
            "fields": "nextPageToken,files(id,name,mimeType,size,modifiedTime,md5Checksum,version)",
            "pageSize": 1000,
            "supportsAllDrives": "true",
            "includeItemsFromAllDrives": "true",
        }
        if page_token:
            params["pageToken"] = page_token
        response = _request(session, API, params=params)
        try:
            payload = response.json()
        except Exception:
            raise DriveSourceError("Google Drive returned an invalid file listing. Retry the refresh.") from None
        finally:
            response.close()
        for item in payload.get("files", []):
            name = item.get("name")
            if name not in REQUIRED:
                continue
            if name in selected:
                raise DriveSourceError(f"Duplicate source filename in Drive: {name}. Keep one current file with this name in the connected folder.")
            if not re.fullmatch(r"[A-Za-z0-9_-]+", str(item.get("id", ""))):
                raise DriveSourceError("Google Drive returned an invalid file ID.")
            if item.get("mimeType", "").startswith(GOOGLE_DOCUMENT):
                raise DriveSourceError(f"Keep the original Excel/PDF/PowerPoint file in Drive, rather than a Google document or shortcut: {name}")
            if int(item.get("size", 0)) > MAX_FILE_BYTES:
                raise DriveSourceError(f"Source file exceeds the 200 MB limit: {name}")
            selected[name] = item
        page_token = payload.get("nextPageToken")
        if not page_token:
            break
        if page_token in seen_pages:
            raise DriveSourceError("Google Drive returned a repeated page. Retry the refresh.")
        seen_pages.add(page_token)
    missing = set(REQUIRED) - set(selected)
    if missing:
        raise DriveSourceError("Missing files in the connected Drive folder (check original filenames and Viewer access): " + "; ".join(sorted(missing)))
    return selected


def _revision(item):
    return tuple(item.get(field) for field in ("id", "version", "modifiedTime", "md5Checksum", "size"))


def read_drive_sources(config, session=None):
    """All eleven required files from one folder; reject ambiguous or changing sources."""
    folder_id = validate_config(config)
    owned_session = session is None
    session = session or authorized_session(config)
    try:
        before = _list_files(session, folder_id)
        sources = {}
        for name in REQUIRED:
            item = before[name]
            response = _request(session, f"{API}/{item['id']}", params={"alt": "media", "supportsAllDrives": "true"}, stream=True)
            try:
                chunks = []
                total = 0
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    total += len(chunk)
                    if total > MAX_FILE_BYTES:
                        raise DriveSourceError(f"Source file exceeds the 200 MB limit: {name}")
                    chunks.append(chunk)
                content = b"".join(chunks)
            except DriveSourceError:
                raise
            except Exception:
                raise DriveSourceError(f"Download interrupted for {name}. Retry the refresh.") from None
            finally:
                response.close()
            if "size" in item and len(content) != int(item["size"]):
                raise DriveSourceError(f"Incomplete or changed download: {name}. Retry the refresh.")
            # Drive uses MD5 for binary-file download integrity, not authentication.
            if item.get("md5Checksum") and md5(content, usedforsecurity=False).hexdigest() != item["md5Checksum"]:
                raise DriveSourceError(f"Download checksum mismatch: {name}. Retry the refresh.")
            sources[name] = content
        after = _list_files(session, folder_id)
        if any(_revision(before[name]) != _revision(after[name]) for name in REQUIRED):
            raise DriveSourceError("A source file changed during the refresh. Wait for editing to finish and retry; the previous data is preserved.")
        manifest = [{"File": name, "Drive file ID": before[name]["id"], "Drive modified time": before[name].get("modifiedTime"), "Drive version": before[name].get("version")} for name in REQUIRED]
        return sources, manifest
    finally:
        if owned_session:
            session.close()
