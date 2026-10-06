"""Drive protocol checks use fake HTTP responses; no Google key or private data is embedded."""
from copy import deepcopy
from hashlib import md5
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

import drive_source
from data_pipeline import REQUIRED, read_sources
from drive_source import DriveSourceError, read_drive_sources

ROOT = Path(__file__).resolve().parents[1]


class Response:
    def __init__(self, payload=None, content=b"", status=200):
        self.payload = payload
        self.content = content
        self.status_code = status

    def json(self):
        return self.payload

    def iter_content(self, chunk_size):
        yield self.content

    def close(self):
        pass


class FakeDrive:
    def __init__(self):
        self.content = b"test source content"
        self.files = [{"id": f"file_{i}", "name": name, "mimeType": "application/octet-stream", "size": str(len(self.content)), "md5Checksum": md5(self.content).hexdigest(), "version": "1", "modifiedTime": "2026-10-06T00:00:00Z"} for i, name in enumerate(REQUIRED)]
        self.calls = []
        self.list_count = 0
        self.duplicate = False
        self.modified_after_download = False
        self.status = 200
        self.corrupt = False

    def get(self, url, **kwargs):
        self.calls.append((url, kwargs))
        assert url.startswith("https://www.googleapis.com/drive/v3/files")
        assert kwargs["timeout"] == (10, 90)
        if self.status != 200:
            return Response(status=self.status)
        if url == drive_source.API:
            self.list_count += 1
            files = deepcopy(self.files)
            if self.duplicate:
                files.append(dict(files[0], id="second_copy"))
            if self.modified_after_download and self.list_count > 1:
                files[0]["version"] = "2"
            if "pageToken" not in kwargs["params"]:
                return Response({"files": files[:5], "nextPageToken": "second_page"})
            return Response({"files": files[5:]})
        assert kwargs["params"]["alt"] == "media"
        assert kwargs["stream"] is True
        return Response(content=b"corrupted content!!" if self.corrupt else self.content)


def test_paginated_read_all_sources_with_provenance():
    session = FakeDrive()
    sources, manifest = read_drive_sources({"folder_id": "valid_folder"}, session)
    assert list(sources) == REQUIRED
    assert all(content == session.content for content in sources.values())
    assert len(manifest) == 11
    assert manifest[0]["Drive version"] == "1"
    assert session.list_count == 4  # two pages before and after the downloads
    assert len(session.calls) == 15


@pytest.mark.parametrize("problem,match", [("duplicate", "Duplicate"), ("missing", "Missing"), ("native", "original Excel"), ("corrupt", "Incomplete|checksum"), ("changed", "changed during"), ("denied", "denied access")])
def test_reject_unsafe_or_incomplete_reads(problem, match):
    session = FakeDrive()
    if problem == "duplicate":
        session.duplicate = True
    elif problem == "missing":
        session.files.pop()
    elif problem == "native":
        session.files[0]["mimeType"] = "application/vnd.google-apps.spreadsheet"
    elif problem == "corrupt":
        session.corrupt = True
    elif problem == "changed":
        session.modified_after_download = True
    else:
        session.status = 403
    with pytest.raises(DriveSourceError, match=match):
        read_drive_sources({"folder_id": "valid_folder"}, session)


def test_key_error_and_folder_error_do_not_expose_secrets():
    with pytest.raises(DriveSourceError, match="key is missing or invalid") as error:
        read_drive_sources({"folder_id": "valid_folder", "service_account_json": "FAKE_PRIVATE_SECRET"})
    assert "FAKE_PRIVATE_SECRET" not in str(error.value)
    session = FakeDrive()
    with pytest.raises(DriveSourceError, match="folder ID"):
        read_drive_sources({"folder_id": "https://drive.google.com/drive/folders/x"}, session)
    assert not session.calls


def test_only_readonly_scope_and_fixed_google_auth_destination(monkeypatch):
    import json
    observed = {}

    def fake_credentials(info, scopes):
        observed.update(info=info, scopes=scopes)
        return "credentials"

    monkeypatch.setattr(drive_source.service_account.Credentials, "from_service_account_info", fake_credentials)
    monkeypatch.setattr(drive_source, "AuthorizedSession", lambda credentials: credentials)
    assert drive_source.authorized_session({"service_account_json": json.dumps({"type": "service_account", "token_uri": "https://untrusted.invalid", "universe_domain": "untrusted.invalid"})}) == "credentials"
    assert observed["scopes"] == ["https://www.googleapis.com/auth/drive.readonly"]
    assert observed["info"]["token_uri"] == "https://oauth2.googleapis.com/token"
    assert observed["info"]["universe_domain"] == "googleapis.com"


def test_drive_app_new_session_filter_and_failed_refresh(monkeypatch, tmp_path):
    # Optional actual-data integration check in the prepared private workspace.
    if not all((ROOT / "data" / name).is_file() for name in REQUIRED):
        pytest.skip("Private source files are not part of the code repository")
    sources = read_sources(ROOT / "data")
    fail = [False]
    reads = []

    def fake_read(config):
        reads.append(config["folder_id"])
        if fail[0]:
            raise DriveSourceError("Google Drive denied access.")
        return sources, [{"File": name, "Drive file ID": f"test_{i}", "Drive version": "1"} for i, name in enumerate(REQUIRED)]

    monkeypatch.setattr(drive_source, "read_drive_sources", fake_read)
    monkeypatch.setenv("DASHBOARD_DATA_DIR", str(tmp_path / "no-local-sources"))

    def new_session():
        app = AppTest.from_file(str(ROOT / "app.py"), default_timeout=60)
        app.secrets["google_drive"] = {"folder_id": "test_folder", "service_account_json": "fake"}
        return app.run()

    at = new_session()
    assert not at.exception
    assert at.metric[0].value == "2,295"
    assert not at.get("file_uploader")
    assert any("Source: Google Drive" in item.value for item in at.caption)
    at.multiselect(key="filter:Students:Campus").set_value(["Puncak Alam"]).run()
    assert len(reads) == 1  # changing a filter must not download again
    fail[0] = True
    next(b for b in at.button if b.label == "Refresh source files").click().run()
    assert not at.exception
    assert at.metric[0].value == "1,049"
    assert at.multiselect(key="filter:Students:Campus").value == ["Puncak Alam"]
    assert any("Refresh failed" in item.value for item in at.error)
    failed_session = new_session()
    assert not failed_session.metric
    assert failed_session.error
    previous_reads = len(reads)
    failed_session.run()
    assert len(reads) == previous_reads  # failure should not loop on every UI change
    fail[0] = False
    at2 = new_session()
    assert not at2.exception
    assert not at2.error
    assert at2.metric[0].value == "2,295"
    assert "replacements" not in at2.session_state
