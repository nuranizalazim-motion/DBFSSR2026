# Deploy FSSR Faculty Insights

This repository contains application code. Faculty source files, generated reports, detailed source profiles and the portable HTML data snapshot are not included.

## Streamlit Community Cloud settings

On the deployment form, use:

| Field | Value |
|---|---|
| Repository | `nuranizalazim-motion/DBFSSR2026` |
| Branch | `main` |
| Main file path | `app.py` |
| Python version in Advanced settings | `3.12` |

An app URL can use the automatically suggested name. `requirements.txt` installs the required packages with versions constrained by `requirements.lock`.

## Access and data

The initial deployment contains no faculty records. After deployment, configure the application's sharing/access settings for private access and your intended users **before uploading faculty records**. The application does not implement separate management-versus-lecturer roles.

1. Open **Replace source files** in the application's sidebar.
2. Upload all eleven original files, retaining their original filenames.
3. Press **Validate and apply uploads**.
4. Check the successful import time and the source-validation report.

Uploads are session-only and do not write files back to GitHub. File refresh is manual, distinct from filter updates. A new session may require re-uploading the sources. Uploaded data is processed on the hosting service; use a faculty-approved hosting arrangement for shared institutional use.

For persistent shared sources, follow [Connect private Google Drive](GOOGLE_DRIVE_SETUP.md). The app supports a dedicated read-only service account configured in Streamlit Secrets. Once access is configured, new sessions load the eleven source files from Drive automatically, and **Refresh source files** reads updates on demand. No background refresh schedule is active. This connection is not active until the key, folder ID and Viewer access are configured. Do not add source records, credential keys or exported dashboard snapshots to the code repository.

## Running the code locally

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run app.py
```

Python 3.12 was used to prepare and test this application. Local `data/` files can provide the sources instead of browser uploads. Those files and generated `reports/` remain Git-ignored.
