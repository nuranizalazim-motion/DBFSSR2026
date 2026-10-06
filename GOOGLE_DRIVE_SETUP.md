# Connect your private Google Drive folder

This connection keeps the original files in Google Drive. The app reads them when a new browser session opens, so closing the browser or restarting the app does not require uploading the sources again. Browser uploads alone do not provide this persistence.

The code is prepared for a Google service account: a dedicated app identity with Viewer access to your source folder. Your Streamlit app should remain private. Viewer invitations to the dashboard are separate from sharing the Drive folder with the app identity.

## 1. Create a Google Cloud project

1. Open [Google Cloud: create a project](https://console.cloud.google.com/projectcreate).
2. Sign in with the account permitted to manage this connection.
3. Name the project **FSSR Dashboard** and click **Create**.
4. Select the new project using the project selector at the top of Google Cloud.
5. Open [Google Drive API](https://console.cloud.google.com/apis/library/drive.googleapis.com) and click **Enable** for that project.

The standard Drive API does not charge for normal API requests, but usage quotas apply. This setup does not require a paid Google Cloud compute service. Organization policies may restrict project creation, service-account keys or external folder sharing. If a step is blocked by UiTM policy, ask your IT administrator to provide an approved project and app identity; keep the folder private.

## 2. Create the app identity and its key

1. In Google Cloud, open **IAM & Admin → Service Accounts**.
2. Click **Create service account**.
3. Use **fssr-dashboard-reader** as its name and click **Create and continue**.
4. Leave optional project permissions and user access blank, then click **Done**. No Owner or Editor project role is needed to read this folder.
5. Open the new service account. Copy its email address, which ends in **iam.gserviceaccount.com**.
6. Open its **Keys** tab, select **Add key → Create new key → JSON**, and click **Create**. Google downloads a JSON key file to your computer.

Keep this key private. Do not upload it to the source folder, GitHub, the dashboard's file uploader or this chat. It belongs only in the app's private Secrets settings. If Google blocks key creation, ask your IT administrator for the approved hosting/authentication route.

## 3. Share the source folder with the app

1. Open the Google Drive folder containing the eleven original Excel, PDF and PowerPoint files.
2. Click **Share** on the folder.
3. Add the service-account email from step 2, choose **Viewer**, and complete sharing. A notification is not needed for this app identity.
4. Leave **General access** set to **Restricted**.
5. Put all eleven original files directly inside that folder. Keep their exact original filenames and original file formats; do not convert them into Google Sheets, Docs or Slides. Shortcuts and nested subfolders are not followed.
6. Keep one current copy of each required filename in this folder. Google Drive allows duplicate names; the dashboard rejects duplicates to avoid reading the wrong file.

The folder ID is the part of its URL immediately after `/folders/`, before any `?` or further `/`. For example, `https://drive.google.com/drive/folders/ABC123?usp=sharing` has folder ID `ABC123`.

## 4. Store the connection privately in Streamlit

1. Open your deployed dashboard.
2. Click **Manage app**, then open **Settings → Secrets**. You can also open the app's settings from its menu on [Streamlit Community Cloud](https://share.streamlit.io/).
3. Add the following TOML block. If other secrets already exist, keep them and append this block only once.

```toml
[google_drive]
folder_id = "YOUR_FOLDER_ID"
service_account_json = '''
PASTE THE COMPLETE CONTENTS OF YOUR DOWNLOADED JSON KEY HERE
'''
```

4. Replace `YOUR_FOLDER_ID` with the folder ID, retaining the double quotes.
5. Open the downloaded JSON file in a plain text editor such as TextEdit in plain text mode or a code editor. Copy all its text from the opening `{` to the closing `}`.
6. Replace the placeholder sentence with that complete JSON text. Retain the three single quotes on the lines above and below it. Do not edit the key or change its `\n` characters. The triple single quotes preserve the JSON exactly.
7. Click **Save** in Streamlit. Wait for the app to restart. If necessary, use **Reboot app** from Manage app.

The app uses the key to request Google's read-only Drive scope. It has no code to modify Drive files. App viewer login remains controlled by Streamlit's private sharing settings.

## 5. Confirm that it works

1. Open the dashboard again. The sidebar should say **Source: private Google Drive**.
2. Wait for the source read and validation to finish. The successful-read caption should include **Source: Google Drive** and **11 source files**.
3. With the original supplied files, Students should show **2,295**, with **14 programmes** and **4 campuses**. Updated source files may legitimately produce different totals.
4. Select only **Puncak Alam**. With the original files, the student total should be **1,049**.
5. Click **Refresh source files** and confirm the selection remains active.
6. Close and reopen the app in a new browser session. It should load from Drive without a file upload. Filters start fresh in a new session; source access persists.
7. Inspect the source validation report and manifest. The manifest includes SHA-256 hashes, Drive file IDs, modification times and versions.

## Refresh and maintenance

The permanent source is Drive, rather than browser-session memory or the code repository. Each new session downloads and validates all eleven files. In an existing session, **Refresh source files** reads changes on demand; filters update immediately against that session's last successful read. There is no background polling or guaranteed hourly refresh. Wait for a refresh to finish before starting another. A large file or slow connection can take longer than a minute.

To update sources, edit or replace the current original-format files in the connected folder, keeping the required filenames. Wait for the Drive uploads to finish, then refresh the dashboard. Do not add a second copy with the same filename. The app refuses missing, duplicated, truncated or changing files. The previous successful snapshot and filter selections remain available in an existing session if refresh fails. A new session with a failed read cannot recover another session's snapshot; it displays an error instead.

All figures, charts, detail tables and exports use the same selected records. Last successful read time is displayed in Malaysia time; it indicates source reading, not confirmation that the underlying business records are current. The stale-data threshold flags how long it has been since a successful read.

To share the dashboard, invite permitted colleagues through the app's **Share** controls while keeping it private. They do not each need Drive access: the server uses the dedicated app identity to read the folder. All app viewers currently have the same views and export capability; separate lecturer/management roles are not implemented.

The connection needs the Drive API enabled, a valid service-account key, Viewer access to the folder, and a running Streamlit host with HTTPS access to `www.googleapis.com` and `oauth2.googleapis.com`. Streamlit Community Cloud may sleep or restart; it will load from Drive again after waking. Hosting and institutional account policies can limit who can access the app.

For key rotation, create a replacement JSON key, update Streamlit Secrets, restart and verify the source read, then revoke the old key in Google Cloud. Removing the app's folder access or deleting its key disconnects future reads. The original source files remain unchanged.

## Common problems

| Message or symptom | Action |
|---|---|
| Missing or invalid service-account key | Check that the complete JSON is inside the triple single quotes in Streamlit Secrets. Do not paste the JSON directly as TOML. |
| Cannot read Streamlit Secrets | Check TOML punctuation and that `[google_drive]` appears only once. |
| Access denied / no files listed | Enable Drive API for the key's project, verify the folder ID, and share the folder with the service-account email as Viewer. |
| Missing source filename | Check every original filename, including spaces and punctuation; keep the eleven files directly inside the folder. |
| Duplicate source filename | Move older copies outside the connected folder, retaining the intended current file. |
| Google document or shortcut | Upload the original `.xlsx`, `.pdf` or `.pptx` file into the folder instead. |
| Source changed during refresh | Wait for edits/uploads to finish and press Refresh again. |
| Rate limit or temporary failure | Wait a few minutes, then retry. Existing validated data is retained in the current session. |
| Still asks for browser uploads | The app has not loaded the `[google_drive]` secrets section; check settings, save and restart. |

No private key, faculty file or data export should be added to the code repository. The repository contains the connection code and this guide only.
