# DBFSSR2026

Project files and setup instructions will be added here.

## Faculty Insights dashboard

The working prototype now includes seven views using the supplied faculty data, consistent filters, tables and charts, CSV/Excel exports, source validation, and manual file refresh.

Start with [the Streamlit deployment guide](DEPLOYMENT.md). For sources that survive browser and app restarts, follow [the private Google Drive connection guide](GOOGLE_DRIVE_SETUP.md). The app supports read-only Drive access, automatic loading in new sessions and refresh on demand. Separate management-versus-lecturer roles and scheduled background refresh are not configured.

Original source files, credential keys and generated reports are excluded from GitHub. A new hosted installation loads sources from Drive once private credentials and folder access are configured. Without that connection it accepts temporary browser uploads after access is configured. The local project package contains detailed data-model and maintenance documentation.

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run app.py
```
