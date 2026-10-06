# DBFSSR2026

Project files and setup instructions will be added here.

## Faculty Insights dashboard

The working prototype now includes seven views using the supplied faculty data, consistent filters, tables and charts, CSV/Excel exports, source validation, and manual file refresh.

Start with [the Streamlit deployment guide](DEPLOYMENT.md). The application can import revised files; live Google Drive synchronisation and role-based shared hosting are not configured.

Original source files and generated reports are excluded from GitHub. A new hosted installation starts with no faculty records and accepts file uploads after access is configured. The local project package contains detailed data-model and maintenance documentation.

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run app.py
```
