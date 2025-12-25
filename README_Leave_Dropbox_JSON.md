# Dropbox-safe Leave Front End (JSON requests + compile to Excel)

This approach avoids Dropbox Excel sync conflicts by **not writing to the shared Excel workbook during routine edits**.

Instead:
- Each leave request is saved as a **separate JSON file** in a Dropbox-synced folder.
- A rota administrator runs **Compile** to write all requests into the workbook's `Leave` sheet.

## Files
- `leave_requests_dropbox_app.py` — Streamlit UI

## Install
```bash
pip install streamlit openpyxl pandas
```

## Run
```bash
streamlit run leave_requests_dropbox_app.py
```

## Configure paths (recommended)
Create `.streamlit/secrets.toml`:
```toml
LEAVE_REQUESTS_DIR = "/Users/<you>/Dropbox/Rota/LeaveRequests"
ROTA_WORKBOOK_PATH = "/Users/<you>/Dropbox/Rota/Rota_Publish_Template_ORtools.xlsx"
```

## Operating model
- Everyone can add/edit/delete leave requests (JSON files sync cleanly in Dropbox).
- Only the rota administrator should run **Compile** to update the workbook.
- Enable workbook backups on compile.

## JSON schema (per request)
Each file is `<RequestID>.json` and includes:
- request_id, name, start_date, end_date, leave_type, approved, notes, created_at, updated_at
