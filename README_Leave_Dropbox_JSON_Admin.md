# Dropbox-safe Leave Front End (JSON requests + admin-only compile)

This approach avoids Dropbox Excel sync conflicts by **not writing to the shared Excel workbook during routine edits**.

Instead:
- Each leave request is saved as a **separate JSON file** in a Dropbox-synced folder.
- The rota administrator runs **Compile** to write all requests into the workbook's `Leave` sheet.

## Files
- `leave_requests_dropbox_app_admin.py` — Streamlit UI (admin-gated compile)

## Install
```bash
pip install streamlit openpyxl pandas
```

## Run
```bash
streamlit run leave_requests_dropbox_app_admin.py
```

## Configure paths + admin password
Create `.streamlit/secrets.toml`:
```toml
LEAVE_REQUESTS_DIR = "/Users/<you>/Dropbox/Rota/LeaveRequests"
ROTA_WORKBOOK_PATH = "/Users/<you>/Dropbox/Rota/Rota_Publish_Template_ORtools.xlsx"

# Only rota admin should know this:
ROTA_ADMIN_PASSWORD = "change-me"
```

## Behaviour
- Users can add/edit/delete leave requests (JSON files) without the admin password.
- The **Compile** button is disabled until the admin password is entered (per browser session).
- Optional workbook backup is created on each compile.

## Notes on security
This is a practical control suitable for small teams. It is not enterprise SSO.
If you need per-user authentication, deploy Streamlit behind your organisation’s identity provider (e.g., reverse proxy with SSO).
