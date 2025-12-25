# Rota Leave Admin (Shared Master Workbook)

This Streamlit UI writes **directly** to a shared master workbook path and supports:
- Add leave rows
- Edit existing leave rows
- Delete leave rows (clears the row rather than shifting sheet rows)

## Install
```bash
pip install streamlit openpyxl pandas
```

## Run
```bash
streamlit run leave_frontend_streamlit_master.py
```

## Configure the shared master workbook path
You can either:
1) Paste the path into the app (top field), or
2) Use Streamlit secrets to prefill it.

Create `.streamlit/secrets.toml`:
```toml
MASTER_WORKBOOK_PATH = "/Users/hakeemyusuff/Dropbox/UHL/Personal/Rota_solution/Rota_Master.xlsx"
```

## Concurrency notes
Excel files are not transactional. To reduce collisions this app can create a simple lock file:
`Rota_Publish_Template_ORtools.xlsx.lock` during write operations.

For high-concurrency teams, consider migrating leave storage to a small database (SQLite/Postgres) and exporting to Excel.
