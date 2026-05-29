# Laser SFG Production Planning Dashboard

A static, single-page dashboard for tracking SFG BOM requirements, equipment scheduling, HOD workload, and per-line production status against an ERP feed and per-line manual inputs synced through Google Apps Script.

## Live Site

Once GitHub Pages is enabled on this repo, the dashboard is served from `index.html`.

## Architecture

- **Frontend:** Single `index.html` file (HTML + CSS + vanilla JS + Chart.js + JSONP).
- **Backend:** Two Google Apps Script Web Apps:
  - **ERP feed** (`Code.gs`) — reads the live ERP tab from a Google Sheet and serves JSON for the pivot + drill-down data, plus EQPT_MASTER / HOD MASTER lookups.
  - **Planning inputs** (`planning_inputs_Code.gs`) — handles per-line manual inputs (Done tick, prod_done qty, EQPT, HOD, commit dates, reschedule count + reason) and user authentication / role management.
- **Optional Python wrapper:** `app.py` (Flask) runs locally at `http://localhost:5000` and proxies the Apps Script feed when running the dashboard outside GitHub Pages.

## Pages / Views

- **SFG Dashboard** — main pivot table with KPI cards, Equipment / HOD / Party balance side tables, drill-down per BOM
- **Team Master** — manpower / team management
- **MC Status** — MC-level production status from a separate Google Sheet tab
- **Modals:** Login, Drill-down child table, Reschedule per equipment, User Panel (Super Admin), Filter Editor

## Roles

- **user** — read-only on Reschedule modal, can set commit dates once (write-once lock thereafter), can mark Done / enter Production Done Qty
- **admin** — full edit on commit dates + reschedule with mandatory reason
- **super_admin** — everything; only role that can edit users, clear Production Done globally, override locks

## Key Features

- Live 30 s auto-refresh with 35 s recent-edit protection per line
- Per-EQPT auto-reschedule modal with reorder + duration edit, mandatory reason, forward-push counter
- Excel-style cascading filter dropdowns (Unit, L3, L4, EQPT, HOD, Party, Approve Status, Commit Start, Commit End)
- Per-table row highlight with combinable filters across tables
- Overdue indicators, commit conflict detection per EQPT, past-date blocking
- Production Done capped at SFG Net Qty per line
- Sheet auto-migration on first run for new columns

## Local Run (optional)

```bash
# Install Python deps (one-time)
pip install -r requirements.txt

# Start local Flask server
python app.py
# OR on Windows
run.bat
```

Then open http://localhost:5000

## Apps Script Setup

1. Open the bound Google Sheet → Extensions → Apps Script.
2. Paste `Code.gs` and `planning_inputs_Code.gs` into separate script projects (or combine).
3. Deploy each as **Web App**, access **Anyone**, copy the URL.
4. Update the URLs in `index.html`:
   - `APP_SCRIPT_URL` → ERP feed URL
   - `INPUT_API_URL` → planning inputs URL

## License

Internal tool — not for public distribution.
