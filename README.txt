================================================================
 Laser SFG Production Planning Dashboard - Standalone Local App
================================================================

WHAT THIS IS
------------
A self-contained dashboard that runs on YOUR computer and reads
LIVE data from your Google Sheet via the Apps Script JSON endpoint.

Files in this folder:
  app.py             - Python server (Flask) that fetches live JSON
  templates/         - HTML dashboard (loads /data from app.py)
  requirements.txt   - Python libraries (Flask, requests)
  run.bat            - Double-click on Windows to start everything
  Code.gs            - Apps Script that produces the JSON feed (already deployed)


ONE-TIME SETUP
--------------
1. Install Python from https://www.python.org/downloads/
   IMPORTANT: tick "Add Python to PATH" during installation.

2. (Optional, only if you change the Apps Script URL):
   Open app.py in Notepad and update the APP_SCRIPT_URL value
   near the top.


HOW TO USE IT EVERY DAY
-----------------------
1. Double-click  run.bat
   (the first run installs Flask + requests; subsequent runs are instant)

2. Wait a few seconds. A browser tab opens at http://localhost:5000
   showing the dashboard with LIVE data from your sheet.

3. Click the blue "Refresh" button (top-right of the page) any time
   you've edited the Google Sheet and want the latest numbers.

4. When you're done, close the browser tab. Then click the black
   server window and press Ctrl+C to stop the server.


WHAT THE DASHBOARD SHOWS
------------------------
* KPI cards: Distinct BOMs, KM total, KG total, top BOMs
* Top BOMs by required quantity (bar chart)
* Required quantity by SFG Level 3 category
* Pivot table:
    - SFG_BOM, SFG_CODE, SFG_NAME, Unit, Level 3, Level 4, Lines, MC Count
    - SFG REQUIRED (sum of SFG_REQUIRED_AG_BAL TO PRD FG)
    - Production Done (you type your own values - saved in browser)
    - Balance (= Required - Done, in red if negative)
* Cascading filters: pick Level 3 -> Level 4 narrows automatically
* Click any row to drill down by Party -> MC -> Production Order
* Sortable columns (click headers)


REFRESHING THE DATA
-------------------
Every click on the "Refresh" button re-pulls from Apps Script.
Apps Script ALWAYS reads the live spreadsheet on every request -
no caching, no staleness.

If you want to skip Refresh and always start fresh, just close
and reopen the browser tab.


TROUBLESHOOTING
---------------
* "pip install failed" -> Python isn't on PATH. Re-install Python
  with the "Add Python to PATH" checkbox.

* "Apps Script fetch failed" -> the URL in app.py might have changed
  after a redeploy. Open Apps Script, Manage deployments, copy the
  current Web app URL, paste into app.py.

* "Authorization required" page in Apps Script -> redeploy with
  access level "Anyone" (not "Anyone with Google account").

* Server won't start, port already in use -> change the port in
  app.py (last line) e.g.  app.run(port=5500).


SCHEDULED AUTO-REFRESH (OPTIONAL)
---------------------------------
If you want the JSON feed to be pre-warmed every 5 minutes (so the
first page load is instant), in Apps Script run the function
scheduleSync() once. It creates a time-based trigger that calls
syncToSheet every 5 minutes.
