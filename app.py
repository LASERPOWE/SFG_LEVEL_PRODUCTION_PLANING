"""
Laser SFG Production Planning Dashboard - local Python server.

Fetches LIVE data from your Apps Script JSON endpoint every time the
dashboard is opened or refreshed, and serves an interactive dashboard
at http://localhost:5000

USAGE
-----
1. (One-time) Install dependencies:
        pip install -r requirements.txt
2. Start the server:
        python app.py
   (or just double-click run.bat on Windows)
3. Open in browser:
        http://localhost:5000
"""

from flask import Flask, render_template, jsonify, request
import requests
import json
import time
from collections import defaultdict
from datetime import datetime

# --------------------------------------------------------------------
# Apps Script endpoint that returns the live ERP DUMP_PROD_ORD_OPEN MC
# tab as JSON.  Change here if you redeploy with a different URL.
# --------------------------------------------------------------------
APP_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzJn0_gDs_nvFeg5X7nlM99dUZ4tnhI3Z1ionkM9RGfU_TS4lIWYThEGjt6EpTnMtHbgQ/exec"

# Cache the live response in memory so we don't hammer Apps Script on
# every chart redraw.  Refresh button in UI bypasses this with ?refresh=1
CACHE_SECONDS = 60

app = Flask(__name__)
# Cache is keyed by tab name so we can fetch the main dump and EQPT_MASTER
# (and any other tab) independently without each call evicting the other.
_cache: dict = {}


def fetch_live_data(tab: str | None = None, force: bool = False) -> dict:
    """Pull the live JSON from Apps Script.  requests follows redirects
    through googleusercontent.com transparently, unlike a sandboxed
    browser fetch."""
    cache_key = tab or "__default__"
    now = time.time()
    entry = _cache.get(cache_key)
    if not force and entry and now - entry["ts"] < CACHE_SECONDS:
        return entry["data"]
    params = {"format": "json"}
    if tab:
        params["tab"] = tab
    r = requests.get(APP_SCRIPT_URL, params=params, allow_redirects=True, timeout=60)
    r.raise_for_status()
    data = r.json()
    _cache[cache_key] = {"data": data, "ts": now}
    return data


def _find_eqpt_key(rows):
    """Detect the EQPT identifier column in the EQPT_MASTER tab."""
    if not rows:
        return None
    keys = list(rows[0].keys())
    candidates = ["EQPT_ID_NAME", "EQPT_NAME", "EQPT_ID", "EQPTID", "EQUIPMENT"]
    for c in candidates:
        if c in keys:
            return c
    for k in keys:
        if "EQPT" in k.upper() or "EQUIP" in k.upper():
            return k
    return None


def fetch_eqpt_master(force: bool = False) -> list:
    """Return a deduplicated, sorted list of EQPT_ID_NAME values from the
    EQPT_MASTER tab of the spreadsheet. Returns [] silently on failure so
    the rest of the dashboard keeps working."""
    try:
        live = fetch_live_data(tab="EQPT_MASTER", force=force)
    except Exception as e:
        print(f"[eqpt_master] fetch failed: {e}")
        return []
    rows = live.get("rows", []) if isinstance(live, dict) else []
    key = _find_eqpt_key(rows)
    if not key:
        return []
    seen = set()
    out = []
    for r in rows:
        v = r.get(key)
        if v is None:
            continue
        s = str(v).strip()
        if not s or s in seen:
            continue
        seen.add(s)
        out.append(s)
    out.sort()
    return out


def _find_hod_key(rows):
    """Detect the HOD identifier column in the HOD MASTER tab."""
    if not rows:
        return None
    keys = list(rows[0].keys())
    candidates = ["HOD_NAME", "HOD NAME", "HOD", "HOD_ID_NAME", "HOD_FULL_NAME"]
    for c in candidates:
        if c in keys:
            return c
    for k in keys:
        ku = k.upper().replace("_", " ")
        if "HOD" in ku:
            return k
    return None


def fetch_hod_master(force: bool = False) -> list:
    """Return a deduplicated, sorted list of HOD names from the HOD MASTER tab.
    Tries multiple tab-name variants since Google Sheets tab names can include
    spaces."""
    rows = []
    last_err = None
    for tab_name in ("HOD MASTER", "HOD_MASTER", "HODMASTER"):
        try:
            live = fetch_live_data(tab=tab_name, force=force)
        except Exception as e:
            last_err = e
            continue
        candidate = live.get("rows", []) if isinstance(live, dict) else []
        if candidate:
            rows = candidate
            break
    if not rows:
        if last_err:
            print(f"[hod_master] fetch failed: {last_err}")
        return []
    key = _find_hod_key(rows)
    if not key:
        return []
    seen = set()
    out = []
    for r in rows:
        v = r.get(key)
        if v is None:
            continue
        s = str(v).strip()
        if not s or s in seen:
            continue
        seen.add(s)
        out.append(s)
    out.sort()
    return out


def _num(x):
    try:
        return float(x or 0)
    except Exception:
        return 0.0


def _find_net_key(rows):
    """Source ERP column for SFG Net Order Qty may be named slightly differently.
    Probe the first row for the closest match and use it consistently."""
    if not rows:
        return "SFG_NET_ORD_QTY"
    candidates = [
        "SFG_NET_ORD_QTY", "SFG_NET_ORDQTY", "SFG_NET_ORDERQTY",
        "SFG_NET_ORDER_QTY", "SFG_NETORDQTY",
    ]
    keys = list(rows[0].keys())
    for c in candidates:
        if c in keys:
            return c
    # Fallback: anything that looks like SFG net qty
    for k in keys:
        ku = k.upper().replace(" ", "").replace("_", "")
        if "SFG" in ku and "NET" in ku and ("ORD" in ku or "QTY" in ku):
            return k
    return "SFG_NET_ORD_QTY"  # not found — sums will stay 0


def _find_target_date_key(rows):
    """Source ERP column for the per-line target production closure date."""
    if not rows:
        return None
    keys = list(rows[0].keys())
    candidates = [
        "TARGET DATE OF PROD CLOSER",
        "TARGET_DATE_OF_PROD_CLOSER",
        "TARGET_DATE_PROD_CLOSER",
        "TARGET DATE PROD CLOSER",
    ]
    for c in candidates:
        if c in keys:
            return c
    for k in keys:
        ku = k.upper().replace("_", " ")
        if "TARGET" in ku and "DATE" in ku and ("CLOSER" in ku or "CLOSURE" in ku or "PROD" in ku):
            return k
    return None


def build_pivot(rows, net_key=None):
    """Group rows by SFG_BOM + SFG_CODE + SFG_NAME, sum SFG_REQUIRED_AG_BAL TO PRD FG
    and SFG_NET_ORD_QTY."""
    REQ_KEY = "SFG_REQUIRED_AG_BAL TO PRD FG"
    NET_KEY = net_key or _find_net_key(rows)
    groups = defaultdict(lambda: {
        "sum": 0, "net_sum": 0, "rows": 0, "um": "", "l3": "", "l4": "",
        "customers": set(), "mc_nos": set(), "ports": set()
    })
    for r in rows:
        bom = str(r.get("SFG_BOM") or "").strip()
        code = str(r.get("SFG_CODE") or "").strip()
        name = str(r.get("SFG_NAME") or "").strip()
        v = _num(r.get(REQ_KEY))
        nv = _num(r.get(NET_KEY))
        k = (bom, code, name)
        g = groups[k]
        g["sum"] += v
        # Always remember UM/L3/L4 so the BOM row has the right labels even if its
        # rows are all zero-requirement.
        g["um"] = r.get("SFG_UM") or g["um"]
        g["l3"] = r.get("SFG_LEVEL_3_NAME") or g["l3"]
        g["l4"] = r.get("SFG_LEVEL_4_NAME") or g["l4"]
        # LINES, MC_COUNT, customers, ports, NET_SUM only reflect rows that
        # actually have an outstanding SFG requirement — keeps the pivot row
        # consistent with the drill-down (which hides SFG_REQ <= 0 lines).
        if v > 0:
            g["rows"] += 1
            g["net_sum"] += nv
            if r.get("ACC_NAME"):
                g["customers"].add(str(r["ACC_NAME"]).strip())
            if r.get("MC_NO"):
                g["mc_nos"].add(str(r["MC_NO"]).strip())
            if r.get("PORD_NO"):
                g["ports"].add(str(r["PORD_NO"]).strip())

    pivot = []
    for (bom, code, name), g in groups.items():
        pivot.append({
            "SFG_BOM": bom,
            "SFG_CODE": code,
            "SFG_NAME": name,
            "REQ_SUM": round(g["sum"], 4),
            "NET_SUM": round(g["net_sum"], 4),
            "UM": g["um"],
            "L3": g["l3"],
            "L4": g["l4"],
            "LINES": g["rows"],
            "CUSTOMERS": len(g["customers"]),
            "MC_COUNT": len(g["mc_nos"]),
            "PORD_COUNT": len(g["ports"]),
        })
    pivot.sort(key=lambda x: (x["SFG_BOM"] or "zzz", x["SFG_CODE"]))
    return pivot


def build_lines(rows, net_key=None, date_key=None):
    """Trim rows to columns needed for drill-down.
    Always emit SFG_NET_ORD_QTY and TARGET_DATE_OF_PROD_CLOSER under canonical
    keys regardless of the source-header spelling."""
    NET_KEY = net_key or _find_net_key(rows)
    DATE_KEY = date_key if date_key is not None else _find_target_date_key(rows)
    KEEP = [
        "PORD_NO", "PORD_DATE", "MC_NO", "ACC_NAME", "SFG_CODE", "SFG_NAME",
        "SFG_BOM", "SFG_UM", "FG_ITEM_NAME", "FG_LEVEL_4_NAME",
        "FG_NET_ORDERQTY", "FG_PRODUCTION_QTY", "FG_BALANCE_TO_PRODUCTION",
        "SFG_REQUIRED_AG_BAL TO PRD FG", "CONTRACT_VRNO",
    ]
    out = []
    for r in rows:
        row = {k: r.get(k) for k in KEEP}
        # Normalize to stable canonical keys the frontend expects.
        row["SFG_NET_ORD_QTY"] = r.get(NET_KEY)
        row["TARGET_DATE_OF_PROD_CLOSER"] = r.get(DATE_KEY) if DATE_KEY else None
        out.append(row)
    return out


# --------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("dashboard.html")


@app.route("/data")
def data():
    """Return the live pivot + lines JSON.  Pass ?refresh=1 to force re-pull."""
    force = request.args.get("refresh", "").lower() in ("1", "true", "yes")
    try:
        live = fetch_live_data(force=force)
    except Exception as e:
        return jsonify({"error": f"Apps Script fetch failed: {e}"}), 500
    if isinstance(live, dict) and live.get("error"):
        return jsonify(live), 500

    # Drop empty rows
    rows = [r for r in live.get("rows", []) if r.get("PORD_NO")]

    net_key = _find_net_key(rows)
    date_key = _find_target_date_key(rows)

    # ----- DIAGNOSTIC -----
    if rows:
        print("=" * 70)
        print(f"[diagnose] Got {len(rows)} non-empty rows from Apps Script")
        print(f"[diagnose] Resolved SFG net-qty column   = {net_key!r}")
        print(f"[diagnose] Resolved target-date column   = {date_key!r}")
        if date_key:
            nz = sum(1 for r in rows if r.get(date_key))
            sample = next((r.get(date_key) for r in rows if r.get(date_key)), None)
            print(f"[diagnose]   target-date column has {nz} rows with a value; sample={sample!r}")
        print("=" * 70)
    # ----------------------

    eqpt_master = fetch_eqpt_master(force=force)
    hod_master = fetch_hod_master(force=force)
    print(f"[diagnose] EQPT_MASTER loaded {len(eqpt_master)} distinct EQPT(s)")
    print(f"[diagnose] HOD_MASTER  loaded {len(hod_master)} distinct HOD(s)")

    return jsonify({
        "pivot": build_pivot(rows, net_key=net_key),
        "lines": build_lines(rows, net_key=net_key, date_key=date_key),
        "eqpt_master": eqpt_master,
        "hod_master": hod_master,
        "exported_at": live.get("exported_at"),
        "row_count": len(rows),
        "fetched_at": datetime.now().isoformat(),
    })


@app.route("/health")
def health():
    now = time.time()
    ages = {k: int(now - v["ts"]) for k, v in _cache.items()}
    return jsonify({"ok": True, "cache_ages_s": ages})


if __name__ == "__main__":
    print("=" * 70)
    print(" Laser SFG Production Planning Dashboard - local server")
    print("=" * 70)
    print(" Open in browser:  http://localhost:5000")
    print(" Live data is fetched from Apps Script on every Refresh click")
    print(" (Apps Script URL is set near the top of app.py)")
    print(" Press Ctrl+C to stop")
    print("=" * 70)
    app.run(host="0.0.0.0", port=5000, debug=False)
