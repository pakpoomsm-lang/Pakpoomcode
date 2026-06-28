"""Local companion server for daily_follow.html.

POST /api/run-zpp0059  — drives SAP (ZPP0059), upserts rows into SQLite,
                         returns aggregated progress to the page.
GET  /api/db-stats     — returns row count and newest timestamp from the DB.
GET  /*                — serves static files from the Daily follow folder.

Requires Python 3.8+ (sqlite3 is in the standard library).
Run on the Windows machine that has SAP GUI open and logged in.
Double-click Start_Daily_Follow.bat or: python serve_daily_follow.py
"""

import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import webbrowser
from collections import defaultdict
from datetime import date, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import openpyxl

import build_daily_follow as bdf

# ---------------------------------------------------------------------------
# Configuration  (edit these to match your machine)
# ---------------------------------------------------------------------------
# BIND_HOST = "0.0.0.0" makes the server reachable from other PCs on the LAN
# (e.g. http://10.235.117.215:8059). Override with the DAILY_FOLLOW_HOST /
# DAILY_FOLLOW_PORT environment variables (set DAILY_FOLLOW_HOST=127.0.0.1 to
# restrict access back to this machine only).
HOST = "127.0.0.1"                                    # local URL for the browser on this PC
BIND_HOST = os.environ.get("DAILY_FOLLOW_HOST", "0.0.0.0")
PORT = int(os.environ.get("DAILY_FOLLOW_PORT", "8059"))
ROOT = bdf.ROOT
OUTPUT_FILE = bdf.OUTPUT_FILE
PROGRESS_FILE = bdf.PROGRESS_FILE          # ZPP0059.xlsx — still kept for fallback

# ---------------------------------------------------------------------------
# SAP GUI recordings, embedded inline (no external .vbs files needed).
# The selection-screen field values are taken verbatim from the recordings;
# the date window in ZPP0059 is rolled automatically before each run.
# ---------------------------------------------------------------------------
SAP_SCRIPT_0059 = r'''If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUISERVER")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").resizeWorkingPane 153,29,false
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_SHOP-LOW").text = "542"
session.findById("wnd[0]/usr/ctxtS_ORDTY-LOW").text = "zp40"
session.findById("wnd[0]/usr/ctxtS_WKDT-LOW").text = "12.06.2026"
session.findById("wnd[0]/usr/ctxtS_WKDT-HIGH").text = "15.06.2026"
session.findById("wnd[0]/usr/ctxtS_WKDT-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtS_WKDT-HIGH").caretPosition = 10
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell 2,"ZTIMESTAMP"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "2"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/cmbGS_EXPORT-DESTINATION").setFocus
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
'''

SAP_SCRIPT_0022 = r'''If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUISERVER")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").resizeWorkingPane 194,25,false
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "1001"
session.findById("wnd[0]/usr/radR_PROC_3").setFocus
session.findById("wnd[0]/usr/radR_PROC_3").select
session.findById("wnd[0]/usr/ctxtS_LINE3-LOW").text = "542"
session.findById("wnd[0]/usr/ctxtS_LINE3-LOW").setFocus
session.findById("wnd[0]/usr/ctxtS_LINE3-LOW").caretPosition = 3
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]").sendVKey 33
' Choose the FULL layout "/PP ASSY ALL" by name instead of by row position.
' Row numbers shift whenever SAP layouts are added/removed (the recording
' happened to land on row 10 = "/TPP&PAINT", the short layout). We scan the
' layout list, match the technical name first, then the description, and only
' fall back to a fixed row if neither is found.
Dim lShell, lRow, lFound, lName, lText
Set lShell = session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
lFound = -1
On Error Resume Next
For lRow = 0 To lShell.RowCount - 1
   lName = ""
   lText = ""
   lName = Trim(lShell.GetCellValue(lRow, "VARIANT"))
   lText = Trim(lShell.GetCellValue(lRow, "TEXT"))
   If lName = "/PP ASSY ALL" Or lText = "PP ASSY ALL LINE (Defalt)" Then
      lFound = lRow
      Exit For
   End If
Next
On Error GoTo 0
If lFound = -1 Then lFound = 2
lShell.setCurrentCell lFound,"TEXT"
lShell.firstVisibleRow = lFound
lShell.selectedRows = CStr(lFound)
lShell.clickCurrentCell
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 5,"PSMNG"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "5"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
'''

# SQLite database path (same folder as the scripts).
DB_FILE = ROOT / "zpp0059.db"

# Persisted page state (imported rows + progress overlay) so the table survives
# a browser refresh. Written by POST /api/save-state, read by GET /api/state.
STATE_FILE = ROOT / "daily_follow_state.json"

# How long to wait for SAP to finish exporting (seconds).
RUN_TIMEOUT = 180

# How many times to drive SAP before giving up. A second try clears most
# transient hiccups (a stray popup, SAP momentarily busy) because the script
# re-navigates to /nZPP0059 from scratch each run. Timeouts are never retried.
RUN_ATTEMPTS = 2

# Rolling date window injected into the VBS before each run.
DYNAMIC_DATES = True
DATE_FROM_OFFSET_DAYS = -7   # S_WKDT-LOW  = today - 7 days
DATE_TO_OFFSET_DAYS   = 0    # S_WKDT-HIGH = today

# Navigate to ZPP0059 before touching its selection screen fields.
# Set to "" to run the recorded VBS exactly as-is.
START_TRANSACTION = "ZPP0059"

# --- ZPP0022 (order master / "Update Progress" source) -----------------------
# Same concept as ZPP0059, but pulls the order export that rebuilds the table.
# The recorded script already fills its own selection fields (plant 1001,
# line 542, radio R_PROC_3), so it runs as-is — no date window to roll.
START_TRANSACTION_0022 = "ZPP0022"
PROGRESS_0022_FILE     = ROOT / "ZPP0022.xlsx"   # latest export, served to the page
ZPP0022_TABLE          = "zpp0022_raw"

# AS/400 network-share base folder that SAP exports into. The mapped DRIVE
# LETTER differs per PC (some map it as J:, the shop-floor PC maps it as Y:).
# Override per machine with the SAP_EXPORT_BASE environment variable, e.g.
#     set SAP_EXPORT_BASE=J:\7.541_HEI\Database follow
SAP_EXPORT_BASE        = os.environ.get("SAP_EXPORT_BASE", r"Y:\7.541_HEI\Database follow")
SAP_EXPORT_DIR_0022    = str(Path(SAP_EXPORT_BASE) / "ZPP0022")

# --- First-lot check sheets --------------------------------------------------
# The first-lot inspection app stores each process in its own SQLite file under
# Server_firstlot. We read them (read-only) to show, per row, whether every
# process check sheet has been confirmed OK.
#
# Resolution order:
#   1. FIRSTLOT_DIR env variable — set this in Start_Daily_Follow.bat per machine.
#   2. Candidate list below — searched in order; first folder that contains at
#      least one expected DB wins. Covers drive-letter differences across PCs.
_FIRSTLOT_CANDIDATES = [
    Path(os.environ.get("FIRSTLOT_DIR", "")),               # env override (blank → skip)
    Path(r"W:\PD\2.HEAT INDOOR\13.Suphamat P\Server_firstlot"),
    Path(r"Y:\PD\2.HEAT INDOOR\13.Suphamat P\Server_firstlot"),
    Path(r"J:\PD\2.HEAT INDOOR\13.Suphamat P\Server_firstlot"),
    ROOT.parent.parent / "Server_firstlot",                 # dev: repo sibling folder
]
_PROBE_DB = "cutting_records.db"   # existence check — all 4 DBs live in the same folder


def _resolve_firstlot_dir():
    for p in _FIRSTLOT_CANDIDATES:
        if p and p.is_dir() and (p / _PROBE_DB).exists():
            return p
    # Fall back to env/default even if DB not found; error will surface at query time.
    return _FIRSTLOT_CANDIDATES[0] if os.environ.get("FIRSTLOT_DIR") else _FIRSTLOT_CANDIDATES[-1]


FIRSTLOT_DIR = _resolve_firstlot_dir()

# (process, db file, table, line col, seq col, month col or None, result col, ts col)
CHECKSHEET_SOURCES = [
    ("cutting",   "cutting_records.db",   "cutting_records", "line", "seq", None,         "status",      "saved_at"),
    ("fp",        "fp_records.db",        "fp_records",      "line", "seq", "prod_month", "final",       "saved_at"),
    ("hp",        "hp_records.db",        "hp_records",      "line", "seq", "prod_month", "result",      "saved_at"),
    ("hp_insert", "hp_insert_records.db", "records",         "line", "seq", "prod_month", "final_check", "created_at"),
]


def _cs_month(value):
    """Check-sheet prod_month 'MM/YYYY' -> 'M.YYYY' to match build's month_display."""
    raw = str(value or "").strip()
    if "/" not in raw:
        return ""
    mm, _, yy = raw.partition("/")
    try:
        return f"{int(mm)}.{yy}"
    except ValueError:
        return ""


def _ts_month(value):
    """ISO timestamp 'YYYY-MM-DDT...' -> 'M.YYYY' (used for cutting which has no prod_month)."""
    raw = str(value or "").strip()[:7]   # 'YYYY-MM'
    if len(raw) < 7 or raw[4] != "-":
        return ""
    try:
        return f"{int(raw[5:7])}.{raw[:4]}"
    except ValueError:
        return ""


def load_checksheets():
    """Read first-lot check sheets, keep the latest OK/NG per (key, process).

    Returns by_key["LINE|month|seq"] = {process: "OK"/"NG"} for all processes.
    cutting_records has no prod_month column, so we derive the month from saved_at
    to prevent records from one month matching orders in another month.
    """
    by_key = {}
    for proc, fname, table, lcol, scol, mcol, rcol, tcol in CHECKSHEET_SOURCES:
        db = FIRSTLOT_DIR / fname
        if not db.is_file() or db.stat().st_size == 0:
            continue
        # Always fetch saved_at/created_at as the last column for month fallback.
        cols = f"{lcol}, {scol}, {rcol}, {mcol}, {tcol}" if mcol else f"{lcol}, {scol}, {rcol}, {tcol}"
        try:
            conn = sqlite3.connect(f"file:{db.as_posix()}?mode=ro", uri=True)
            # Oldest first so a later record overwrites (latest result wins).
            cur = conn.execute(f"SELECT {cols} FROM {table} ORDER BY {tcol}")
            for r in cur:
                line = str(r[0] or "").strip().upper()
                seq = bdf.seq_key(r[1])
                res = str(r[2] or "").strip().upper()
                if not line or not seq or res not in ("OK", "NG"):
                    continue
                if mcol:
                    month = _cs_month(r[3])   # r[3] = prod_month, r[4] = ts
                else:
                    month = _ts_month(r[3])   # r[3] = saved_at (no prod_month col)
                if month:
                    by_key.setdefault(f"{line}|{month}|{seq}", {})[proc] = res
            conn.close()
        except sqlite3.Error:
            continue  # skip a missing/locked/corrupt DB rather than break the page
    return by_key


# --- Part incoming (HEI Smart Stock Management) ------------------------------
# The warehouse receiving app stores every scanned-in part in incoming.db. We
# read it (read-only) to show, per row, which parts have physically arrived for
# that lot. Same idea/keying as the check sheets above.
#
# Resolution: STOCK_DB env first, then candidate paths (first existing wins).
_STOCK_DB_CANDIDATES = [
    Path(os.environ.get("STOCK_DB", "")),                    # env override (blank → skip)
    Path(r"W:\PD\2.HEAT INDOOR\13.Suphamat P\HEI Smart Stock Management\incoming.db"),
    Path(r"Y:\PD\2.HEAT INDOOR\13.Suphamat P\HEI Smart Stock Management\incoming.db"),
    Path(r"J:\PD\2.HEAT INDOOR\13.Suphamat P\HEI Smart Stock Management\incoming.db"),
    ROOT.parent.parent / "HEI Smart Stock Management" / "incoming.db",  # dev: repo sibling
]


def _resolve_stock_db():
    for p in _STOCK_DB_CANDIDATES:
        if p and p.is_file():
            return p
    return _STOCK_DB_CANDIDATES[0] if os.environ.get("STOCK_DB") else _STOCK_DB_CANDIDATES[-1]


STOCK_DB = _resolve_stock_db()


def _stock_month(value):
    """incoming pro_month 'MMYYYY' (e.g. '062026') -> 'M.YYYY' to match build."""
    raw = str(value or "").strip()
    if len(raw) != 6 or not raw.isdigit():
        return ""
    try:
        return f"{int(raw[:2])}.{raw[2:]}"
    except ValueError:
        return ""


def load_incoming():
    """Read incoming.db, group received parts per (line, month, seq).

    Returns one map the page joins onto its rows:
      by_key["LINE|month|seq"] = [{part, desc, partType, qty, last}]
    """
    by_key = {}
    if not STOCK_DB.is_file() or STOCK_DB.stat().st_size == 0:
        return by_key
    try:
        conn = sqlite3.connect(f"file:{STOCK_DB.as_posix()}?mode=ro", uri=True)
        # Load description lookup first (small table, fits in memory).
        desc_map = {}  # part_no -> (description, part_type)
        try:
            for item_code, description, part_type in conn.execute(
                "SELECT item_code, description, part_type FROM item_descriptions"
            ):
                desc_map[str(item_code or "").strip()] = (
                    str(description or "").strip(),
                    str(part_type or "").strip(),
                )
        except sqlite3.Error:
            pass  # table missing → descriptions stay blank
        cur = conn.execute(
            "SELECT line_num, seq, pro_month, part_no, qty, receive_date "
            "FROM incoming ORDER BY receive_date, receive_time"
        )
        # Sum qty per (key, part); keep the newest receive_date seen.
        parts = {}  # key -> {part: {"qty": n, "last": date}}
        for line_num, seq, pro_month, part_no, qty, receive_date in cur:
            line = str(line_num or "").strip().upper()
            seq_n = bdf.seq_key(seq)
            month = _stock_month(pro_month)
            part = str(part_no or "").strip()
            if not line or not seq_n or not month or not part:
                continue
            key = f"{line}|{month}|{seq_n}"
            q = qty if isinstance(qty, (int, float)) else 0
            bucket = parts.setdefault(key, {})
            entry = bucket.setdefault(part, {"qty": 0, "last": ""})
            entry["qty"] += q
            if receive_date:
                entry["last"] = str(receive_date)
        conn.close()
        for key, bucket in parts.items():
            by_key[key] = [
                {
                    "part": p,
                    "desc": desc_map.get(p, ("", ""))[0],
                    "partType": desc_map.get(p, ("", ""))[1],
                    "qty": bdf.clean_qty(v["qty"]),
                    "last": v["last"],
                }
                for p, v in bucket.items()
            ]
    except sqlite3.Error:
        return {}  # missing/locked/corrupt DB -> blank column, page still works
    return by_key


# --- OT program bridge -------------------------------------------------------
# The OT recording app (Flask) runs on the same PC. We proxy its working-hours
# API server-side so the Daily Follow page can read per-shift hours without
# running into cross-origin (CORS) restrictions. Override with the OT_APP_URL
# environment variable if the OT app uses a different host/port.
OT_APP_URL = os.environ.get("OT_APP_URL", "http://127.0.0.1:5000")


# SAP opens the exported file in Excel automatically. When True, the workbook
# is closed again right after we have read it (only that file is closed; any
# other Excel windows you have open are left alone).
CLOSE_EXCEL_AFTER = True

# Folders watched for the file SAP exports (searched in order, newest file wins).
HOME = Path(os.path.expanduser("~"))
EXPORT_DIRS = [
    Path(SAP_EXPORT_BASE) / "ZPP0059",              # network share (primary)
    Path(SAP_EXPORT_BASE) / "ZPP0022",              # ZPP0022 order export
    # Also scan the J: variant in case a PC maps the same share as J:.
    Path(r"J:\7.541_HEI\Database follow\ZPP0059"),
    Path(r"J:\7.541_HEI\Database follow\ZPP0022"),
    Path(r"C:\TEMP"),                               # SAP default save dir
    ROOT,
    HOME / "Documents" / "SAP" / "SAP GUI",
    HOME / "Downloads",
    HOME / "Documents",
    HOME / "Desktop",
]
# Target directory for SAP to save the export file into.
# Must match one of the EXPORT_DIRS entries above.
SAP_EXPORT_DIR = str(Path(SAP_EXPORT_BASE) / "ZPP0059")

# Business key columns — a row is considered duplicate when ALL of these match.
# If any column is missing from the export, falls back to SHA-256 of the whole row.
UNIQUE_KEY_COLS = ["Order", "Activity", "Short Time Stamp", "Tag ID"]

# Columns used to aggregate progress (must match bdf.load_progress logic).
PROGRESS_COLS = [
    "Production Line", "Production Month", "Sequence",
    "Material", "Assembly Order", "Operation Short Text", "Posted Quantity",
]

# File extensions accepted when scanning EXPORT_DIRS for SAP's export file.
EXPORT_EXTS = {".xlsx", ".xlsm", ".xls"}


# ---------------------------------------------------------------------------
# SQLite helpers
# ---------------------------------------------------------------------------
def _get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_FILE), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def _col_name(raw: str) -> str:
    """Sanitise a header string to a safe SQL column name."""
    return re.sub(r"[^\w]", "_", raw.strip()).strip("_")


def _row_hash(values) -> str:
    blob = "|".join("" if v is None else str(v) for v in values)
    return hashlib.sha256(blob.encode("utf-8")).hexdigest()


def db_init(conn: sqlite3.Connection, headers: list[str],
            table: str = "zpp0059_raw") -> list[str]:
    """Ensure the table exists and has all required columns. Returns safe col names."""
    safe = [_col_name(h) for h in headers]
    # Base table always has the two internal columns; data columns are added
    # below via ALTER (handles both new tables and empty-headers calls cleanly).
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS {table} (
            _row_hash TEXT PRIMARY KEY,
            _inserted_at TEXT DEFAULT (datetime('now','localtime'))
        )
    """)
    # Add any columns that appeared in a newer export but not in the original table.
    existing = {r[1] for r in conn.execute(f"PRAGMA table_info({table})")}
    for c in safe:
        if c not in existing:
            conn.execute(f'ALTER TABLE {table} ADD COLUMN "{c}" TEXT')
    # Index the production-date columns so api/actual-production does an index
    # range scan instead of a full-table scan (the table grows unbounded).
    for date_header in ("Working day", "Posting Date"):
        dc = _col_name(date_header)
        if dc in existing or dc in safe:
            conn.execute(f'CREATE INDEX IF NOT EXISTS "idx_{dc}" ON {table}("{dc}")')
    conn.commit()
    return safe


def db_upsert(conn: sqlite3.Connection, headers: list[str], rows,
              table: str = "zpp0059_raw") -> tuple[int, int]:
    """Upsert rows keyed by the business key (latest export wins). Returns (new, updated).

    A row with the same business key (Order+Activity+Short Time Stamp+Tag ID)
    overwrites the stored copy, so a later edit or cancellation in SAP replaces
    the earlier values instead of being ignored. _inserted_at is left untouched
    on update, so it still marks when the row was first seen.
    """
    safe_headers = db_init(conn, headers, table)

    # Determine which columns to use for the business key.
    key_cols = [c for c in UNIQUE_KEY_COLS if c in headers]
    use_hash = len(key_cols) < 2  # fallback: SHA-256 of entire row

    placeholders = ", ".join("?" for _ in safe_headers)
    col_list = ", ".join(f'"{c}"' for c in safe_headers)
    set_clause = ", ".join(f'"{c}" = excluded."{c}"' for c in safe_headers)
    sql = (f'INSERT INTO {table} (_row_hash, {col_list}) VALUES (?, {placeholders}) '
           f'ON CONFLICT(_row_hash) DO UPDATE SET {set_clause}')

    # Pre-load existing keys so we can report new vs updated honestly: the upsert
    # itself counts both as a change, so rowcount alone can't tell them apart.
    seen = {r[0] for r in conn.execute(f"SELECT _row_hash FROM {table}")}

    new = updated = 0
    batch = []
    for row in rows:
        values = [str(v).strip() if v is not None else "" for v in row]
        if use_hash:
            key = _row_hash(values)
        else:
            key_vals = [values[headers.index(c)] for c in key_cols]
            key = hashlib.sha256("|".join(key_vals).encode()).hexdigest()
        if key in seen:
            updated += 1
        else:
            new += 1
            seen.add(key)
        batch.append((key, *values))
        if len(batch) >= 500:
            conn.executemany(sql, batch)
            batch.clear()

    if batch:
        conn.executemany(sql, batch)
    conn.commit()
    return new, updated


def db_aggregate(conn: sqlite3.Connection) -> tuple[dict, dict]:
    """Re-aggregate the full DB into (progressMat, progressLot) for the page."""
    # Map header -> safe col name.
    col_map = {r[1]: r[1] for r in conn.execute("PRAGMA table_info(zpp0059_raw)")}

    def c(name):
        return _col_name(name)

    # Check all needed columns exist.
    needed = [c(x) for x in PROGRESS_COLS]
    missing = [n for n in needed if n not in col_map]
    if missing:
        return {}, {}

    op_field = {"Insert": "fp", "Brazing": "auto", "Cutting": "cutting"}
    hp_op = "H/P bender"

    def clean(v):
        if v is None:
            return ""
        v = str(v).strip()
        try:
            f = float(v)
            return int(f) if f.is_integer() else round(f, 3)
        except ValueError:
            return v

    by_mat: dict = defaultdict(lambda: defaultdict(float))
    by_lot: dict = defaultdict(float)

    # Exclude cancelled / deleted postings so the aggregate doesn't over-count.
    # These columns may be absent in an older table, so only filter when present.
    where = [f'"{c("Production Line")}" != ""']
    for name in ("Deletion Flag", "Cancel Date"):
        cc = c(name)
        if cc in col_map:
            where.append(f'IFNULL("{cc}", "") = ""')
    sql = (f'SELECT "{c("Production Line")}", "{c("Production Month")}", '
           f'"{c("Sequence")}", "{c("Material")}", "{c("Assembly Order")}", '
           f'"{c("Operation Short Text")}", "{c("Posted Quantity")}" '
           f'FROM zpp0059_raw WHERE ' + " AND ".join(where))

    for row in conn.execute(sql):
        line, month_raw, seq_raw, mat, ao, op, qty_raw = row
        if not line:
            continue
        try:
            qty = float(qty_raw) if qty_raw else 0.0
        except ValueError:
            qty = 0.0
        head = f"{line}|{bdf.month_display(month_raw)}|{bdf.seq_key(seq_raw)}"
        if op == hp_op:
            by_lot[f"{head}|{ao}"] += qty
        elif op in op_field:
            by_mat[f"{head}|{mat}"][op_field[op]] += qty

    def rnd(v):
        return int(v) if float(v).is_integer() else round(v, 3)

    mat_out = {k: {f: rnd(q) for f, q in fields.items()} for k, fields in by_mat.items()}
    lot_out = {k: rnd(q) for k, q in by_lot.items()}
    return mat_out, lot_out


def actual_production_by_date(date_iso):
    """Final-op qty (Brazing->auto, Cutting->cutting) produced on one working day,
    grouped by shift D/N and by line|month|seq|material key."""
    conn = _get_db()
    col_map = {r[1]: r[1] for r in conn.execute("PRAGMA table_info(zpp0059_raw)")}
    c = _col_name
    needed = ["Production Line", "Production Month", "Sequence", "Material",
              "Operation Short Text", "Posted Quantity", "Shift D/N"]
    if any(c(x) not in col_map for x in needed):
        conn.close(); return {"D": {}, "N": {}}
    date_col = c("Working day") if c("Working day") in col_map else (
        c("Posting Date") if c("Posting Date") in col_map else None)
    if not date_col:
        conn.close(); return {"D": {}, "N": {}}
    # Range scan (index-friendly) instead of LIKE: stored timestamps sort
    # lexicographically like their ISO date prefix.
    try:
        next_day = (date.fromisoformat(date_iso) + timedelta(days=1)).isoformat()
        where = [f'"{date_col}" >= ? AND "{date_col}" < ?']
        params = [date_iso, next_day]
    except ValueError:
        where = [f'"{date_col}" LIKE ?']
        params = [date_iso + "%"]
    for name in ("Deletion Flag", "Cancel Date"):
        cc = c(name)
        if cc in col_map:
            where.append(f'IFNULL("{cc}", "") = ""')
    sql = (f'SELECT "{c("Production Line")}", "{c("Production Month")}", "{c("Sequence")}", '
           f'"{c("Material")}", "{c("Operation Short Text")}", "{c("Posted Quantity")}", '
           f'"{c("Shift D/N")}" FROM zpp0059_raw WHERE ' + " AND ".join(where))
    op_field = {"Brazing": "auto", "Cutting": "cutting"}
    out = {"D": {}, "N": {}}
    for line, month_raw, seq_raw, mat, op, qty_raw, dn in conn.execute(sql, params):
        if not line:
            continue
        field = op_field.get((op or "").strip())
        if not field:
            continue
        dn = (dn or "").strip().upper()
        if dn not in ("D", "N"):
            continue
        try:
            qty = float(qty_raw) if qty_raw else 0.0
        except ValueError:
            qty = 0.0
        key = f"{line}|{bdf.month_display(month_raw)}|{bdf.seq_key(seq_raw)}|{mat}"
        bucket = out[dn].setdefault(key, {"auto": 0.0, "cutting": 0.0})
        bucket[field] += qty
    conn.close()
    def rnd(v):
        return int(v) if float(v).is_integer() else round(v, 3)
    for dn in out:
        for k in out[dn]:
            for f in out[dn][k]:
                out[dn][k][f] = rnd(out[dn][k][f])
    return out


def db_stats(conn: sqlite3.Connection, table: str = "zpp0059_raw") -> dict:
    total = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
    newest = conn.execute(
        f"SELECT MAX(_inserted_at) FROM {table}"
    ).fetchone()[0]
    return {"total_rows": total, "newest_inserted_at": newest}


# Columns shown (in order) by the "View Database" viewer.
VIEWER_COLS = [
    "Production Line", "Production Month", "Sequence", "Order", "Activity",
    "Material", "Material Description", "Operation Short Text",
    "Posted Quantity", "Assembly Order", "Working day", "Short Time Stamp",
    "Receive Date", "Receive Time", "Stock Type", "Process",
]


def db_query(conn: sqlite3.Connection, q: str = "", limit: int = 500,
             offset: int = 0, filters: dict | None = None) -> dict:
    """Return rows from the DB for the viewer.

    `q` is a free-text search OR-ed across all visible columns.
    `filters` maps a visible-column index -> substring; each is AND-ed
    (and AND-ed with `q`) so the viewer can narrow one column at a time."""
    existing = {r[1] for r in conn.execute("PRAGMA table_info(zpp0059_raw)")}
    cols = [c for c in VIEWER_COLS if _col_name(c) in existing]
    if not cols:  # fall back to whatever columns exist (minus internal)
        cols = [c for c in existing if not c.startswith("_")][:16]
        safe = cols
        headers = cols
    else:
        safe = [_col_name(c) for c in cols]
        headers = cols
    col_sql = ", ".join(f'"{s}"' for s in safe)

    where_parts, params = [], []
    if q:
        like = f"%{q}%"
        where_parts.append("(" + " OR ".join(f'"{s}" LIKE ?' for s in safe) + ")")
        params += [like] * len(safe)
    if filters:
        for i, val in filters.items():
            if 0 <= i < len(safe) and val:
                where_parts.append(f'"{safe[i]}" LIKE ?')
                params.append(f"%{val}%")
    where = "WHERE " + " AND ".join(where_parts) if where_parts else ""

    total = conn.execute(
        f"SELECT COUNT(*) FROM zpp0059_raw {where}", params
    ).fetchone()[0]

    order_col = _col_name("Short Time Stamp")
    order_sql = f'ORDER BY "{order_col}" DESC' if order_col in existing else ""
    rows = conn.execute(
        f'SELECT {col_sql} FROM zpp0059_raw {where} {order_sql} LIMIT ? OFFSET ?',
        params + [limit, offset],
    ).fetchall()
    return {
        "columns": headers,
        "rows": [list(r) for r in rows],
        "total": total,
        "shown": len(rows),
        "offset": offset,
        "limit": limit,
    }


# ---------------------------------------------------------------------------
# SAP + export helpers
# ---------------------------------------------------------------------------
def _prepare_script(script_text: str = SAP_SCRIPT_0059,
                    transaction: str = START_TRANSACTION,
                    roll_dates: bool = DYNAMIC_DATES,
                    export_dir_win: str = SAP_EXPORT_DIR,
                    tmp_name: str = "_zpp0059_run.vbs") -> Path:
    try:
        text = script_text

        # 1) Roll the work-date window (only for scripts that have S_WKDT fields).
        if roll_dates:
            low  = (date.today() + timedelta(days=DATE_FROM_OFFSET_DAYS)).strftime("%d.%m.%Y")
            high = (date.today() + timedelta(days=DATE_TO_OFFSET_DAYS)).strftime("%d.%m.%Y")
            text = re.sub(r'(S_WKDT-LOW"\)\.text\s*=\s*")[^"]*(")',  rf"\g<1>{low}\g<2>",  text)
            text = re.sub(r'(S_WKDT-HIGH"\)\.text\s*=\s*")[^"]*(")', rf"\g<1>{high}\g<2>", text)

        # 2) Navigate to the transaction before touching its selection fields.
        if transaction and "/tbar[0]/okcd" not in text:
            nav = (f'session.findById("wnd[0]/tbar[0]/okcd").text = "/n{transaction}"\r\n'
                   f'session.findById("wnd[0]").sendVKey 0\r\n')
            m = re.search(r'^\s*session\.findById\("wnd\[0\]"\)\.resizeWorkingPane', text, re.M)
            if not m:
                m = re.search(r'^\s*session\.findById\("wnd\[0\]', text, re.M)
            if m:
                text = text[:m.start()] + nav + text[m.start():]

        # 3) Set export directory + filename in the SAP "Save File" dialog.
        #    SAP uses wnd[1]/usr/ctxtDY_PATH and ctxtDY_FILENAME for the local file dialog.
        #    We inject these lines just before the final wnd[1] btn[0] (Generate) press.
        export_dir  = export_dir_win.replace("\\", "\\\\")
        save_inject = (
            f'On Error Resume Next\r\n'
            f'session.findById("wnd[1]/usr/ctxtDY_PATH").text = "{export_dir}"\r\n'
            f'session.findById("wnd[2]/usr/ctxtDY_PATH").text = "{export_dir}"\r\n'
            f'On Error GoTo 0\r\n'
        )
        # Find the last wnd[1]/tbar[0]/btn[0] press (the Generate button).
        gen_pat = re.compile(
            r'^(\s*session\.findById\("wnd\[1\]/tbar\[0\]/btn\[0\]"\)\.press)', re.M)
        last = None
        for m in gen_pat.finditer(text):
            last = m
        if last:
            text = text[:last.start()] + save_inject + text[last.start():]

        tmp = ROOT / tmp_name
        tmp.write_text(text, encoding="utf-8")
        return tmp
    except Exception:
        # Fall back to the unmodified recording.
        tmp = ROOT / tmp_name
        tmp.write_text(script_text, encoding="utf-8")
        return tmp


def _start_security_handler() -> subprocess.Popen | None:
    """Launch a background PowerShell that watches for every 'SAP GUI Security'
    popup and clicks its 'Allow' button via UI Automation — focus-independent,
    so it does not depend on accelerators or tab order. Written to a .ps1 file
    and run with -File (passing multi-line scripts via -Command is unreliable)."""
    ps = shutil.which("powershell") or shutil.which("pwsh")
    if not ps:
        return None

    script = r"""
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$CT = [System.Windows.Automation.ControlType]
$Desc = [System.Windows.Automation.TreeScope]::Descendants
$Child = [System.Windows.Automation.TreeScope]::Children
$TrueCond = [System.Windows.Automation.Condition]::TrueCondition
$Toggle = [System.Windows.Automation.TogglePattern]::Pattern
$Invoke = [System.Windows.Automation.InvokePattern]::Pattern
$On = [System.Windows.Automation.ToggleState]::On

$deadline = (Get-Date).AddSeconds(160)
while ((Get-Date) -lt $deadline) {
    $root = $AE::RootElement
    $win = $null
    foreach ($w in $root.FindAll($Child, $TrueCond)) {
        if ($w.Current.Name -like '*SAP GUI Security*') { $win = $w; break }
    }
    if ($win -ne $null) {
        # Tick "Remember My Decision" so it stops asking next time.
        $cbCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::CheckBox)
        $cb = $win.FindFirst($Desc, $cbCond)
        if ($cb -ne $null) {
            try {
                $tp = $cb.GetCurrentPattern($Toggle)
                if ($tp.Current.ToggleState -ne $On) { $tp.Toggle() }
            } catch {}
        }
        # Click the Allow button by name (ignore any '&' accelerator marker).
        $btnCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Button)
        foreach ($b in $win.FindAll($Desc, $btnCond)) {
            $name = ($b.Current.Name) -replace '&',''
            if ($name -match '^(Allow|Zulassen|Permitir)') {
                try { $b.GetCurrentPattern($Invoke).Invoke() } catch {}
                break
            }
        }
        Start-Sleep -Milliseconds 300
    }
    Start-Sleep -Milliseconds 200
}
"""
    tmp = ROOT / "_sec_handler.ps1"
    tmp.write_text(script, encoding="utf-8")
    try:
        return subprocess.Popen(
            [ps, "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-WindowStyle", "Hidden", "-File", str(tmp)],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
    except Exception:
        return None


def _close_excel(filename: str) -> None:
    """Close the workbook SAP auto-opened after the export, without saving.
    Only the matching file is closed; other open Excel windows are untouched.
    Excel itself is quit only if no workbooks remain."""
    if not CLOSE_EXCEL_AFTER:
        return
    wscript = shutil.which("wscript")
    if not wscript:
        return
    safe_name = filename.replace('"', '')
    vbs = f'''
On Error Resume Next
Dim target : target = "{safe_name}"
Dim i, xl, wb, closed
closed = False
For i = 1 To 30
    Set xl = GetObject(, "Excel.Application")
    If Not (xl Is Nothing) Then
        For Each wb In xl.Workbooks
            If LCase(wb.Name) = LCase(target) Then
                wb.Close False
                closed = True
            End If
        Next
        If closed Then
            If xl.Workbooks.Count = 0 Then xl.Quit
            Exit For
        End If
    End If
    WScript.Sleep 500
Next
'''
    tmp = ROOT / "_close_excel.vbs"
    try:
        tmp.write_text(vbs, encoding="utf-8")
        subprocess.Popen(
            [wscript, "//nologo", str(tmp)],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
    except Exception:
        pass


def _newest_export(since: float, dirs: list[Path] | None = None) -> Path | None:
    # When the caller knows exactly where SAP just saved the export (e.g. the
    # ZPP0022 order folder), search only that folder. Otherwise the globally
    # newest file across EXPORT_DIRS could be from a *different* transaction
    # (e.g. a ZPP0059 stock export), which then gets handed back as the wrong
    # file and rejected downstream by the header check.
    search_dirs = dirs if dirs is not None else EXPORT_DIRS
    best, best_mtime = None, since - 2
    for d in search_dirs:
        try:
            if not d.is_dir():
                continue
            for f in d.iterdir():
                if f.suffix.lower() not in EXPORT_EXTS:
                    continue
                if f.resolve() in (PROGRESS_FILE.resolve(), PROGRESS_0022_FILE.resolve()):
                    continue
                m = f.stat().st_mtime
                if m >= best_mtime:
                    best, best_mtime = f, m
        except OSError:
            continue
    return best


def _read_xlsx(path: Path) -> tuple[list[str], list]:
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    headers = [str(v).strip() if v is not None else f"col{i}" for i, v in enumerate(next(it))]
    rows = list(it)
    wb.close()
    return headers, rows


class SapNotReady(RuntimeError):
    """SAP GUI is not open / scripting off / not logged in — cannot export yet."""


# Short reason code (from the probe) -> message shown to the user.
_SAP_NOT_READY_MSG = {
    "NO_SAPGUI": "ไม่พบ SAP GUI — กรุณาเปิด SAP Logon แล้ว login ก่อนกดปุ่มนี้ "
                 "(SAP GUI is not open. Open SAP Logon and log in first.)",
    "NO_ENGINE": "SAP GUI Scripting ยังไม่ได้เปิดใช้งาน — เปิด scripting แล้ว login ก่อน "
                 "(SAP GUI Scripting is disabled. Enable it and log in first.)",
    "NO_CONNECTION": "ยังไม่ได้ login เข้า SAP — กรุณา login ก่อนกดปุ่มนี้ "
                     "(Not logged in to SAP. Please log in first.)",
    "NO_SESSION": "ยังไม่ได้ login เข้า SAP — กรุณา login ก่อนกดปุ่มนี้ "
                  "(Not logged in to SAP. Please log in first.)",
    "NOT_LOGGED_IN": "ยังไม่ได้ login เข้า SAP — กรุณา login ให้เรียบร้อยก่อนกดปุ่มนี้ "
                     "(Not logged in to SAP. Please log in first.)",
}


def _sap_session_ready() -> tuple[bool, str]:
    """Probe for a logged-in SAP GUI session. Returns (ready, reason_code).
    reason_code is 'OK' when ready, otherwise one of the keys in _SAP_NOT_READY_MSG."""
    cscript = shutil.which("cscript")
    if not cscript:
        return False, "NO_SAPGUI"
    probe = (
        'On Error Resume Next\r\n'
        'Dim g, app, conn, sess\r\n'
        'Set g = GetObject("SAPGUI")\r\n'
        'If g Is Nothing Then Err.Clear : Set g = GetObject("SAPGUISERVER")\r\n'
        'If g Is Nothing Then WScript.Echo "NO_SAPGUI" : WScript.Quit 0\r\n'
        'Set app = g.GetScriptingEngine\r\n'
        'If app Is Nothing Then WScript.Echo "NO_ENGINE" : WScript.Quit 0\r\n'
        'If app.Children.Count = 0 Then WScript.Echo "NO_CONNECTION" : WScript.Quit 0\r\n'
        'Set conn = app.Children(0)\r\n'
        'If conn.Children.Count = 0 Then WScript.Echo "NO_SESSION" : WScript.Quit 0\r\n'
        'Set sess = conn.Children(0)\r\n'
        'If sess.Info.User = "" Then WScript.Echo "NOT_LOGGED_IN" : WScript.Quit 0\r\n'
        'WScript.Echo "OK:" & sess.Info.User\r\n'
    )
    tmp = ROOT / "_sap_probe.vbs"
    try:
        tmp.write_text(probe, encoding="utf-8")
        proc = subprocess.run(
            [cscript, "//nologo", str(tmp)],
            capture_output=True, text=True, timeout=20,
        )
        out = (proc.stdout or "").strip()
    except Exception:
        return False, "NO_SAPGUI"
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass
    return (out.startswith("OK:"), out if out.startswith("OK:") else (out or "NO_SAPGUI"))


# Fragments seen in SAP scripting COM errors -> a hint the user can act on.
_SAP_ERROR_HINTS = [
    ("could not be found", "SAP อาจไม่ได้อยู่ที่หน้าจอ ZPP0059 หรือมี popup ค้างอยู่ "
                           "— ปิด popup แล้วกลับไปหน้า ZPP0059 (SAP not on the ZPP0059 "
                           "screen, or a dialog is open.)"),
    ("enable scripting", "เปิด SAP GUI Scripting ก่อนใช้งาน (Enable SAP GUI Scripting.)"),
    ("scripting is disabled", "เปิด SAP GUI Scripting ก่อนใช้งาน (Enable SAP GUI Scripting.)"),
    ("is busy", "SAP กำลังทำงานอื่นอยู่ — รอสักครู่แล้วลองใหม่ (SAP is busy, try again.)"),
]


def _friendly_sap_error(raw: str) -> str:
    low = raw.lower()
    for frag, hint in _SAP_ERROR_HINTS:
        if frag in low:
            return f"{hint} [รายละเอียด: {raw[:200]}]"
    return f"SAP script error: {raw[:300] or 'unknown'}"


def _drive_sap_export(script_text: str = SAP_SCRIPT_0059,
                     transaction: str = START_TRANSACTION,
                     roll_dates: bool = DYNAMIC_DATES,
                     export_dir_win: str = SAP_EXPORT_DIR,
                     tmp_name: str = "_zpp0059_run.vbs") -> Path:
    """Drive SAP once and return the file it just exported.
    Raises RuntimeError on a (often transient) failure; TimeoutExpired on timeout."""
    cscript = shutil.which("cscript")
    script = _prepare_script(script_text, transaction, roll_dates, export_dir_win, tmp_name)
    # Start background handler BEFORE cscript so it's ready to catch the popup.
    sec_handler = _start_security_handler()
    started = time.time()
    try:
        proc = subprocess.run(
            [cscript, "//nologo", str(script)],
            capture_output=True, text=True, timeout=RUN_TIMEOUT,
        )
    finally:
        if sec_handler:
            try:
                sec_handler.terminate()
            except Exception:
                pass
        try:
            (ROOT / "_sec_handler.ps1").unlink(missing_ok=True)
        except Exception:
            pass
    if proc.returncode != 0:
        msg = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(_friendly_sap_error(msg))

    # Prefer the folder SAP was told to save into, so we never pick up a newer
    # file produced by a different transaction (e.g. a ZPP0059 stock export
    # landing in its own folder while we're running ZPP0022). Only fall back to
    # scanning every EXPORT_DIRS if nothing showed up there (SAP may have used a
    # default save location like C:\TEMP when the network drive was unwritable).
    target_dir = Path(export_dir_win)
    export = None
    for _ in range(20):
        if target_dir.is_dir():
            export = _newest_export(started, dirs=[target_dir])
        if not export:
            export = _newest_export(started)
        if export:
            break
        time.sleep(0.5)
    if not export:
        raise RuntimeError(
            "Ran the script but found no exported file. "
            "Add your SAP export folder to EXPORT_DIRS in serve_daily_follow.py.")
    return export


def run_zpp0059(date_from: str = "", date_to: str = "") -> tuple[dict, dict, str, dict]:
    """Drive SAP → upsert into DB → aggregate. Returns (mat, lot, filename, stats).
    date_from / date_to override the rolling window when supplied (format DD.MM.YYYY).
    """
    cscript = shutil.which("cscript")
    if not cscript:
        raise RuntimeError("cscript not found — must run on Windows with SAP GUI.")

    ready, code = _sap_session_ready()
    if not ready:
        raise SapNotReady(_SAP_NOT_READY_MSG.get(code, _SAP_NOT_READY_MSG["NOT_LOGGED_IN"]))

    # Retry transient GUI hiccups. Never retry a timeout (would double the wait),
    # and bail out early if the SAP session dropped between attempts.
    # Build the SAP script with the requested date window.
    # Explicit dates from the caller win; otherwise use the rolling offset.
    if date_from or date_to:
        d_low  = date_from or (date.today() + timedelta(days=DATE_FROM_OFFSET_DAYS)).strftime("%d.%m.%Y")
        d_high = date_to   or date.today().strftime("%d.%m.%Y")
        patched = re.sub(r'(S_WKDT-LOW"\)\.text\s*=\s*")[^"]*(")',  rf"\g<1>{d_low}\g<2>",  SAP_SCRIPT_0059)
        patched = re.sub(r'(S_WKDT-HIGH"\)\.text\s*=\s*")[^"]*(")', rf"\g<1>{d_high}\g<2>", patched)
        script_text = patched
        use_roll = False
    else:
        script_text = SAP_SCRIPT_0059
        use_roll = DYNAMIC_DATES

    export = None
    for attempt in range(1, RUN_ATTEMPTS + 1):
        try:
            export = _drive_sap_export(script_text=script_text, roll_dates=use_roll)
            break
        except subprocess.TimeoutExpired:
            raise
        except RuntimeError as exc:
            if attempt >= RUN_ATTEMPTS:
                raise
            print(f"[SAP] attempt {attempt} failed: {exc} — retrying…")
            time.sleep(3)
            ok, _ = _sap_session_ready()
            if not ok:
                raise SapNotReady(_SAP_NOT_READY_MSG["NOT_LOGGED_IN"])

    # SAP auto-opens the export in Excel — close that workbook again.
    _close_excel(export.name)

    # Keep a copy as ZPP0059.xlsx (backward compat for build_daily_follow.py).
    shutil.copyfile(export, PROGRESS_FILE)

    # Read the export and upsert into SQLite.
    headers, rows = _read_xlsx(export)
    conn = _get_db()
    new, updated = db_upsert(conn, headers, rows)
    print(f"[DB] {new} new rows, {updated} updated "
          f"(total export: {len(rows)})")

    # Aggregate from the full DB (not just this export).
    mat, lot = db_aggregate(conn)
    stats = db_stats(conn)
    stats["last_export"] = export.name
    stats["new_rows"] = new
    stats["updated_rows"] = updated
    conn.close()
    return mat, lot, export.name, stats


def run_zpp0022() -> tuple[str, dict]:
    """Drive SAP with the embedded ZPP0022 recording → save the order export →
    store raw rows in SQLite. The page then loads ZPP0022.xlsx and rebuilds the
    table from it (same pipeline as the manual 'Update Progress' import).
    Returns (filename, stats)."""
    cscript = shutil.which("cscript")
    if not cscript:
        raise RuntimeError("cscript not found — must run on Windows with SAP GUI.")

    ready, code = _sap_session_ready()
    if not ready:
        raise SapNotReady(_SAP_NOT_READY_MSG.get(code, _SAP_NOT_READY_MSG["NOT_LOGGED_IN"]))

    export = None
    for attempt in range(1, RUN_ATTEMPTS + 1):
        try:
            export = _drive_sap_export(
                script_text=SAP_SCRIPT_0022,
                transaction=START_TRANSACTION_0022,
                roll_dates=False,   # the 0022 script has no S_WKDT date window
                export_dir_win=SAP_EXPORT_DIR_0022,
                tmp_name="_zpp0022_run.vbs",
            )
            break
        except subprocess.TimeoutExpired:
            raise
        except RuntimeError as exc:
            if attempt >= RUN_ATTEMPTS:
                raise
            print(f"[SAP] ZPP0022 attempt {attempt} failed: {exc} — retrying…")
            time.sleep(3)
            ok, _ = _sap_session_ready()
            if not ok:
                raise SapNotReady(_SAP_NOT_READY_MSG["NOT_LOGGED_IN"])

    # SAP auto-opens the export in Excel — close that workbook again.
    _close_excel(export.name)

    # Keep the latest export as ZPP0022.xlsx so the page can fetch + parse it.
    shutil.copyfile(export, PROGRESS_0022_FILE)

    # Accumulate the raw rows in their own SQLite table (dedup, never wipe).
    headers, rows = _read_xlsx(export)
    conn = _get_db()
    new, updated = db_upsert(conn, headers, rows, table=ZPP0022_TABLE)
    print(f"[DB:0022] {new} new rows, {updated} updated "
          f"(total export: {len(rows)})")
    stats = db_stats(conn, table=ZPP0022_TABLE)
    stats["last_export"] = export.name
    stats["new_rows"] = new
    stats["updated_rows"] = updated
    conn.close()
    return export.name, stats


# ---------------------------------------------------------------------------
# HTTP server
# ---------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="application/json"):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        route = self.path.rstrip("/")
        if route == "/api/run-zpp0059":
            try:
                length = int(self.headers.get("Content-Length", 0))
                body = json.loads(self.rfile.read(length)) if length else {}
                mat, lot, fname, stats = run_zpp0059(
                    date_from=body.get("date_from", ""),
                    date_to=body.get("date_to", ""),
                )
                self._send(200, json.dumps(
                    {"ok": True, "mat": mat, "lot": lot, "file": fname, "stats": stats}
                ))
            except SapNotReady as exc:
                self._send(409, json.dumps(
                    {"ok": False, "error": str(exc), "code": "SAP_NOT_READY"}
                ))
            except subprocess.TimeoutExpired:
                self._send(504, json.dumps({"ok": False, "error": "SAP timed out."}))
            except Exception as exc:
                self._send(500, json.dumps({"ok": False, "error": str(exc)}))
        elif route == "/api/save-state":
            try:
                length = int(self.headers.get("Content-Length", 0))
                raw = self.rfile.read(length) if length else b"{}"
                # Validate it is JSON before persisting; reject garbage.
                json.loads(raw.decode("utf-8"))
                STATE_FILE.write_bytes(raw)
                self._send(200, json.dumps({"ok": True}))
            except Exception as exc:
                self._send(500, json.dumps({"ok": False, "error": str(exc)}))
        elif route == "/api/run-zpp0022":
            try:
                fname, stats = run_zpp0022()
                self._send(200, json.dumps(
                    {"ok": True, "file": PROGRESS_0022_FILE.name, "export": fname, "stats": stats}
                ))
            except SapNotReady as exc:
                self._send(409, json.dumps(
                    {"ok": False, "error": str(exc), "code": "SAP_NOT_READY"}
                ))
            except subprocess.TimeoutExpired:
                self._send(504, json.dumps({"ok": False, "error": "SAP timed out."}))
            except Exception as exc:
                self._send(500, json.dumps({"ok": False, "error": str(exc)}))
        else:
            self._send(404, json.dumps({"ok": False, "error": "not found"}))

    def do_GET(self):
        from urllib.parse import urlparse, parse_qs
        parsed = urlparse(self.path)
        path = parsed.path.lstrip("/")
        query = parse_qs(parsed.query)
        if path in ("", "index.html"):
            path = OUTPUT_FILE.name
        if path == OUTPUT_FILE.name:
            _rebuild_page_if_stale()
        if path == "api/state":
            try:
                if STATE_FILE.is_file():
                    self._send(200, STATE_FILE.read_bytes())
                else:
                    self._send(200, json.dumps({}))
            except Exception as exc:
                self._send(500, json.dumps({"error": str(exc)}))
            return
        if path == "api/checksheets":
            try:
                by_key = load_checksheets()
                self._send(200, json.dumps({"byKey": by_key}))
            except Exception as exc:
                self._send(500, json.dumps({"error": str(exc)}))
            return
        if path == "api/incoming":
            try:
                self._send(200, json.dumps({"byKey": load_incoming()}, ensure_ascii=False))
            except Exception as exc:
                self._send(500, json.dumps({"error": str(exc)}))
            return
        if path == "api/db-stats":
            try:
                conn = _get_db()
                db_init(conn, [])   # ensure table exists
                s = db_stats(conn)
                conn.close()
                self._send(200, json.dumps(s))
            except Exception as exc:
                self._send(500, json.dumps({"error": str(exc)}))
            return
        if path == "api/db-data":
            try:
                conn = _get_db()
                db_init(conn, [])
                q = (query.get("q", [""])[0] or "").strip()
                limit = min(int(query.get("limit", ["500"])[0]), 5000)
                offset = max(int(query.get("offset", ["0"])[0]), 0)
                filters = {}
                for k, vals in query.items():
                    m = re.fullmatch(r"f(\d+)", k)
                    if m and vals and vals[0].strip():
                        filters[int(m.group(1))] = vals[0].strip()
                data = db_query(conn, q, limit, offset, filters)
                conn.close()
                self._send(200, json.dumps(data))
            except Exception as exc:
                self._send(500, json.dumps({"error": str(exc)}))
            return
        if path == "api/actual-production":
            req_date = (query.get("date", [""])[0] or "").strip()
            if not req_date:
                return self._send(400, json.dumps({"success": False, "error": "missing date"}))
            try:
                by_shift = actual_production_by_date(req_date)
                self._send(200, json.dumps({"success": True, "byShift": by_shift}, ensure_ascii=False))
            except Exception as exc:
                self._send(200, json.dumps({"success": False, "error": str(exc)}))
            return
        if path == "api/ot-working-hours":
            # Proxy the OT app's per-shift working hours (server-side, no CORS).
            import urllib.request
            from urllib.error import URLError, HTTPError
            req_date = (query.get("date", [""])[0] or "").strip()
            if not req_date:
                return self._send(400, json.dumps(
                    {"success": False, "error": "missing date"}))
            url = f"{OT_APP_URL}/api/working-hours?work_date={req_date}"
            try:
                # Bypass any HTTP proxy for the local OT app.
                opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
                with opener.open(url, timeout=5) as resp:
                    body = resp.read()
                self._send(200, body)
            except (URLError, HTTPError, OSError) as exc:
                self._send(200, json.dumps({
                    "success": False,
                    "error": f"เชื่อมต่อโปรแกรม OT ไม่ได้ ({OT_APP_URL}). "
                             f"ตรวจสอบว่าโปรแกรม OT เปิดอยู่. [{exc}]"
                }))
            return
        target = (ROOT / path).resolve()
        if ROOT.resolve() not in target.parents and target != ROOT.resolve():
            return self._send(403, "forbidden", "text/plain")
        if not target.is_file():
            return self._send(404, "not found", "text/plain")
        ctype = "text/html; charset=utf-8" if target.suffix == ".html" else "application/octet-stream"
        self._send(200, target.read_bytes(), ctype)

    def log_message(self, *args):
        pass


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def _detect_lan_ip():
    """Best-effort: find this PC's LAN IP (the address other PCs would use).
    Opens a dummy UDP socket — no packets are actually sent."""
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return None
    finally:
        s.close()


# Files whose change should trigger a rebuild of daily_follow.html when the
# page is requested, so a browser refresh always reflects the latest data/code
# without restarting the server.
_BUILD_INPUTS = [
    bdf.TEMPLATE_FILE, bdf.EXPORT_FILE, bdf.ROUTING_FILE, bdf.PROGRESS_FILE,
    bdf.SAMPLE_FILE, bdf.MASTER_TS_FILE, bdf.MASTER_TS_DB_FILE,
    bdf.CT_XLSX_FILE, bdf.CT_DB_FILE,
    Path(__file__).resolve(), Path(bdf.__file__).resolve(),
]
_rebuild_lock = threading.Lock()


def _rebuild_page_if_stale():
    """Rebuild daily_follow.html if any source file or the build code is newer
    than it. Lets an edited MasterTS/CT/export show up on a plain refresh."""
    with _rebuild_lock:
        try:
            if not OUTPUT_FILE.exists():
                bdf.main()
                return
            out_mtime = OUTPUT_FILE.stat().st_mtime
            newest = 0.0
            for p in _BUILD_INPUTS:
                try:
                    if p.exists():
                        newest = max(newest, p.stat().st_mtime)
                except OSError:
                    pass
            if newest > out_mtime:
                print("Source changed — rebuilding daily_follow.html …")
                bdf.main()
        except Exception as exc:
            # Serve the existing page rather than 500 if a rebuild fails.
            print(f"[warn] rebuild skipped: {exc}")


def main():
    # Pre-populate DB from the existing ZPP0059.xlsx if the DB is new.
    if PROGRESS_FILE.exists():
        conn = _get_db()
        total = 0
        try:
            total = conn.execute("SELECT COUNT(*) FROM zpp0059_raw").fetchone()[0]
        except sqlite3.OperationalError:
            pass
        if total == 0:
            print("Pre-loading existing ZPP0059.xlsx into the database …")
            headers, rows = _read_xlsx(PROGRESS_FILE)
            inserted, _ = db_upsert(conn, headers, rows)
            print(f"  → {inserted} rows loaded.")
        conn.close()

    print("Building daily_follow.html …")
    bdf.main()

    url = f"http://{HOST}:{PORT}/{OUTPUT_FILE.name}"
    server = ThreadingHTTPServer((BIND_HOST, PORT), Handler)
    print(f"Daily Follow server (this PC): {url}")
    if BIND_HOST == "0.0.0.0":
        lan_ip = _detect_lan_ip()
        if lan_ip:
            print(f"Daily Follow server (LAN):     http://{lan_ip}:{PORT}/{OUTPUT_FILE.name}")
        print("(เครื่องอื่นเข้าผ่าน IP ของเครื่องนี้ได้ — เปิด Firewall ขาเข้า"
              f" TCP {PORT} ด้วย)")
    print(f"SQLite database:     {DB_FILE}")
    print("Press Ctrl+C to stop.")
    try:
        webbrowser.open(url)
    except Exception:
        pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
        server.shutdown()


if __name__ == "__main__":
    sys.exit(main())
