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
HOST = "127.0.0.1"
PORT = 8059
ROOT = bdf.ROOT
VBS_FILE = ROOT.parent / "Script_0059.vbs"
OUTPUT_FILE = bdf.OUTPUT_FILE
PROGRESS_FILE = bdf.PROGRESS_FILE          # ZPP0059.xlsx — still kept for fallback

# SQLite database path (same folder as the scripts).
DB_FILE = ROOT / "zpp0059.db"

# How long to wait for SAP to finish exporting (seconds).
RUN_TIMEOUT = 180

# Rolling date window injected into the VBS before each run.
DYNAMIC_DATES = True
DATE_FROM_OFFSET_DAYS = -3   # S_WKDT-LOW  = today - 3 days
DATE_TO_OFFSET_DAYS   = 0    # S_WKDT-HIGH = today

# Navigate to ZPP0059 before touching its selection screen fields.
# Set to "" to run the recorded VBS exactly as-is.
START_TRANSACTION = "ZPP0059"

# Folders watched for the file SAP exports (searched in order, newest file wins).
HOME = Path(os.path.expanduser("~"))
EXPORT_DIRS = [
    Path(r"J:\7.541_HEI\Database follow\ZPP0059"),  # shared drive (primary)
    Path(r"C:\TEMP"),                               # SAP default save dir
    ROOT,
    HOME / "Documents" / "SAP" / "SAP GUI",
    HOME / "Downloads",
    HOME / "Documents",
    HOME / "Desktop",
]
# Target directory for SAP to save the export file into.
# Must match one of the EXPORT_DIRS entries above.
SAP_EXPORT_DIR = str(Path(r"J:\7.541_HEI\Database follow\ZPP0059"))

# Business key columns — a row is considered duplicate when ALL of these match.
# If any column is missing from the export, falls back to SHA-256 of the whole row.
UNIQUE_KEY_COLS = ["Order", "Activity", "Short Time Stamp", "Tag ID"]

# Columns used to aggregate progress (must match bdf.load_progress logic).
PROGRESS_COLS = [
    "Production Line", "Production Month", "Sequence",
    "Material", "Assembly Order", "Operation Short Text", "Posted Quantity",
]


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


def db_init(conn: sqlite3.Connection, headers: list[str]) -> list[str]:
    """Ensure the table exists and has all required columns. Returns safe col names."""
    safe = [_col_name(h) for h in headers]
    cols_ddl = ", ".join(f'"{c}" TEXT' for c in safe)
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS zpp0059_raw (
            _row_hash TEXT PRIMARY KEY,
            _inserted_at TEXT DEFAULT (datetime('now','localtime')),
            {cols_ddl}
        )
    """)
    # Add any columns that appeared in a newer export but not in the original table.
    existing = {r[1] for r in conn.execute("PRAGMA table_info(zpp0059_raw)")}
    for c in safe:
        if c not in existing:
            conn.execute(f'ALTER TABLE zpp0059_raw ADD COLUMN "{c}" TEXT')
    conn.commit()
    return safe


def db_upsert(conn: sqlite3.Connection, headers: list[str], rows) -> tuple[int, int]:
    """Insert rows that don't already exist. Returns (inserted, skipped)."""
    safe_headers = db_init(conn, headers)
    h_to_safe = {h: s for h, s in zip(headers, safe_headers)}

    # Determine which columns to use for the business key.
    key_cols = [c for c in UNIQUE_KEY_COLS if c in headers]
    use_hash = len(key_cols) < 2  # fallback: SHA-256 of entire row

    inserted = skipped = 0
    placeholders = ", ".join("?" for _ in safe_headers)
    col_list = ", ".join(f'"{c}"' for c in safe_headers)
    sql = (f'INSERT OR IGNORE INTO zpp0059_raw (_row_hash, {col_list}) '
           f'VALUES (?, {placeholders})')

    batch = []
    for row in rows:
        values = [str(v).strip() if v is not None else "" for v in row]
        if use_hash:
            key = _row_hash(values)
        else:
            key_vals = [values[headers.index(c)] for c in key_cols]
            key = hashlib.sha256("|".join(key_vals).encode()).hexdigest()
        batch.append((key, *values))
        if len(batch) >= 500:
            cur = conn.executemany(sql, batch)
            inserted += cur.rowcount
            skipped += len(batch) - cur.rowcount
            batch.clear()

    if batch:
        cur = conn.executemany(sql, batch)
        inserted += cur.rowcount
        skipped += len(batch) - cur.rowcount
    conn.commit()
    return inserted, skipped


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

    def month_display(v):
        try:
            from datetime import datetime
            for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%m/%d/%Y", "%Y%m%d"):
                try:
                    return datetime.strptime(str(v).strip(), fmt).strftime("%-m.%Y")
                except ValueError:
                    pass
        except Exception:
            pass
        return str(v).strip()

    def seq_key(v):
        try:
            return str(int(float(str(v).strip())))
        except Exception:
            return str(v).strip()

    by_mat: dict = defaultdict(lambda: defaultdict(float))
    by_lot: dict = defaultdict(float)

    sql = (f'SELECT "{c("Production Line")}", "{c("Production Month")}", '
           f'"{c("Sequence")}", "{c("Material")}", "{c("Assembly Order")}", '
           f'"{c("Operation Short Text")}", "{c("Posted Quantity")}" '
           f'FROM zpp0059_raw WHERE "{c("Production Line")}" != ""')

    for row in conn.execute(sql):
        line, month_raw, seq_raw, mat, ao, op, qty_raw = row
        if not line:
            continue
        try:
            qty = float(qty_raw) if qty_raw else 0.0
        except ValueError:
            qty = 0.0
        head = f"{line}|{month_display(month_raw)}|{seq_key(seq_raw)}"
        if op == hp_op:
            by_lot[f"{head}|{ao}"] += qty
        elif op in op_field:
            by_mat[f"{head}|{mat}"][op_field[op]] += qty

    def rnd(v):
        return int(v) if float(v).is_integer() else round(v, 3)

    mat_out = {k: {f: rnd(q) for f, q in fields.items()} for k, fields in by_mat.items()}
    lot_out = {k: rnd(q) for k, q in by_lot.items()}
    return mat_out, lot_out


def db_stats(conn: sqlite3.Connection) -> dict:
    total = conn.execute("SELECT COUNT(*) FROM zpp0059_raw").fetchone()[0]
    newest = conn.execute(
        "SELECT MAX(_inserted_at) FROM zpp0059_raw"
    ).fetchone()[0]
    return {"total_rows": total, "newest_inserted_at": newest}


# ---------------------------------------------------------------------------
# SAP + export helpers
# ---------------------------------------------------------------------------
def _read_vbs_text() -> str:
    raw = VBS_FILE.read_bytes()
    for enc in ("utf-16", "utf-16-le", "utf-8-sig", "utf-8", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", "ignore")


def _prepare_script() -> Path:
    try:
        text = _read_vbs_text()

        # 1) Roll the work-date window.
        if DYNAMIC_DATES:
            low  = (date.today() + timedelta(days=DATE_FROM_OFFSET_DAYS)).strftime("%d.%m.%Y")
            high = (date.today() + timedelta(days=DATE_TO_OFFSET_DAYS)).strftime("%d.%m.%Y")
            text = re.sub(r'(S_WKDT-LOW"\)\.text\s*=\s*")[^"]*(")',  rf"\g<1>{low}\g<2>",  text)
            text = re.sub(r'(S_WKDT-HIGH"\)\.text\s*=\s*")[^"]*(")', rf"\g<1>{high}\g<2>", text)

        # 2) Navigate to the transaction before touching its selection fields.
        if START_TRANSACTION and "/tbar[0]/okcd" not in text:
            nav = (f'session.findById("wnd[0]/tbar[0]/okcd").text = "/n{START_TRANSACTION}"\r\n'
                   f'session.findById("wnd[0]").sendVKey 0\r\n')
            m = re.search(r'^\s*session\.findById\("wnd\[0\]"\)\.resizeWorkingPane', text, re.M)
            if not m:
                m = re.search(r'^\s*session\.findById\("wnd\[0\]', text, re.M)
            if m:
                text = text[:m.start()] + nav + text[m.start():]

        # 3) Set export directory + filename in the SAP "Save File" dialog.
        #    SAP uses wnd[1]/usr/ctxtDY_PATH and ctxtDY_FILENAME for the local file dialog.
        #    We inject these lines just before the final wnd[1] btn[0] (Generate) press.
        export_dir  = SAP_EXPORT_DIR.replace("\\", "\\\\")
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

        tmp = ROOT / "_zpp0059_run.vbs"
        tmp.write_text(text, encoding="utf-8")
        return tmp
    except Exception:
        return VBS_FILE


def _start_security_handler() -> subprocess.Popen | None:
    """Launch a background PowerShell that watches for the SAP GUI Security
    popup, ticks 'Remember My Decision', then clicks Allow — so the user
    never has to touch it manually.  Returns the Popen handle (or None on
    non-Windows / if powershell not found)."""
    ps = shutil.which("powershell") or shutil.which("pwsh")
    if not ps:
        return None
    script = r"""
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$AutomationElement  = [System.Windows.Automation.AutomationElement]
$ControlType        = [System.Windows.Automation.ControlType]
$Condition          = [System.Windows.Automation.PropertyCondition]
$InvokePattern      = [System.Windows.Automation.InvokePattern]::Pattern
$TogglePattern      = [System.Windows.Automation.TogglePattern]::Pattern

# Keep handling popups for the whole run; SAP raises several in sequence
# (scripting attach, create file, modify directory, ...). The parent process
# kills this handler once cscript returns, so we just loop until then.
$deadline = (Get-Date).AddSeconds(150)
while ((Get-Date) -lt $deadline) {
    $root = $AutomationElement::RootElement
    $wins = $root.FindAll(
        [System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($w in $wins) {
        $title = $w.Current.Name
        if ($title -notmatch "SAP GUI Security") { continue }
        # Tick every "Remember My Decision" checkbox in the dialog.
        $chks = $w.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            (New-Object $Condition(
                $AutomationElement::ControlTypeProperty, $ControlType::CheckBox)))
        foreach ($chk in $chks) {
            try {
                $tp = $chk.GetCurrentPattern($TogglePattern)
                if ($tp.Current.ToggleState -ne [System.Windows.Automation.ToggleState]::On) {
                    $tp.Toggle()
                }
            } catch {}
        }
        Start-Sleep -Milliseconds 150
        # Click "Allow".
        $btns = $w.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            (New-Object $Condition(
                $AutomationElement::ControlTypeProperty, $ControlType::Button)))
        foreach ($btn in $btns) {
            if ($btn.Current.Name -match "Allow|Zulassen|อนุญาต") {
                try { $btn.GetCurrentPattern($InvokePattern).Invoke() } catch {}
                break
            }
        }
        Start-Sleep -Milliseconds 400
    }
    Start-Sleep -Milliseconds 250
}
"""
    try:
        return subprocess.Popen(
            [ps, "-NoProfile", "-NonInteractive", "-Command", script],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
    except Exception:
        return None


def _newest_export(since: float) -> Path | None:
    best, best_mtime = None, since - 2
    for d in EXPORT_DIRS:
        try:
            if not d.is_dir():
                continue
            for f in d.iterdir():
                if f.suffix.lower() not in EXPORT_EXTS:
                    continue
                if f.resolve() == PROGRESS_FILE.resolve():
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


def run_zpp0059() -> tuple[dict, dict, str, dict]:
    """Drive SAP → upsert into DB → aggregate. Returns (mat, lot, filename, stats)."""
    if not VBS_FILE.exists():
        raise RuntimeError(f"Script not found: {VBS_FILE.name}")
    cscript = shutil.which("cscript")
    if not cscript:
        raise RuntimeError("cscript not found — must run on Windows with SAP GUI.")

    script = _prepare_script()
    # Start background handler BEFORE cscript so it's ready to catch the popup.
    sec_handler = _start_security_handler()
    started = time.time()
    proc = subprocess.run(
        [cscript, "//nologo", str(script)],
        capture_output=True, text=True, timeout=RUN_TIMEOUT,
    )
    if sec_handler:
        try:
            sec_handler.terminate()
        except Exception:
            pass
    if proc.returncode != 0:
        msg = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(f"SAP script error: {msg[:300] or 'unknown'}")

    export = None
    for _ in range(20):
        export = _newest_export(started)
        if export:
            break
        time.sleep(0.5)
    if not export:
        raise RuntimeError(
            "Ran the script but found no exported file. "
            "Add your SAP export folder to EXPORT_DIRS in serve_daily_follow.py.")

    # Keep a copy as ZPP0059.xlsx (backward compat for build_daily_follow.py).
    shutil.copyfile(export, PROGRESS_FILE)

    # Read the export and upsert into SQLite.
    headers, rows = _read_xlsx(export)
    conn = _get_db()
    inserted, skipped = db_upsert(conn, headers, rows)
    print(f"[DB] {inserted} new rows inserted, {skipped} duplicates skipped "
          f"(total export: {len(rows)})")

    # Aggregate from the full DB (not just this export).
    mat, lot = db_aggregate(conn)
    stats = db_stats(conn)
    stats["last_export"] = export.name
    stats["new_rows"] = inserted
    stats["skipped_rows"] = skipped
    conn.close()
    return mat, lot, export.name, stats


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
        if self.path.rstrip("/") == "/api/run-zpp0059":
            try:
                mat, lot, fname, stats = run_zpp0059()
                self._send(200, json.dumps(
                    {"ok": True, "mat": mat, "lot": lot, "file": fname, "stats": stats}
                ))
            except subprocess.TimeoutExpired:
                self._send(504, json.dumps({"ok": False, "error": "SAP timed out."}))
            except Exception as exc:
                self._send(500, json.dumps({"ok": False, "error": str(exc)}))
        else:
            self._send(404, json.dumps({"ok": False, "error": "not found"}))

    def do_GET(self):
        path = self.path.split("?", 1)[0].lstrip("/")
        if path in ("", "index.html"):
            path = OUTPUT_FILE.name
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

    if not OUTPUT_FILE.exists():
        print("Building daily_follow.html …")
        bdf.main()

    url = f"http://{HOST}:{PORT}/{OUTPUT_FILE.name}"
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"Daily Follow server: {url}")
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
