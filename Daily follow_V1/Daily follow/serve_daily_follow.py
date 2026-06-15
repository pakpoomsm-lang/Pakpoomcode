"""Local companion server for daily_follow.html.

Serves the dashboard and exposes one endpoint, POST /api/run-zpp0059, which:
  1. runs the SAP GUI script Script_0059.vbs (transaction ZPP0059) via cscript,
  2. locates the spreadsheet SAP just exported,
  3. copies it over ZPP0059.xlsx and re-aggregates progress,
  4. returns the fresh progress so the page updates itself.

This must run on the same Windows machine where SAP GUI is open and logged in.

Usage:
    python serve_daily_follow.py
then open the URL it prints (default http://127.0.0.1:8059/daily_follow.html).
Windows users can just double-click Start_Daily_Follow.bat.
"""

import json
import os
import re
import shutil
import subprocess
import sys
import time
import webbrowser
from datetime import date, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import build_daily_follow as bdf

# ---------------------------------------------------------------------------
# Configuration  (edit these to match your machine)
# ---------------------------------------------------------------------------
HOST = "127.0.0.1"
PORT = 8059
ROOT = bdf.ROOT
VBS_FILE = ROOT.parent / "Script_0059.vbs"   # uploaded one folder up
OUTPUT_FILE = bdf.OUTPUT_FILE
PROGRESS_FILE = bdf.PROGRESS_FILE             # ZPP0059.xlsx (overwritten each run)

# How long to wait for SAP to finish exporting before giving up (seconds).
RUN_TIMEOUT = 180

# Replace the dates inside the .vbs with a rolling window based on today.
# Set DYNAMIC_DATES = False to run the script exactly as recorded.
DYNAMIC_DATES = True
DATE_FROM_OFFSET_DAYS = -3   # S_WKDT-LOW  = today - 3 days
DATE_TO_OFFSET_DAYS = 0      # S_WKDT-HIGH = today

# Folders to watch for the file SAP exports. The newest spreadsheet created
# while the script runs is taken as the result. Add your SAP export folder
# here if the pull says it couldn't find the exported file.
HOME = Path(os.path.expanduser("~"))
EXPORT_DIRS = [
    ROOT,                                   # this "Daily follow" folder
    HOME / "Documents" / "SAP" / "SAP GUI",  # SAP GUI default working dir
    HOME / "Downloads",
    HOME / "Documents",
    HOME / "Desktop",
]
EXPORT_EXTS = (".xlsx", ".xls", ".xlsm", ".mhtml")


# ---------------------------------------------------------------------------
# SAP run helpers
# ---------------------------------------------------------------------------
def _read_vbs_text():
    raw = VBS_FILE.read_bytes()
    for enc in ("utf-16", "utf-16-le", "utf-8-sig", "utf-8", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", "ignore")


def _prepare_script():
    """Return the path of the .vbs to execute (a temp copy with rolling dates,
    or the original if DYNAMIC_DATES is off / substitution fails)."""
    if not DYNAMIC_DATES:
        return VBS_FILE
    try:
        text = _read_vbs_text()
        low = (date.today() + timedelta(days=DATE_FROM_OFFSET_DAYS)).strftime("%d.%m.%Y")
        high = (date.today() + timedelta(days=DATE_TO_OFFSET_DAYS)).strftime("%d.%m.%Y")
        text, n_low = re.subn(
            r'(S_WKDT-LOW"\)\.text\s*=\s*")[^"]*(")', rf"\g<1>{low}\g<2>", text)
        text, n_high = re.subn(
            r'(S_WKDT-HIGH"\)\.text\s*=\s*")[^"]*(")', rf"\g<1>{high}\g<2>", text)
        if not (n_low and n_high):
            return VBS_FILE  # pattern not found -> run original untouched
        tmp = ROOT / "_zpp0059_run.vbs"
        tmp.write_text(text, encoding="utf-8")
        return tmp
    except Exception:
        return VBS_FILE


def _newest_export(since):
    best, best_mtime = None, since - 2  # small slack for clock/fs jitter
    for d in EXPORT_DIRS:
        try:
            if not d.is_dir():
                continue
            for f in d.iterdir():
                if f.suffix.lower() not in EXPORT_EXTS:
                    continue
                if f.resolve() == PROGRESS_FILE.resolve():
                    continue  # don't pick the file we overwrite
                m = f.stat().st_mtime
                if m >= best_mtime:
                    best, best_mtime = f, m
        except OSError:
            continue
    return best


def run_zpp0059():
    """Drive SAP, refresh ZPP0059.xlsx, return (mat, lot, filename)."""
    if not VBS_FILE.exists():
        raise RuntimeError(f"Script not found: {VBS_FILE.name}")
    cscript = shutil.which("cscript")
    if not cscript:
        raise RuntimeError("cscript not found - this must run on Windows with SAP GUI.")

    script = _prepare_script()
    started = time.time()
    proc = subprocess.run(
        [cscript, "//nologo", str(script)],
        capture_output=True, text=True, timeout=RUN_TIMEOUT,
    )
    if proc.returncode != 0:
        msg = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(f"SAP script error: {msg[:300] or 'unknown'}")

    # SAP writes the file slightly after the script returns; poll briefly.
    export = None
    for _ in range(20):
        export = _newest_export(started)
        if export:
            break
        time.sleep(0.5)
    if not export:
        raise RuntimeError(
            "Ran the script but found no exported file. Add your SAP export "
            "folder to EXPORT_DIRS in serve_daily_follow.py.")

    shutil.copyfile(export, PROGRESS_FILE)
    mat, lot = bdf.load_progress()
    return mat, lot, export.name


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
        if self.path.rstrip("/") != "/api/run-zpp0059":
            return self._send(404, json.dumps({"ok": False, "error": "not found"}))
        try:
            mat, lot, fname = run_zpp0059()
            self._send(200, json.dumps({"ok": True, "mat": mat, "lot": lot, "file": fname}))
        except subprocess.TimeoutExpired:
            self._send(504, json.dumps({"ok": False, "error": "SAP took too long (timeout)."}))
        except Exception as exc:  # noqa: BLE001 - surface message to the page
            self._send(500, json.dumps({"ok": False, "error": str(exc)}))

    def do_GET(self):
        path = self.path.split("?", 1)[0].lstrip("/")
        if path in ("", "index.html"):
            path = OUTPUT_FILE.name
        target = (ROOT / path).resolve()
        if ROOT.resolve() not in target.parents and target != ROOT.resolve():
            return self._send(403, "forbidden", "text/plain")
        if not target.is_file():
            return self._send(404, "not found", "text/plain")
        ctype = "text/html; charset=utf-8" if target.suffix == ".html" else "application/octet-stream"
        self._send(200, target.read_bytes(), ctype)

    def log_message(self, *args):  # quieter console
        pass


def main():
    if not OUTPUT_FILE.exists():
        print("Building daily_follow.html ...")
        bdf.main()
    url = f"http://{HOST}:{PORT}/{OUTPUT_FILE.name}"
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"Daily Follow server running at {url}")
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
