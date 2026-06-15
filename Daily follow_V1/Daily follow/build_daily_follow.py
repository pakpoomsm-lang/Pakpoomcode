import json
from collections import Counter, defaultdict
from datetime import date, datetime, time
from html import escape
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parent
EXPORT_FILE = ROOT / "EXPORT_20260601070215.xlsx"
SAMPLE_FILE = ROOT / "HEI-18-05-PP-Original-SAP-1 (1).xlsm"
ROUTING_FILE = ROOT / "Routing.xlsx"
PROGRESS_FILE = ROOT / "ZPP0059.xlsx"
OUTPUT_FILE = ROOT / "daily_follow.html"

# Map ZPP0059 "Operation Short Text" -> progress column (material-level)
OP_TO_FIELD = {"Insert": "fp", "Brazing": "auto", "Cutting": "cutting"}
# H/P bender is recorded on sub-component materials, so it is summed per assembly lot.
HP_OP = "H/P bender"


def norm(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def number(value):
    if value in (None, ""):
        return ""
    try:
        f = float(str(value).strip())
    except ValueError:
        return value
    return int(f) if f.is_integer() else round(f, 3)


def excel_date(value):
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    raw = str(value).strip()
    if " " in raw:
        raw = raw.split(" ", 1)[0]
    try:
        return datetime.fromisoformat(raw).date().isoformat()
    except ValueError:
        return raw


def excel_time(value):
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%H:%M")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    raw = str(value).strip()
    if " " in raw:
        raw = raw.split(" ", 1)[-1]
    parts = raw.split(":")
    if len(parts) >= 2:
        return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}"
    return raw


def parse_dt(d_value, t_value):
    d = excel_date(d_value)
    t = excel_time(t_value) or "00:00"
    if not d:
        return None
    try:
        return datetime.fromisoformat(f"{d}T{t}")
    except ValueError:
        return None


def month_display(value):
    raw = text(value)
    if "." in raw:
        left, right = raw.split(".", 1)
        try:
            return f"{int(left)}.{right}"
        except ValueError:
            return raw
    return raw


def load_routing():
    wb = openpyxl.load_workbook(ROUTING_FILE, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    headers = [text(v) for v in next(rows)]
    idx = {h: i for i, h in enumerate(headers)}
    routing = {}
    for row in rows:
        material = text(row[idx["Material"]])
        if not material or material in routing:
            continue
        routing[material] = {
            "description": text(row[idx["Description"]]),
            "operation": text(row[idx["Operation Text"]]),
            "attribute1": text(row[idx["Attribute1"]]),
            "attribute2": text(row[idx["Attribute2"]]),
            "unit": number(row[idx["Standard lot"]]),
            "ts": number(row[idx["OPR Time Standard"]]),
        }
    wb.close()
    return routing


def seq_key(value):
    n = number(value)
    return str(n) if n != "" else ""


def final_qty(attribute, matp):
    """Finished quantity = qty of the last operation, decided by Attribute suffix.

    -BZ -> ends at Brazing (matp["auto"]); -CT -> ends at Cutting (matp["cutting"]).
    Anything else -> blank.
    """
    a = (attribute or "").strip().upper()
    if a.endswith("-BZ"):
        return matp.get("auto", "")
    if a.endswith("-CT"):
        return matp.get("cutting", "")
    return ""


def clean_qty(value):
    if not isinstance(value, (int, float)):
        return 0
    return int(value) if float(value).is_integer() else round(float(value), 2)


def load_progress():
    """Aggregate real per-process progress from ZPP0059.

    Returns (by_mat, by_lot):
      by_mat["line|month|seq|material"] = {"fp": q, "auto": q, "cutting": q}
      by_lot["line|month|seq|assyOrder"] = q   (H/P bender, summed per lot)
    """
    by_mat = defaultdict(lambda: defaultdict(float))
    by_lot = defaultdict(float)
    if not PROGRESS_FILE.exists():
        return {}, {}
    wb = openpyxl.load_workbook(PROGRESS_FILE, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    headers = [text(v) for v in next(rows)]
    idx = {h: i for i, h in enumerate(headers)}
    c_line = idx["Production Line"]
    c_month = idx["Production Month"]
    c_seq = idx["Sequence"]
    c_mat = idx["Material"]
    c_ao = idx["Assembly Order"]
    c_op = idx["Operation Short Text"]
    c_qty = idx["Posted Quantity"]
    for row in rows:
        line = text(row[c_line])
        if not line:
            continue
        key_head = f"{line}|{month_display(row[c_month])}|{seq_key(row[c_seq])}"
        op = text(row[c_op])
        qty = row[c_qty] if isinstance(row[c_qty], (int, float)) else 0
        if op == HP_OP:
            by_lot[f"{key_head}|{text(row[c_ao])}"] += qty
        elif op in OP_TO_FIELD:
            by_mat[f"{key_head}|{text(row[c_mat])}"][OP_TO_FIELD[op]] += qty
    wb.close()
    mat_out = {
        k: {f: clean_qty(q) for f, q in fields.items()}
        for k, fields in by_mat.items()
    }
    lot_out = {k: clean_qty(q) for k, q in by_lot.items()}
    return mat_out, lot_out


def most_common(counter, default=""):
    if not counter:
        return default
    return counter.most_common(1)[0][0]


def load_speed_maps():
    by_line = defaultdict(Counter)
    by_model = defaultdict(Counter)
    wb = openpyxl.load_workbook(SAMPLE_FILE, read_only=True, data_only=True, keep_vba=False)
    for sheet in wb.sheetnames:
        if not sheet.startswith("Line "):
            continue
        ws = wb[sheet]
        for row in ws.iter_rows(min_row=7, values_only=True):
            line = text(row[0]) if len(row) > 0 else ""
            code = text(row[3]) if len(row) > 3 else ""
            model = text(row[4]) if len(row) > 4 else ""
            speed = row[11] if len(row) > 11 else None
            if not line or not isinstance(speed, (int, float)):
                continue
            speed = int(speed) if float(speed).is_integer() else round(float(speed), 2)
            by_line[line][speed] += 1
            if code and model:
                by_model[(line, code, model)][speed] += 1
    wb.close()
    return (
        {line: most_common(counter) for line, counter in by_line.items()},
        {key: most_common(counter) for key, counter in by_model.items()},
    )


def build_rows():
    routing = load_routing()
    speed_by_line, speed_by_model = load_speed_maps()
    prog_mat, prog_lot = load_progress()
    wb = openpyxl.load_workbook(EXPORT_FILE, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    next(rows)
    out = []
    for row in rows:
        line = text(row[0])
        if not line:
            continue

        item = text(row[6])
        route = routing.get(item, {})
        prod_dt = parse_dt(row[9], row[10])
        finish_dt = parse_dt(row[51] or row[52], row[53] or row[45])
        lead = ""
        if prod_dt and finish_dt:
            lead = round((prod_dt - finish_dt).total_seconds() / 3600, 2)

        assy_qty = number(row[12])
        status = text(row[8])
        complete = status.upper() == "DLV" or text(row[40]).lower() == "complete"
        progress_value = assy_qty if complete else ""
        code = text(row[3])
        model = text(row[4])
        month = month_display(row[2])
        assy_order_no = text(row[13])
        head = f"{line}|{month}|{seq_key(row[1])}"
        matp = prog_mat.get(f"{head}|{item}", {})
        hp_val = prog_lot.get(f"{head}|{assy_order_no}", "")
        fin = final_qty(route.get("attribute2"), matp)
        remark = " ".join(
            part for part in [text(row[23]), text(row[24]), text(row[50])] if part
        )

        out.append(
            {
                "line": line,
                "seq": number(row[1]),
                "month": month,
                "code": code,
                "model": model,
                "finished": "",
                "orderQty": number(row[5]),
                "item": item,
                "description": route.get("description") or text(row[7]),
                "attribute": route.get("attribute2") or "",
                "attribute1": route.get("attribute1") or "",
                "productionOrder": text(row[11]),
                "speed": speed_by_model.get((line, code, model), speed_by_line.get(line, "")),
                "prodDate": excel_date(row[9]),
                "prodTime": excel_time(row[10]),
                "prodSort": prod_dt.isoformat() if prod_dt else "",
                "assyOrder": assy_qty,
                "assyOrderNo": assy_order_no,
                "status": status,
                "remark": remark,
                "hp": hp_val,
                "fp": matp.get("fp", ""),
                "exp": progress_value,
                "auto": matp.get("auto", ""),
                "cutting": matp.get("cutting", ""),
                "fg": fin,
                "unit": route.get("unit") or "",
                "stockFg": fin,
                "subcooler": fin,
                "lead": lead,
                "leadRemark": text(row[41]) or text(row[44]),
                "ts": route.get("ts") or "",
                "operation": route.get("operation") or "",
                "sourceReady": {
                    "metal": norm(row[27]),
                    "paint2": norm(row[28]),
                    "pipe1": norm(row[29]),
                    "pipe2": norm(row[30]),
                    "hex1out": norm(row[31]),
                    "hex1in": norm(row[32]),
                    "hex2": norm(row[33]),
                    "dirOrder": norm(row[39]),
                    "dirGr": norm(row[40]),
                },
            }
        )
    wb.close()
    out.sort(
        key=lambda r: (
            r["line"],
            r["prodSort"] or "9999",
            float(r["seq"]) if isinstance(r["seq"], (int, float)) else 999999,
            r["assyOrderNo"],
            r["productionOrder"],
            r["item"],
        )
    )
    return out


def render_html(rows):
    default_line = "ALL"
    default_date = ""
    routing = load_routing()
    speed_by_line, speed_by_model = load_speed_maps()
    prog_mat, prog_lot = load_progress()
    payload = json.dumps(
        {
            "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "sourceFile": EXPORT_FILE.name,
            "sampleFile": SAMPLE_FILE.name,
            "routingFile": ROUTING_FILE.name,
            "progressFile": PROGRESS_FILE.name,
            "defaultLine": default_line,
            "defaultDate": default_date,
            "routing": routing,
            "speedByLine": speed_by_line,
            "speedByModel": {
                "|".join(key): value for key, value in speed_by_model.items()
            },
            "progressMat": prog_mat,
            "progressLot": prog_lot,
            "rows": rows,
        },
        ensure_ascii=False,
        separators=(",", ":"),
    )
    return f"""<!doctype html>
<html lang="th">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>HEAT EXCHANGER SHOP DAILY FOLLOW</title>
  <style>
    :root {{
      --bg: #eef1f7;
      --surface: #ffffff;
      --surface-2: #f8fafc;
      --border: #dce3ee;
      --grid: #e6ebf3;
      --ink: #1e293b;
      --ink-soft: #475569;
      --muted: #94a3b8;
      --primary: #2f56c4;
      --primary-dark: #213f96;
      --head: #3a548f;
      --head-cyan: #0e7c8b;
      --head-green: #1f9254;
      --head-amber: #c08109;
      --head-purple: #6d3bd1;
      --head-pink: #d6457f;
      --head-edit: #3f5bd9;
      --edit-tint: #fff9e8;
      --late: #e23b3b;
      --ready: #1f9254;
      --missing: #d99b00;
      --shadow-sm: 0 1px 2px rgba(16,24,40,.06);
      --shadow: 0 1px 3px rgba(16,24,40,.08), 0 12px 28px rgba(16,24,40,.07);
      --yellow: #ffe600;
      --green: #00b050;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: "Segoe UI", "Inter", "Noto Sans Thai", Tahoma, Arial, sans-serif;
      font-size: 12.5px;
      -webkit-font-smoothing: antialiased;
    }}
    .app {{
      height: 100vh;
      display: grid;
      grid-template-rows: auto 1fr;
    }}
    .top {{
      background: var(--surface);
      border-bottom: 1px solid var(--border);
      box-shadow: var(--shadow-sm);
      z-index: 5;
    }}
    .banner {{
      height: 60px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 20px;
      background: linear-gradient(180deg, #ffffff 0%, #fbfcfe 100%);
    }}
    .brand {{
      display: flex;
      align-items: center;
      gap: 13px;
      min-width: 280px;
    }}
    .logo {{
      display: inline-grid;
      place-items: center;
      width: 42px;
      height: 28px;
      border-radius: 8px;
      background: linear-gradient(135deg, #f4495a, #d62f40);
      color: #fff;
      font-size: 12px;
      font-weight: 800;
      letter-spacing: .5px;
      box-shadow: 0 4px 12px rgba(214, 47, 64, .3);
    }}
    .brand-title {{
      color: #1f2937;
      font-size: 14px;
      font-weight: 800;
      letter-spacing: .2px;
      line-height: 1.1;
    }}
    .brand-subtitle {{
      margin-top: 4px;
      color: var(--muted);
      font-size: 9px;
      font-weight: 700;
      letter-spacing: 2px;
      line-height: 1;
    }}
    .connection {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      color: #138a55;
      font-size: 10px;
      font-weight: 700;
      border: 1px solid #b7ecd0;
      border-radius: 999px;
      padding: 6px 13px;
      background: #f1fdf6;
      text-transform: uppercase;
      letter-spacing: .6px;
    }}
    .connection::before {{
      content: "";
      width: 7px;
      height: 7px;
      border-radius: 999px;
      background: #21c177;
      box-shadow: 0 0 0 3px rgba(33, 193, 119, .18);
      animation: pulse 2s ease-in-out infinite;
    }}
    @keyframes pulse {{
      0%, 100% {{ box-shadow: 0 0 0 3px rgba(33, 193, 119, .18); }}
      50% {{ box-shadow: 0 0 0 5px rgba(33, 193, 119, .06); }}
    }}
    .dash-btn {{
      display: inline-flex;
      align-items: center;
      gap: 7px;
      height: 34px;
      padding: 0 17px;
      border: none;
      border-radius: 999px;
      background: var(--primary);
      color: #fff;
      font: 800 11px/1 inherit;
      letter-spacing: .5px;
      text-transform: uppercase;
      cursor: pointer;
      box-shadow: 0 2px 8px rgba(47, 86, 196, .3);
    }}
    .dash-btn:hover {{ background: var(--primary-dark); box-shadow: 0 4px 12px rgba(47, 86, 196, .35); }}
    .dash-overlay {{
      position: fixed;
      inset: 0;
      z-index: 100;
      background: rgba(15, 23, 42, .45);
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 30px;
    }}
    .dash-overlay[hidden] {{ display: none; }}
    .dash-panel {{
      width: min(900px, 96vw);
      max-height: 90vh;
      display: flex;
      flex-direction: column;
      background: #fff;
      border-radius: 16px;
      box-shadow: 0 24px 64px rgba(0, 0, 0, .32);
      overflow: hidden;
    }}
    .dash-head {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 16px 22px;
      border-bottom: 1px solid var(--border);
      background: linear-gradient(180deg, #fff, #fbfcfe);
    }}
    .dash-title {{ font-size: 16px; font-weight: 800; color: var(--ink); }}
    .dash-close {{
      width: 32px;
      height: 32px;
      border: 1px solid var(--border);
      border-radius: 8px;
      background: #fff;
      cursor: pointer;
      font-size: 14px;
      color: var(--ink-soft);
    }}
    .dash-close:hover {{ background: #fff1f2; border-color: #e8939b; color: #d62f40; }}
    .db-panel {{ width: min(1280px, 98vw); }}
    .db-toolbar {{
      display: flex; align-items: center; gap: 10px;
      padding: 12px 22px; border-bottom: 1px solid var(--border); flex-wrap: wrap;
    }}
    .db-toolbar input[type="search"] {{
      flex: 1 1 240px; height: 32px; padding: 0 12px;
      border: 1px solid var(--border); border-radius: 8px; font: inherit;
    }}
    .db-meta {{ font-size: 12px; color: var(--ink-soft); }}
    .db-refresh {{
      height: 32px; padding: 0 14px; border: 1px solid transparent; border-radius: 8px;
      background: #0a8043; color: #fff; font: 700 11.5px/1 inherit; cursor: pointer;
    }}
    .db-refresh:hover {{ background: #0b6e3a; }}
    .db-table-wrap {{ overflow: auto; flex: 1; padding: 0 0 8px; }}
    .db-table {{ border-collapse: collapse; font-size: 12px; width: max-content; min-width: 100%; }}
    .db-table th, .db-table td {{
      border: 1px solid var(--border); padding: 5px 9px; white-space: nowrap; text-align: left;
    }}
    .db-table th {{
      position: sticky; top: 0; background: #f1f5fb; font-weight: 700;
      color: var(--ink); z-index: 1;
    }}
    .db-table tbody tr:nth-child(even) {{ background: #fafbfd; }}
    .db-table tbody tr:hover {{ background: #eef4ff; }}
    .db-empty {{ padding: 40px; text-align: center; color: var(--ink-soft); }}
    .dash-body {{ padding: 20px 22px; overflow: auto; }}
    .dash-stats {{ display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 18px; }}
    .dash-stat {{
      flex: 1 1 0;
      min-width: 130px;
      padding: 12px 16px;
      border: 1px solid var(--border);
      border-left: 3px solid var(--primary);
      border-radius: 10px;
      background: var(--surface-2);
    }}
    .dash-stat .l {{ font-size: 9.5px; font-weight: 700; letter-spacing: .8px; text-transform: uppercase; color: var(--muted); }}
    .dash-stat .v {{ font-size: 22px; font-weight: 800; color: var(--ink); font-variant-numeric: tabular-nums; }}
    .dash-stat.assy {{ border-left-color: var(--primary); background: #eef2fc; }}
    .dash-stat.assy .v {{ color: var(--primary-dark); }}
    .dash-stat.sc {{ border-left-color: #14b8a6; background: #e7f5f5; }}
    .dash-stat.sc .v {{ color: #0c6b78; }}
    .dash-legend {{
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 12px;
      padding: 10px 14px;
      margin-bottom: 16px;
      border: 1px solid var(--border);
      border-radius: 10px;
      background: var(--surface-2);
    }}
    .dash-leg-title {{ font-size: 10px; font-weight: 800; letter-spacing: .5px; text-transform: uppercase; color: var(--muted); }}
    .dash-leg-item {{ display: inline-flex; align-items: center; gap: 5px; font-size: 11px; font-weight: 700; color: var(--ink-soft); }}
    .dash-leg-sw {{ width: 14px; height: 14px; border-radius: 4px; display: inline-block; }}
    .dline {{
      display: grid;
      grid-template-columns: 60px 1fr 1fr;
      gap: 12px;
      align-items: stretch;
      padding: 10px;
      margin-bottom: 8px;
      border: 1px solid var(--border);
      border-radius: 12px;
      background: #fff;
      cursor: pointer;
    }}
    .dline:hover {{ border-color: var(--primary); box-shadow: var(--shadow-sm); }}
    .dline-name {{
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
      font-size: 24px;
      font-weight: 800;
      color: var(--primary);
      line-height: 1.1;
    }}
    .dline-name small {{ font-size: 8.5px; font-weight: 700; color: var(--muted); margin-top: 4px; text-align: center; }}
    .dmetric {{ padding: 10px 13px; border-radius: 10px; }}
    .dmetric.assy {{ background: #eef2fc; border: 1px solid #cdddf8; }}
    .dmetric.sc {{ background: #e7f5f5; border: 1px solid #b8e6e6; }}
    .dmetric-label {{ font-size: 10px; font-weight: 800; letter-spacing: .6px; text-transform: uppercase; }}
    .dmetric.assy .dmetric-label {{ color: var(--primary); }}
    .dmetric.sc .dmetric-label {{ color: #0e7c8b; }}
    .dmetric-num {{ font-size: 27px; font-weight: 800; line-height: 1.15; font-variant-numeric: tabular-nums; }}
    .dmetric.assy .dmetric-num {{ color: var(--primary-dark); }}
    .dmetric.sc .dmetric-num {{ color: #0c6b78; }}
    .dmetric-num span {{ font-size: 12px; font-weight: 700; color: var(--muted); }}
    .dmetric-bar {{ margin-top: 7px; height: 8px; background: #fff; border: 1px solid rgba(0,0,0,.07); border-radius: 5px; overflow: hidden; }}
    .dmetric-bar i {{ display: block; height: 100%; border-radius: 5px; }}
    .dmetric.assy .dmetric-bar i {{ background: linear-gradient(90deg, #3a548f, #2f56c4); }}
    .dmetric.sc .dmetric-bar i {{ background: linear-gradient(90deg, #0e7c8b, #14b8a6); }}
    .dmetric-sub {{ margin-top: 6px; font-size: 10px; color: var(--muted); font-variant-numeric: tabular-nums; }}
    .dash-rowhead, .dash-row {{
      display: grid;
      grid-template-columns: 64px 1fr 130px;
      align-items: center;
      gap: 14px;
      padding: 9px 0;
    }}
    .dash-rowhead {{
      font-size: 10px;
      font-weight: 800;
      letter-spacing: .5px;
      text-transform: uppercase;
      color: var(--muted);
      border-bottom: 2px solid var(--border);
    }}
    .dash-row {{ border-bottom: 1px solid #eef2f7; }}
    .dash-row .line {{ font-weight: 800; color: var(--primary); font-size: 14px; }}
    .dash-bar-wrap {{ background: #eef2f7; border-radius: 6px; padding: 3px; display: flex; flex-direction: column; gap: 3px; }}
    .dash-bar {{ height: 9px; background: linear-gradient(90deg, #3a548f, #2f56c4); border-radius: 5px; min-width: 2px; }}
    .dash-bar.sc {{ background: linear-gradient(90deg, #0e7c8b, #14b8a6); }}
    .dash-val {{ text-align: right; }}
    .dash-val .v {{ font-weight: 800; font-variant-numeric: tabular-nums; color: var(--ink); }}
    .dash-val .v.sc {{ font-size: 12px; color: #0e7c8b; }}
    .dash-val .s {{ font-size: 10px; color: var(--muted); font-variant-numeric: tabular-nums; }}
    .dash-row.clickable {{ cursor: pointer; }}
    .dash-row.clickable:hover {{ background: #f4f7fe; }}
    .dash-detail {{ padding: 2px 0 10px 64px; }}
    .dash-detail-row {{
      display: grid;
      grid-template-columns: 92px 1fr 64px 60px 60px;
      gap: 10px;
      font-size: 11px;
      color: var(--ink-soft);
      font-variant-numeric: tabular-nums;
      padding: 3px 0;
      border-bottom: 1px dashed #eef2f7;
    }}
    .dash-detail-row .r {{ text-align: right; }}
    .controls {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px 14px;
      align-items: stretch;
      padding: 11px 20px;
      border-top: 1px solid #eef2f7;
      background: var(--surface);
    }}
    .left-stack {{
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      gap: 14px;
      flex: 0 1 auto;
      min-width: 0;
    }}
    .filter-row {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px 14px;
      align-items: end;
    }}
    label {{
      display: grid;
      gap: 4px;
      font-size: 10px;
      font-weight: 700;
      letter-spacing: .5px;
      text-transform: uppercase;
      color: var(--ink-soft);
    }}
    select, input {{
      height: 32px;
      border: 1px solid #cbd5e1;
      border-radius: 8px;
      padding: 4px 10px;
      font: inherit;
      font-size: 12.5px;
      color: var(--ink);
      background: #fff;
      min-width: 110px;
      transition: border-color .15s, box-shadow .15s;
    }}
    select {{
      appearance: none;
      background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2364748b' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");
      background-repeat: no-repeat;
      background-position: right 9px center;
      padding-right: 26px;
    }}
    select:focus, input:focus {{
      outline: none;
      border-color: var(--primary);
      box-shadow: 0 0 0 3px rgba(47, 86, 196, .15);
    }}
    select[multiple] {{
      height: 70px;
      min-width: 190px;
      padding: 4px;
      background-image: none;
    }}
    select[multiple] option {{ padding: 2px 6px; border-radius: 4px; }}
    input[type="search"] {{ min-width: 0; width: 100%; }}
    .control-search {{ flex: 0 1 300px; }}
    .control-actions {{
      display: flex;
      gap: 10px;
      align-items: center;
      margin-left: auto;
    }}
    .import-btn {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      height: 32px;
      border: 1px solid transparent;
      border-radius: 8px;
      padding: 0 16px;
      background: var(--primary);
      color: #fff;
      font: 700 11.5px/1 inherit;
      letter-spacing: .3px;
      cursor: pointer;
      box-shadow: 0 2px 6px rgba(47, 86, 196, .25);
      white-space: nowrap;
      transition: background .15s, box-shadow .15s, transform .05s;
    }}
    .import-btn:hover {{ background: var(--primary-dark); box-shadow: 0 4px 10px rgba(47, 86, 196, .3); }}
    .import-btn:active {{ transform: translateY(1px); }}
    .import-btn::before {{ content: "\\2191"; font-size: 13px; }}
    .import-btn input {{ display: none; }}
    .sap-btn {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      height: 32px;
      border: 1px solid transparent;
      border-radius: 8px;
      padding: 0 16px;
      background: #0a8043;
      color: #fff;
      font: 700 11.5px/1 inherit;
      letter-spacing: .3px;
      cursor: pointer;
      box-shadow: 0 2px 6px rgba(10, 128, 67, .25);
      white-space: nowrap;
      transition: background .15s, box-shadow .15s, transform .05s;
    }}
    .sap-btn::before {{ content: "\\21BB"; font-size: 14px; }}
    .sap-btn:hover {{ background: #0b6e3a; box-shadow: 0 4px 10px rgba(10, 128, 67, .3); }}
    .sap-btn:active {{ transform: translateY(1px); }}
    .sap-btn[disabled] {{ opacity: .6; cursor: progress; }}
    .sap-btn.busy::before {{ animation: sapSpin .8s linear infinite; }}
    @keyframes sapSpin {{ to {{ transform: rotate(360deg); }} }}
    .clear-btn {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      height: 32px;
      border: 1px solid #f1c4c8;
      border-radius: 8px;
      padding: 0 16px;
      background: #fff;
      color: #d62f40;
      font: 700 11.5px/1 inherit;
      letter-spacing: .3px;
      cursor: pointer;
      white-space: nowrap;
      transition: background .15s, border-color .15s;
    }}
    .clear-btn:hover {{ background: #fff1f2; border-color: #e8939b; }}
    .import-status {{
      min-width: 160px;
      max-width: 260px;
      font-size: 11px;
      color: var(--muted);
      align-self: center;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }}
    .lead-panel {{
      display: flex;
      flex-direction: column;
      gap: 7px;
      align-self: stretch;
      flex: 1 1 240px;
      min-width: 0;
      padding: 9px 14px 11px;
      border: 1px solid #d7e0ee;
      border-radius: 14px;
      background: var(--surface-2);
      box-shadow: var(--shadow-sm);
    }}
    .lead-panel-title {{
      font-size: 11px;
      font-weight: 800;
      letter-spacing: .7px;
      text-transform: uppercase;
      color: var(--ink-soft);
      padding-left: 2px;
    }}
    .lead-head {{
      display: flex;
      align-items: center;
      gap: 16px;
      flex-wrap: wrap;
    }}
    .lead-legend {{ display: flex; align-items: center; gap: 11px; flex-wrap: wrap; }}
    .lead-leg-title {{
      font-size: 9px;
      font-weight: 800;
      letter-spacing: .5px;
      text-transform: uppercase;
      color: var(--muted);
    }}
    .ll-item {{
      display: inline-flex;
      align-items: center;
      gap: 5px;
      font-size: 10.5px;
      font-weight: 700;
      color: var(--ink-soft);
      white-space: nowrap;
    }}
    .ll-sw {{ width: 11px; height: 11px; border-radius: 3px; display: inline-block; }}
    .lead-edit-btn {{
      margin-left: auto;
      height: 26px;
      padding: 0 11px;
      border: 1px solid #cbd5e1;
      border-radius: 8px;
      background: #fff;
      color: var(--ink-soft);
      font: 700 10.5px/1 inherit;
      letter-spacing: .3px;
      cursor: pointer;
      white-space: nowrap;
      transition: border-color .15s, color .15s;
    }}
    .lead-edit-btn:hover {{ border-color: var(--primary); color: var(--primary); }}
    .target-panel {{ width: min(560px, 96vw); max-height: 86vh; }}
    .target-note {{ font-size: 11px; color: var(--ink-soft); margin: 4px 2px 12px; line-height: 1.5; }}
    .target-body {{ overflow-y: auto; }}
    .target-grid {{
      display: grid;
      grid-template-columns: 1fr 1fr 1fr;
      gap: 8px 12px;
      align-items: center;
    }}
    .target-grid .th {{
      font-size: 10px;
      font-weight: 800;
      letter-spacing: .5px;
      text-transform: uppercase;
      color: var(--muted);
      padding-bottom: 4px;
      border-bottom: 1px solid var(--border);
    }}
    .target-grid .tline {{ font-size: 15px; font-weight: 800; color: var(--primary-dark); }}
    .target-grid input {{ width: 100%; height: 34px; }}
    .target-actions {{
      display: flex;
      justify-content: flex-end;
      gap: 10px;
      margin-top: 16px;
      padding-top: 14px;
      border-top: 1px solid var(--border);
    }}
    .target-save {{
      display: inline-flex;
      align-items: center;
      height: 34px;
      padding: 0 20px;
      border: 1px solid transparent;
      border-radius: 8px;
      background: var(--primary);
      color: #fff;
      font: 700 12px/1 inherit;
      letter-spacing: .3px;
      cursor: pointer;
      box-shadow: 0 2px 6px rgba(47, 86, 196, .25);
    }}
    .target-save:hover {{ background: var(--primary-dark); }}
    .lead-strip {{
      display: flex;
      align-items: center;
      gap: 8px;
      flex: 1 1 auto;
      overflow-x: auto;
      padding-bottom: 2px;
      scrollbar-width: thin;
    }}
    .lead-strip.empty {{ min-height: 44px; }}
    .lchip {{
      display: flex;
      align-items: center;
      gap: 9px;
      padding: 8px 12px;
      border: 1px solid var(--border);
      border-radius: 12px;
      background: #fff;
      white-space: nowrap;
      flex: 0 0 auto;
    }}
    .lchip-name {{ font-size: 20px; font-weight: 800; color: var(--primary-dark); min-width: 22px; }}
    .lmetric {{
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      line-height: 1.1;
      padding: 5px 11px;
      border: 1px solid transparent;
      border-radius: 9px;
      min-width: 64px;
    }}
    .lmetric-l {{ font-size: 9px; font-weight: 800; letter-spacing: .4px; text-transform: uppercase; }}
    .lmetric-v {{ font-size: 19px; font-weight: 800; color: #111; font-variant-numeric: tabular-nums; }}
    .sim-panel {{
      display: flex;
      flex-direction: column;
      gap: 7px;
      align-self: stretch;
      margin-left: auto;
      padding: 9px 14px 11px;
      border: 1px solid #c2d4f4;
      border-radius: 14px;
      background: linear-gradient(180deg, #eaf1ff 0%, #f4f7fe 100%);
      box-shadow: var(--shadow-sm);
    }}
    .sim-panel-title {{
      font-size: 11px;
      font-weight: 800;
      letter-spacing: .7px;
      text-transform: uppercase;
      color: var(--primary-dark);
      padding-left: 2px;
    }}
    .sim {{
      display: inline-flex;
      align-items: stretch;
      gap: 18px;
      padding: 10px 16px;
      border: 1px solid #cdddf8;
      border-radius: 10px;
      background: #fff;
    }}
    .sim-cell {{
      display: flex;
      flex-direction: column;
      justify-content: flex-start;
      gap: 6px;
      padding-right: 18px;
      border-right: 1px solid #d4e1f6;
    }}
    .sim-cell:last-of-type {{ border-right: none; padding-right: 0; }}
    .sim-input {{
      width: 96px;
      height: 34px;
      border: 1px solid #cdddf8;
      border-radius: 8px;
      padding: 0 10px;
      font-family: inherit;
      font-weight: 800;
      font-size: 22px;
      line-height: 1;
      color: var(--primary);
      background: #fff;
      font-variant-numeric: tabular-nums;
    }}
    .sim-title {{
      font-size: 11px;
      font-weight: 800;
      letter-spacing: .5px;
      text-transform: uppercase;
      color: var(--primary);
      white-space: nowrap;
    }}
    .sim-badge {{
      display: inline-flex;
      align-items: center;
      height: 34px;
      font-size: 22px;
      font-weight: 800;
      line-height: 1;
      color: var(--primary);
      font-variant-numeric: tabular-nums;
    }}
    .sim-badge.danger {{ color: var(--late); }}
    .sim-sub {{ font-size: 11px; font-weight: 700; color: var(--muted); white-space: nowrap; }}
    .sim-clear {{
      height: 32px;
      padding: 0 14px;
      border: 1px solid #cdddf8;
      border-radius: 8px;
      background: #fff;
      color: var(--primary);
      font: 700 11px/1 inherit;
      text-transform: uppercase;
      letter-spacing: .3px;
      cursor: pointer;
      box-shadow: none;
    }}
    .sim-clear:hover {{ background: #fff; border-color: var(--primary); }}
    .sim-clear {{ align-self: center; }}
    .action-bar {{
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      align-items: center;
    }}
    .action-bar .toggle-btn,
    .action-bar .import-btn,
    .action-bar .clear-btn {{
      height: 38px;
      min-width: 170px;
      justify-content: center;
      align-self: auto;
      margin: 0;
    }}
    .action-bar .import-status {{ margin-left: auto; }}
    .toggle-btn {{
      height: 44px;
      align-self: end;
      padding: 0 16px;
      border: 1px solid #cbd5e1;
      border-radius: 10px;
      background: #fff;
      color: var(--ink-soft);
      font: 700 11.5px/1 inherit;
      letter-spacing: .3px;
      cursor: pointer;
      box-shadow: none;
      white-space: nowrap;
      transition: background .15s, border-color .15s, color .15s;
    }}
    .toggle-btn::before {{ content: "\\1F441  "; opacity: .65; }}
    .toggle-btn:hover {{ border-color: var(--primary); color: var(--primary); }}
    .toggle-btn.active {{
      background: var(--primary);
      border-color: var(--primary);
      color: #fff;
    }}
    .toggle-btn.active::before {{ content: "\\1F6AB  "; opacity: 1; }}
    th.selcol {{ cursor: default; padding: 0; }}
    td.selcol {{ text-align: center; padding: 0; }}
    .selall, .rowchk {{
      width: 15px;
      height: 15px;
      min-width: 0;
      margin: 0;
      vertical-align: middle;
      accent-color: var(--primary);
      cursor: pointer;
    }}
    tbody tr:hover td.selcol {{ background: #dfeaff; }}
    .subbar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 14px;
      padding: 9px 20px;
      border-top: 1px solid #eef2f7;
      background: var(--surface-2);
    }}
    .subbar-right {{
      display: flex;
      align-items: center;
      gap: 14px;
      flex: 0 0 auto;
      margin-left: auto;
    }}
    .hidden-bar {{
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 7px;
      min-width: 0;
    }}
    .hidden-bar.empty {{ display: none; }}
    .hb-label {{
      font-size: 10px;
      font-weight: 700;
      letter-spacing: .5px;
      text-transform: uppercase;
      color: var(--muted);
    }}
    .chip {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 3px 9px;
      border: 1px solid var(--border);
      border-radius: 999px;
      background: #fff;
      font-size: 11px;
      font-weight: 600;
      color: var(--ink-soft);
      cursor: pointer;
      transition: border-color .15s, color .15s, background .15s;
    }}
    .chip:hover {{ border-color: var(--primary); color: var(--primary); background: #f4f7fe; }}
    .chip .x {{ font-size: 13px; line-height: 1; opacity: .7; }}
    .hb-showall {{
      border: none;
      background: none;
      padding: 2px 6px;
      font: 700 11px/1 inherit;
      color: var(--primary);
      cursor: pointer;
      text-decoration: underline;
    }}
    .table-wrap {{
      overflow: auto;
      min-height: 0;
      height: 100%;
      background: var(--surface);
    }}
    table {{
      border-collapse: separate;
      border-spacing: 0;
      table-layout: fixed;
      width: max-content;
      min-width: 100%;
    }}
    tr.spacer td {{
      padding: 0;
      border: 0;
      height: 0;
      background: transparent !important;
    }}
    th, td {{
      border-right: 1px solid var(--grid);
      border-bottom: 1px solid var(--grid);
      padding: 4px 9px;
      height: 26px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      max-width: 260px;
    }}
    th {{
      position: sticky;
      top: 0;
      z-index: 4;
      color: #fff;
      background: var(--head);
      text-align: center;
      font-size: 11px;
      font-weight: 700;
      letter-spacing: .2px;
      cursor: default;
      user-select: none;
      box-shadow: inset 0 -2px 0 rgba(0,0,0,.12);
    }}
    th:hover {{ filter: brightness(1.08); }}
    th .th-hide {{
      position: absolute;
      top: 50%;
      right: 3px;
      transform: translateY(-50%);
      width: 15px;
      height: 15px;
      display: none;
      align-items: center;
      justify-content: center;
      border-radius: 4px;
      font-size: 12px;
      line-height: 1;
      color: #fff;
      background: rgba(0, 0, 0, .22);
    }}
    th:hover .th-hide {{ display: inline-flex; }}
    th .th-hide:hover {{ background: rgba(0, 0, 0, .45); }}
    th .arrow {{
      display: inline-block;
      width: 0;
      opacity: 0;
      margin-left: 4px;
      font-size: 9px;
      transition: opacity .15s;
    }}
    th.sorted .arrow {{ width: auto; opacity: .95; }}
    th.c-cyan {{ background: var(--head-cyan); }}
    th.c-green {{ background: var(--head-green); }}
    th.c-yellow {{ background: var(--head-amber); }}
    th.c-purple {{ background: var(--head-purple); }}
    th.c-pink {{ background: var(--head-pink); }}
    th.c-blue {{ background: var(--head-edit); }}
    td.num, td.time, td.date {{ text-align: right; font-variant-numeric: tabular-nums; }}
    td.center {{ text-align: center; }}
    td.edit {{ background: var(--edit-tint); }}
    tbody td {{ color: var(--ink); background: var(--surface); }}
    tbody tr:nth-child(odd) td {{ background: var(--surface-2); }}
    tbody tr:nth-child(odd) td.edit {{ background: #fef6df; }}
    tbody tr:hover td {{ background: #eaf1ff; }}
    tbody tr:hover td.edit {{ background: #fdeeca; }}
    tbody tr.row-locked td {{ background: #d6dae1 !important; color: #8b94a3 !important; }}
    tbody tr.row-locked:hover td {{ background: #ccd1d9 !important; }}
    tbody tr.row-locked td.editable {{ cursor: default; }}
    tbody tr.hidden-wip td {{ display: none; }}
    td.lead {{ font-weight: 700; }}
    tbody tr.row-late td.lead {{ background: #fde7e7 !important; color: var(--late); }}
    tbody tr.row-ready td.lead {{ background: #e6f6ee !important; color: var(--ready); }}
    tbody tr.row-missing td.lead {{ background: #fdf4dc !important; color: #a9790a; }}
    td.editable {{ cursor: text; }}
    td.blocked {{
      background: #9aa5b5 !important;
      cursor: not-allowed;
    }}
    td.editable:focus {{
      outline: 2px solid var(--primary);
      outline-offset: -2px;
      background: #fff !important;
      border-radius: 2px;
    }}
    td.fill-preview {{ outline: 2px solid var(--primary); outline-offset: -2px; background: #dbe6ff !important; }}
    td.fill-sel {{ outline: 2px dashed var(--primary); outline-offset: -2px; background: #e3ecff !important; }}
    td.okcell::after {{
      content: "Done";
      margin-left: 6px;
      font-size: 9px;
      font-weight: 800;
      color: #138a55;
      background: #e6f6ee;
      border: 1px solid #b7ecd0;
      border-radius: 4px;
      padding: 0 4px;
      vertical-align: middle;
    }}
    #fillHandle {{
      position: fixed;
      width: 10px;
      height: 10px;
      background: var(--primary);
      border: 1.5px solid #fff;
      border-radius: 2px;
      box-shadow: 0 0 0 1px rgba(0, 0, 0, .25);
      cursor: crosshair;
      z-index: 50;
      display: none;
    }}
    body.fill-dragging {{ cursor: crosshair; user-select: none; }}
    body.fill-dragging td.editable {{ cursor: crosshair; }}
    .badge {{
      display: inline-block;
      padding: 2px 9px;
      border-radius: 999px;
      font-size: 10.5px;
      font-weight: 700;
      letter-spacing: .3px;
      line-height: 1.5;
      background: #eef2f7;
      color: var(--ink-soft);
      border: 1px solid #dce3ee;
    }}
    .badge.ok {{ background: #e6f6ee; color: #138a55; border-color: #b7ecd0; }}
    .badge.warn {{ background: #fdf4dc; color: #a9790a; border-color: #f0deb0; }}
    .badge.info {{ background: #e8effc; color: #2f56c4; border-color: #cdddf8; }}
    .blank {{ color: transparent; }}
    .muted {{ color: var(--muted); }}
    .red {{ color: var(--late); font-weight: 700; }}
    .small {{ font-size: 11px; }}
    @media print {{
      .top {{ display: none; }}
      .table-wrap {{ height: auto; overflow: visible; }}
      body {{ background: #fff; }}
      table {{ min-width: 0; font-size: 9px; }}
      th, td {{ padding: 1px 2px; }}
    }}
    @media (max-width: 900px) {{
      .banner {{ height: 54px; padding: 0 12px; }}
      .brand {{ min-width: 0; }}
      .brand-title {{ font-size: 12px; }}
      .brand-subtitle {{ font-size: 8px; letter-spacing: 1.4px; }}
      .connection {{ padding: 5px 10px; font-size: 9px; }}
      .controls {{ padding: 10px 12px; }}
      .control-actions {{ margin-left: 0; }}
      input[type="search"] {{ min-width: 0; }}
    }}
  </style>
</head>
<body>
  <div class="app">
    <section class="top">
      <div class="banner">
        <div class="brand" aria-label="HEI Part Order System">
          <span class="logo">HEI</span>
          <div>
            <div class="brand-title">DAILY PRODUCTION PLANNING</div>
            <div class="brand-subtitle">SMART FACTORY - REAL-TIME</div>
          </div>
        </div>
        <button id="dashBtn" class="dash-btn" type="button">&#128202; Dashboard</button>
      </div>
      <div class="controls">
        <div class="left-stack">
          <div class="filter-row">
            <label>Line
              <select id="lineSelect"></select>
            </label>
            <label>Prod Month
              <select id="monthSelect"></select>
            </label>
            <label>Status
              <select id="statusSelect"></select>
            </label>
            <label class="control-search">Search
              <input id="searchBox" type="search" placeholder="Code / Model / Item / Order / Remark…">
            </label>
          </div>
          <div class="action-bar">
            <button id="hideHairBtn" class="toggle-btn" type="button" aria-pressed="false" title="Hide / show rows whose Description is HAIR PIN, HAIR TUBE, or HPIN">Hide Hair Pin</button>
            <label class="import-btn">Update Progress
              <input id="importFile" type="file" accept=".xlsx,.xlsm">
            </label>
            <label class="import-btn">Update Stock
              <input id="importStock" type="file" accept=".xlsx,.xlsm">
            </label>
            <button id="zpp0059Btn" class="sap-btn" type="button" title="Pull progress straight from SAP (runs transaction ZPP0059, then refreshes the table automatically)">ZPP0059</button>
            <button id="dbViewBtn" class="dash-btn" type="button" title="View the ZPP0059 database (all data collected from SAP)">&#128451; View Database</button>
          </div>
        </div>
        <div class="lead-panel">
          <div class="lead-head">
            <span class="lead-panel-title">Target Assy Leadtime (Hrs.)</span>
            <div class="lead-legend">
              <span class="lead-leg-title">Target</span>
              <span class="ll-item"><span class="ll-sw" style="background:#e23b3b"></span>&lt; 4 hr</span>
              <span class="ll-item"><span class="ll-sw" style="background:#e0a200"></span>4-10 hr</span>
              <span class="ll-item"><span class="ll-sw" style="background:#1f9254"></span>&gt;10-20 hr</span>
              <span class="ll-item"><span class="ll-sw" style="background:#2f56c4"></span>&gt;20-40 hr</span>
              <span class="ll-item"><span class="ll-sw" style="background:#7c3aed"></span>&gt;40 hr</span>
            </div>
            <button id="editTargetBtn" class="lead-edit-btn" type="button" title="Edit Assy / SC leadtime targets per line">&#9881; Edit Target</button>
          </div>
          <div class="lead-strip empty" id="leadStrip"></div>
        </div>
        <div class="sim-panel">
          <span class="sim-panel-title">Productivity Simulator</span>
          <div class="sim" id="simBox" title="Tick rows at the left to plan today's run. Unit Rate = sum of (TS x Assy Order qty) for the ticked rows.">
            <div class="sim-cell">
              <span class="sim-title">Unit Rate</span>
              <span class="sim-badge" id="simBadge">0</span>
              <span class="sim-sub" id="simSub">0 rows</span>
            </div>
            <div class="sim-cell">
              <span class="sim-title">Working</span>
              <input id="workHour" class="sim-input" type="number" min="0" step="0.5" placeholder="0" title="Enter available working hours">
              <span class="sim-sub">Hrs</span>
            </div>
            <div class="sim-cell">
              <span class="sim-title">Result</span>
              <span class="sim-badge" id="prodBadge" title="Working hour / Unit rate">0</span>
              <span class="sim-sub">Hrs./Unit rate</span>
            </div>
            <button id="simClear" class="sim-clear" type="button">Clear</button>
          </div>
        </div>
      </div>
      <div class="subbar">
        <div class="hidden-bar empty" id="hiddenBar"></div>
        <div class="subbar-right">
          <span id="importStatus" class="import-status">Loaded: {escape(EXPORT_FILE.name)}</span>
          <button id="clearProgressBtn" class="clear-btn" type="button" title="Clear entered values in H/P, FP, EXP, Auto, Cutting, FG, Assy, Subcooler">Clear Progress</button>
          <button id="clearDataBtn" class="clear-btn" type="button" title="Remove all loaded data">Reset</button>
        </div>
      </div>
    </section>
    <main class="table-wrap">
      <table id="planTable">
        <colgroup id="colGroup"></colgroup>
        <thead><tr id="headerRow"></tr></thead>
        <tbody id="bodyRows"></tbody>
      </table>
    </main>
  </div>
  <div id="fillHandle"></div>
  <div id="dashOverlay" class="dash-overlay" hidden>
    <div class="dash-panel">
      <div class="dash-head">
        <div class="dash-title">Lead Time Summary by Line</div>
        <button id="dashClose" class="dash-close" type="button">&#10005;</button>
      </div>
      <div id="dashBody" class="dash-body"></div>
    </div>
  </div>
  <div id="dbOverlay" class="dash-overlay" hidden>
    <div class="dash-panel db-panel">
      <div class="dash-head">
        <div class="dash-title">ZPP0059 Database</div>
        <button id="dbClose" class="dash-close" type="button">&#10005;</button>
      </div>
      <div class="db-toolbar">
        <input id="dbSearch" type="search" placeholder="Search all columns…">
        <button id="dbRefresh" class="db-refresh" type="button">&#8635; Refresh</button>
        <span id="dbMeta" class="db-meta"></span>
      </div>
      <div id="dbTableWrap" class="db-table-wrap"></div>
    </div>
  </div>
  <div id="targetOverlay" class="dash-overlay" hidden>
    <div class="dash-panel target-panel">
      <div class="dash-head">
        <div class="dash-title">Edit Leadtime Targets (Hrs.)</div>
        <button id="targetClose" class="dash-close" type="button">&#10005;</button>
      </div>
      <div class="target-note">Assy: leadtime below the target shows red. SC: leadtime at or below the target shows red, above shows yellow. Green / blue / purple bands (10 / 20 / 40 hr) are shared.</div>
      <div id="targetBody" class="target-body"></div>
      <div class="target-actions">
        <button id="targetReset" class="clear-btn" type="button">Reset to default</button>
        <button id="targetSave" class="target-save" type="button">Save</button>
      </div>
    </div>
  </div>
  <script id="daily-data" type="application/json">{payload}</script>
  <script>
    const DATA = JSON.parse(document.getElementById('daily-data').textContent);
    let currentRows = DATA.rows;
    const columns = [
      ['_sel','','selcol',36],
      ['line','Line','center',46], ['seq','Seq.','num',54], ['month','Month','num',62], ['code','Code','',88],
      ['model','Model Name','',150], ['finished','Finished','center',70], ['orderQty','Order Quantity','num',92],
      ['item','Item','',112], ['description','Description','',180], ['attribute','Attribute','',112],
      ['productionOrder','Production Order','num',112], ['speed','Speed','num',58], ['prodDate','Prod.Date','date',92],
      ['prodTime','Prod.Time','time',72], ['assyOrder','Assy Order','num',80], ['assyOrderNo','Assy Order No.','num',104],
      ['status','Status','center',76], ['remark','Remark','',230], ['hp','H/P','num edit',52],
      ['fp','FP','num edit',52], ['exp','EXP','num edit',52], ['auto','Auto','num edit',54],
      ['cutting','Cutting','num edit',64], ['fg','FG.','num edit',52], ['unit','UNIT','num',58],
      ['stockFg','Assy','num edit',74], ['subcooler','Subcooler','num edit',84], ['lead','Assy LT (Hrs.)','num lead',96], ['scLead','SC LT (Hrs.)','num lead',90],
      ['leadRemark','Remark','',130], ['ts','TS','num',58], ['unitRate','Unit Rate','num',100]
    ];
    const headerClass = {{
      item: 'c-cyan', description: 'c-green', attribute: 'c-yellow', productionOrder: 'c-purple',
      speed: 'c-pink', status: 'c-purple', hp: 'c-blue', fp: 'c-blue', exp: 'c-blue',
      auto: 'c-blue', cutting: 'c-blue', fg: 'c-blue', unit: 'c-blue', stockFg: 'c-blue',
      lead: 'c-blue', leadRemark: 'c-green', ts: 'c-green', unitRate: 'c-green',
      subcooler: 'c-blue', scLead: 'c-blue'
    }};
    const editable = new Set(['hp','fp','exp','auto','cutting','fg','stockFg','subcooler','leadRemark']);
    const els = {{
      line: document.getElementById('lineSelect'), month: document.getElementById('monthSelect'),
      status: document.getElementById('statusSelect'),
      search: document.getElementById('searchBox'), hiddenBar: document.getElementById('hiddenBar'),
      body: document.getElementById('bodyRows'), colgroup: document.getElementById('colGroup'),
      wrap: document.querySelector('.table-wrap'),
      header: document.getElementById('headerRow'), importFile: document.getElementById('importFile'),
      importStock: document.getElementById('importStock'), zpp0059Btn: document.getElementById('zpp0059Btn'),
      importStatus: document.getElementById('importStatus'), clearData: document.getElementById('clearDataBtn'),
      clearProgress: document.getElementById('clearProgressBtn'),
      simBadge: document.getElementById('simBadge'), simSub: document.getElementById('simSub'),
      leadStrip: document.getElementById('leadStrip'),
      editTargetBtn: document.getElementById('editTargetBtn'),
      targetOverlay: document.getElementById('targetOverlay'), targetClose: document.getElementById('targetClose'),
      targetBody: document.getElementById('targetBody'), targetReset: document.getElementById('targetReset'),
      targetSave: document.getElementById('targetSave'),
      workHour: document.getElementById('workHour'), prodBadge: document.getElementById('prodBadge'),
      simClear: document.getElementById('simClear'), hideHairBtn: document.getElementById('hideHairBtn'),
      fillHandle: document.getElementById('fillHandle'),
      dashBtn: document.getElementById('dashBtn'), dashOverlay: document.getElementById('dashOverlay'),
      dashClose: document.getElementById('dashClose'), dashBody: document.getElementById('dashBody'),
      dbViewBtn: document.getElementById('dbViewBtn'), dbOverlay: document.getElementById('dbOverlay'),
      dbClose: document.getElementById('dbClose'), dbSearch: document.getElementById('dbSearch'),
      dbRefresh: document.getElementById('dbRefresh'), dbMeta: document.getElementById('dbMeta'),
      dbTableWrap: document.getElementById('dbTableWrap')
    }};
    let hideSkip = false;
    let hideHairPin = localStorage.getItem('dailyFollowHideHair') === '1';
    let okSet = new Set(JSON.parse(localStorage.getItem('dailyFollowOk') || '[]'));
    let hidden = new Set(JSON.parse(localStorage.getItem('dailyFollowHiddenCols') || '[]'));
    const columnTitle = Object.fromEntries(columns.map(([field, title]) => [field, title]));
    let sortKey = 'prodSort';
    let sortDir = 1;
    let saved = JSON.parse(localStorage.getItem('dailyFollowProgress') || '{{}}');
    let sortedAll = [];
    let filtered = [];
    let needSort = true;
    let activeCols = [];
    let winStart = -1, winEnd = -1;
    let ROW_H = 27, measuredRowH = false;
    let searchTimer = 0;
    let scrollTicking = false;
    let selected = new Set();
    let byId = new Map();
    let fillSource = null;
    let fillDragging = false;
    let fillRange = null;
    let pendingSel = null;
    let undoStack = [];
    let fillScrollTimer = 0;
    let lastMouse = {{ x: 0, y: 0 }};
    let lastChkIdx = null;
    let pendingChkModifier = false;
    let targets = loadTargets();
    function loadTargets() {{
      let t = {{ assy: {{}}, sc: {{}} }};
      try {{
        const saved = JSON.parse(localStorage.getItem('dailyFollowTargets') || '{{}}');
        if (saved && typeof saved === 'object') {{
          t.assy = saved.assy && typeof saved.assy === 'object' ? saved.assy : {{}};
          t.sc = saved.sc && typeof saved.sc === 'object' ? saved.sc : {{}};
        }}
      }} catch (e) {{ t = {{ assy: {{}}, sc: {{}} }}; }}
      return t;
    }}
    function saveTargets() {{
      localStorage.setItem('dailyFollowTargets', JSON.stringify(targets));
    }}

    let leadGroups = new Map();
    let leadDirty = true;
    let progressCleared = localStorage.getItem('dailyFollowProgCleared') === '1';
    const progressFields = new Set(['hp','fp','exp','auto','cutting','fg','stockFg','subcooler']);
    function keyOf(row, field) {{
      return `${{row.productionOrder}}|${{row.item}}|${{field}}`;
    }}
    function valueOf(row, field) {{
      const k = keyOf(row, field);
      if (Object.prototype.hasOwnProperty.call(saved, k)) return saved[k];
      if (progressCleared && progressFields.has(field)) return '';
      return row[field];
    }}
    function groupKey(row) {{
      return `${{row.line}}|${{row.month}}|${{row.seq}}`;
    }}
    function buildLeadGroups() {{
      leadGroups = new Map();
      for (const row of currentRows) {{
        if (isHairPin(row)) continue;
        const k = groupKey(row);
        let g = leadGroups.get(k);
        if (!g) {{ g = {{ stocks: [], scs: [], stockGate: true, scGate: true }}; leadGroups.set(k, g); }}
        if (!okSet.has(keyOf(row, 'stockFg'))) {{
          const s = valueOf(row, 'stockFg');
          if (s === '' || s === null || s === undefined) g.stockGate = false;
          else g.stocks.push(num(s));
        }}
        if (!okSet.has(keyOf(row, 'subcooler'))) {{
          const s = valueOf(row, 'subcooler');
          if (s === '' || s === null || s === undefined) g.scGate = false;
          else g.scs.push(num(s));
        }}
      }}
      for (const g of leadGroups.values()) {{
        g.stockMin = g.stocks.length ? Math.min(...g.stocks) : null;
        g.scMin = g.scs.length ? Math.min(...g.scs) : null;
      }}
      leadDirty = false;
    }}
    function ensureLeadGroups() {{
      if (leadDirty) buildLeadGroups();
    }}
    function fmt(value) {{
      if (value === null || value === undefined || value === '') return '';
      return String(value);
    }}
    function num(value) {{
      const n = Number(value);
      return Number.isFinite(n) ? n : 0;
    }}
    function focusCell(td) {{
      td.focus();
      const range = document.createRange();
      range.selectNodeContents(td);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }}
    function computeUnitRate(row) {{
      if (isHairPin(row) || assyOk(row)) return '';
      const ts = valueOf(row, 'ts');
      const stock = valueOf(row, 'stockFg');
      if (ts === '' || ts === null || ts === undefined) return '';
      if (stock === '' || stock === null || stock === undefined) return '';
      const v = num(ts) * num(stock);
      if (!v) return '';
      return Math.round(v * 10) / 10;
    }}
    function computeLead(row) {{
      if (isHairPin(row) || assyOk(row)) return '';
      ensureLeadGroups();
      const g = leadGroups.get(groupKey(row));
      if (!g || !g.stockGate || g.stockMin === null) return '';
      const speed = row.speed;
      if (speed === '' || speed === null || speed === undefined) return '';
      return Math.round((g.stockMin * num(speed) / 3600) * 100) / 100;
    }}
    function isHairPin(row) {{
      return /hair\\s*(pin|tube)|hpin/i.test(String(row.description || ''));
    }}
    function attrEndsBz(row) {{
      return /-BZ$/i.test(String(row.attribute || '').trim());
    }}
    function assyOk(row) {{
      return okSet.has(keyOf(row, 'stockFg'));
    }}
    function scOk(row) {{
      return okSet.has(keyOf(row, 'subcooler'));
    }}
    function computeScLead(row) {{
      if (isHairPin(row) || scOk(row)) return '';
      ensureLeadGroups();
      const g = leadGroups.get(groupKey(row));
      if (!g || !g.scGate || g.scMin === null) return '';
      const speed = row.speed;
      if (speed === '' || speed === null || speed === undefined) return '';
      return Math.round((g.scMin * num(speed) / 3600) * 100) / 100;
    }}
    function option(select, value, label = value) {{
      const opt = document.createElement('option');
      opt.value = value;
      opt.textContent = label;
      select.appendChild(opt);
      return opt;
    }}
    function fillSelect(select, values, selected = 'ALL') {{
      select.replaceChildren();
      option(select, 'ALL', 'ALL');
      values.forEach(v => option(select, v));
      select.value = values.includes(selected) ? selected : 'ALL';
    }}
    function saveHidden() {{
      localStorage.setItem('dailyFollowHiddenCols', JSON.stringify([...hidden]));
    }}
    function hideColumn(field) {{
      hidden.add(field);
      saveHidden();
      render();
    }}
    function showColumn(field) {{
      hidden.delete(field);
      saveHidden();
      render();
    }}
    function renderHiddenBar() {{
      els.hiddenBar.replaceChildren();
      if (!hidden.size) {{
        els.hiddenBar.classList.add('empty');
        return;
      }}
      els.hiddenBar.classList.remove('empty');
      const label = document.createElement('span');
      label.className = 'hb-label';
      label.textContent = 'Hidden columns:';
      els.hiddenBar.appendChild(label);
      for (const [field, title] of columns) {{
        if (!hidden.has(field)) continue;
        const chip = document.createElement('button');
        chip.type = 'button';
        chip.className = 'chip';
        chip.title = 'Click to show this column again';
        chip.append(title);
        const x = document.createElement('span');
        x.className = 'x';
        x.textContent = '\\u00D7';
        chip.appendChild(x);
        chip.addEventListener('click', () => showColumn(field));
        els.hiddenBar.appendChild(chip);
      }}
      const showAll = document.createElement('button');
      showAll.type = 'button';
      showAll.className = 'hb-showall';
      showAll.textContent = 'Show all';
      showAll.addEventListener('click', () => {{ hidden.clear(); saveHidden(); render(); }});
      els.hiddenBar.appendChild(showAll);
    }}
    function refreshFilters() {{
      fillSelect(els.line, [...new Set(currentRows.map(r => r.line).filter(Boolean))].sort(), els.line.value || DATA.defaultLine || 'ALL');
      fillSelect(els.month, [...new Set(currentRows.map(r => r.month).filter(Boolean))].sort((a,b) => a.localeCompare(b, 'th', {{numeric:true}})), els.month.value || 'ALL');
      fillSelect(els.status, [...new Set(currentRows.map(r => r.status).filter(Boolean))].sort(), els.status.value || 'ALL');
    }}
    function setup() {{
      indexRows(currentRows);
      restoreSelected();
      refreshFilters();
      els.workHour.value = localStorage.getItem('dailyFollowWorkHour') || '';
      els.line.value = DATA.defaultLine || 'ALL';
      els.month.value = 'ALL';
      els.hideHairBtn.classList.toggle('active', hideHairPin);
      els.hideHairBtn.setAttribute('aria-pressed', hideHairPin ? 'true' : 'false');
      els.line.addEventListener('change', () => recompute(true));
      els.month.addEventListener('change', () => recompute(true));
      els.status.addEventListener('change', () => recompute(true));
      els.search.addEventListener('input', () => {{
        clearTimeout(searchTimer);
        searchTimer = setTimeout(() => recompute(true), 160);
      }});
      els.hideHairBtn.addEventListener('click', () => {{
        hideHairPin = !hideHairPin;
        localStorage.setItem('dailyFollowHideHair', hideHairPin ? '1' : '0');
        els.hideHairBtn.classList.toggle('active', hideHairPin);
        els.hideHairBtn.setAttribute('aria-pressed', hideHairPin ? 'true' : 'false');
        recompute(true);
      }});
      els.simClear.addEventListener('click', () => {{
        selected.clear();
        saveSelected();
        if (els.selAll) {{ els.selAll.checked = false; els.selAll.indeterminate = false; }}
        renderWindow(true);
        updateSimReadout();
      }});
      els.workHour.addEventListener('input', () => {{
        localStorage.setItem('dailyFollowWorkHour', els.workHour.value);
        updateSimReadout();
      }});
      els.body.addEventListener('click', (e) => {{
        const cb = e.target;
        if (cb && cb.classList && cb.classList.contains('rowchk')) {{
          pendingChkModifier = e.shiftKey || e.ctrlKey || e.metaKey;
        }}
      }});
      els.body.addEventListener('change', (e) => {{
        const cb = e.target;
        if (!cb || !cb.classList || !cb.classList.contains('rowchk')) return;
        const tr = cb.closest('tr');
        const row = tr && tr.__row;
        if (!row) return;
        const idx = filtered.indexOf(row);
        const useRange = pendingChkModifier && lastChkIdx !== null && idx >= 0 && idx !== lastChkIdx;
        pendingChkModifier = false;
        if (useRange) {{
          const a = Math.min(lastChkIdx, idx), b = Math.max(lastChkIdx, idx);
          for (let i = a; i <= b; i++) {{
            const r = filtered[i];
            if (!r || isHairPin(r)) continue;
            if (cb.checked) selected.add(r._id); else selected.delete(r._id);
          }}
          renderWindow(true);
        }} else {{
          if (cb.checked) selected.add(row._id); else selected.delete(row._id);
        }}
        lastChkIdx = idx;
        saveSelected();
        updateSimReadout();
        if (els.selAll) {{
          const all = allFilteredSelected();
          els.selAll.checked = all;
          els.selAll.indeterminate = !all && filtered.some(r => selected.has(r._id));
        }}
      }});
      els.wrap.addEventListener('scroll', onScroll, {{ passive: true }});
      window.addEventListener('resize', () => {{ renderWindow(true); positionFillHandle(); }});
      els.body.addEventListener('dblclick', (e) => {{
        const td = e.target && e.target.closest ? e.target.closest('td') : null;
        if (!td) return;
        const field = td.dataset.field;
        if (field !== 'stockFg' && field !== 'subcooler') return;
        const tr = td.closest('tr');
        const row = tr && tr.__row;
        if (!row) return;
        e.preventDefault();
        const k = keyOf(row, field);
        if (okSet.has(k)) okSet.delete(k); else okSet.add(k);
        localStorage.setItem('dailyFollowOk', JSON.stringify([...okSet]));
        leadDirty = true;
        renderWindow(true);
        renderLeadStrip();
      }});
      els.body.addEventListener('keydown', (e) => {{
        if (e.isComposing) return;
        if (pendingSel) {{
          if (e.key === 'Enter') {{ e.preventDefault(); commitPendingFill(); return; }}
          if (e.key === 'Delete') {{ e.preventDefault(); clearPendingSel(); return; }}
          if (e.key === 'Escape') {{ e.preventDefault(); cancelPendingSel(); return; }}
        }}
        if (e.key !== 'Enter') return;
        const td = e.target;
        if (!td || !td.classList || !td.classList.contains('editable')) return;
        e.preventDefault();
        let next = td.nextElementSibling;
        while (next && !next.classList.contains('editable')) next = next.nextElementSibling;
        if (next) focusCell(next);
        else td.blur();
      }});
      els.body.addEventListener('focusin', (e) => {{
        const td = e.target;
        if (!td || !td.classList || !td.classList.contains('editable')) return;
        td.classList.remove('blank');
        cancelPendingSel();
        const tr = td.closest('tr');
        fillSource = {{ td: td, field: td.dataset.field, idx: Number(tr.dataset.idx), ci: Number(td.dataset.ci) }};
        positionFillHandle();
      }});
      els.fillHandle.addEventListener('mousedown', (e) => {{
        e.preventDefault();
        if (!fillSource || !document.body.contains(fillSource.td)) return;
        cancelPendingSel();
        fillDragging = true;
        fillSource.value = fillSource.td.textContent.trim();
        fillRange = [fillSource.idx, fillSource.idx, fillSource.ci, fillSource.ci];
        document.body.classList.add('fill-dragging');
        fillScrollTimer = setInterval(fillAutoScroll, 60);
      }});
      document.addEventListener('mousemove', (e) => {{
        if (!fillDragging) return;
        lastMouse.x = e.clientX; lastMouse.y = e.clientY;
        updateFillTargetFromPoint();
        highlightFillRange();
      }});
      document.addEventListener('mouseup', () => {{
        if (!fillDragging) return;
        fillDragging = false;
        document.body.classList.remove('fill-dragging');
        if (fillScrollTimer) {{ clearInterval(fillScrollTimer); fillScrollTimer = 0; }}
        els.body.querySelectorAll('td.fill-preview').forEach(td => td.classList.remove('fill-preview'));
        if (fillRange && fillSource) {{
          pendingSel = {{ rect: fillRange.slice(), value: fillSource.value }};
          fillRange = null;
          highlightPending();
        }}
      }});
      document.addEventListener('keydown', (e) => {{
        if ((e.ctrlKey || e.metaKey) && (e.key === 'z' || e.key === 'Z')) {{
          e.preventDefault();
          undo();
        }} else if (e.key === 'Escape' && !els.dashOverlay.hidden) {{
          closeDashboard();
        }} else if (e.key === 'Escape' && els.dbOverlay && !els.dbOverlay.hidden) {{
          els.dbOverlay.hidden = true;
        }} else if (e.key === 'Escape' && !els.targetOverlay.hidden) {{
          closeTargetEditor();
        }}
      }});
      els.body.addEventListener('focusout', (e) => {{
        const td = e.target;
        if (!td || !td.classList || !td.classList.contains('editable')) return;
        if (!fillDragging) els.fillHandle.style.display = 'none';
        const row = td.parentElement && td.parentElement.__row;
        const field = td.dataset.field;
        if (!row || !field) return;
        const text = td.textContent.trim();
        applyChanges([{{ key: keyOf(row, field), value: text }}]);
        td.classList.toggle('blank', !text);
        const ao = row.assyOrder;
        td.classList.toggle('red', field !== 'leadRemark' && text !== '' && ao !== '' && ao !== null && ao !== undefined && Number(text) !== Number(ao));
        // An edit to stockFg/subcooler changes the whole seq group's leadtime,
        // so re-render the visible window rather than just this row's cell.
        if (field === 'stockFg' || field === 'subcooler') {{ renderWindow(true); renderLeadStrip(); }}
      }});
      els.importFile.addEventListener('change', importRawData);
      els.importStock.addEventListener('change', importStockData);
      if (els.zpp0059Btn) els.zpp0059Btn.addEventListener('click', runZpp0059);
      els.clearData.addEventListener('click', clearData);
      els.clearProgress.addEventListener('click', clearProgress);
      els.dashBtn.addEventListener('click', openDashboard);
      if (els.dbViewBtn) els.dbViewBtn.addEventListener('click', openDatabase);
      if (els.dbClose) els.dbClose.addEventListener('click', () => {{ els.dbOverlay.hidden = true; }});
      if (els.dbOverlay) els.dbOverlay.addEventListener('click', (e) => {{ if (e.target === els.dbOverlay) els.dbOverlay.hidden = true; }});
      if (els.dbRefresh) els.dbRefresh.addEventListener('click', () => loadDatabase(els.dbSearch.value.trim()));
      if (els.dbSearch) els.dbSearch.addEventListener('keydown', (e) => {{ if (e.key === 'Enter') loadDatabase(els.dbSearch.value.trim()); }});
      els.dashClose.addEventListener('click', closeDashboard);
      els.dashOverlay.addEventListener('click', (e) => {{ if (e.target === els.dashOverlay) closeDashboard(); }});
      els.editTargetBtn.addEventListener('click', openTargetEditor);
      els.targetClose.addEventListener('click', closeTargetEditor);
      els.targetSave.addEventListener('click', saveTargetEditor);
      els.targetReset.addEventListener('click', resetTargetEditor);
      els.targetOverlay.addEventListener('click', (e) => {{ if (e.target === els.targetOverlay) closeTargetEditor(); }});
      recompute(true);
    }}
    function searchText(row) {{
      return [row.line,row.seq,row.month,row.code,row.model,row.item,row.description,row.attribute,row.productionOrder,row.assyOrderNo,row.status,row.remark,row.leadRemark].join(' ').toLowerCase();
    }}
    function indexRows(rows) {{
      byId = new Map();
      rows.forEach((row, i) => {{
        row._hay = searchText(row);
        row._id = i;
        byId.set(i, row);
      }});
      selected.clear();
    }}
    function passes(row, q) {{
      if (els.line.value !== 'ALL' && row.line !== els.line.value) return false;
      if (els.month.value !== 'ALL' && row.month !== els.month.value) return false;
      if (els.status.value !== 'ALL' && row.status !== els.status.value) return false;
      if (hideSkip && row.attribute1 && row.attribute1 !== 'FG') return false;
      if (hideHairPin && isHairPin(row)) return false;
      if (q && !(row._hay || '').includes(q)) return false;
      return true;
    }}
    function sortRows(rows) {{
      const computed = {{ prodSort: r => r.prodSort, lead: computeLead, unitRate: computeUnitRate }};
      const getv = computed[sortKey];
      return rows.sort((a,b) => {{
        let av = getv ? getv(a) : valueOf(a, sortKey);
        let bv = getv ? getv(b) : valueOf(b, sortKey);
        const an = Number(av), bn = Number(bv);
        if (Number.isFinite(an) && Number.isFinite(bn)) return (an - bn) * sortDir;
        av = fmt(av); bv = fmt(bv);
        return av.localeCompare(bv, 'th', {{numeric:true}}) * sortDir;
      }});
    }}
    function excelSerialToDate(serial) {{
      if (!Number.isFinite(serial)) return '';
      const utc = Math.round((serial - 25569) * 86400 * 1000);
      return new Date(utc).toISOString().slice(0, 10);
    }}
    function excelSerialToTime(serial) {{
      if (!Number.isFinite(serial)) return '';
      const seconds = Math.round((serial - Math.floor(serial)) * 86400);
      const h = String(Math.floor(seconds / 3600) % 24).padStart(2, '0');
      const m = String(Math.floor((seconds % 3600) / 60)).padStart(2, '0');
      return `${{h}}:${{m}}`;
    }}
    function normalizeDate(value) {{
      if (value === null || value === undefined || value === '') return '';
      if (typeof value === 'number') return excelSerialToDate(value);
      const raw = String(value).trim();
      if (!raw) return '';
      const first = raw.split(' ')[0];
      if (/^\\d{{4}}-\\d{{2}}-\\d{{2}}/.test(first)) return first.slice(0, 10);
      const parsed = new Date(raw);
      return Number.isNaN(parsed.getTime()) ? first : parsed.toISOString().slice(0, 10);
    }}
    function normalizeTime(value) {{
      if (value === null || value === undefined || value === '') return '';
      if (typeof value === 'number') return excelSerialToTime(value);
      const raw = String(value).trim().split(' ').pop();
      const parts = raw.split(':');
      if (parts.length >= 2) return `${{parts[0].padStart(2, '0')}}:${{parts[1].padStart(2, '0')}}`;
      return raw;
    }}
    function parseDateTime(d, t) {{
      const date = normalizeDate(d);
      const time = normalizeTime(t) || '00:00';
      const parsed = date ? new Date(`${{date}}T${{time}}`) : null;
      return parsed && !Number.isNaN(parsed.getTime()) ? parsed : null;
    }}
    function monthDisplay(value) {{
      if (value === null || value === undefined || value === '') return '';
      const raw = String(value).trim();
      const match = raw.match(/^(\\d+)(?:\\.0+)?\\.(\\d{{4}})$/);
      return match ? `${{Number(match[1])}}.${{match[2]}}` : raw;
    }}
    function cleanNumber(value) {{
      if (value === null || value === undefined || value === '') return '';
      const n = Number(String(value).trim());
      if (!Number.isFinite(n)) return value;
      return Number.isInteger(n) ? n : Math.round(n * 1000) / 1000;
    }}
    function cleanText(value) {{
      if (value === null || value === undefined) return '';
      return String(value).trim();
    }}
    async function inflateRaw(bytes) {{
      const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      return new Uint8Array(await new Response(stream).arrayBuffer());
    }}
    function u16(bytes, offset) {{
      return bytes[offset] | (bytes[offset + 1] << 8);
    }}
    function u32(bytes, offset) {{
      return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0;
    }}
    async function unzipXlsx(arrayBuffer) {{
      if (!('DecompressionStream' in window)) {{
        throw new Error('This browser cannot read XLSX files locally. Please use current Chrome or Edge.');
      }}
      const bytes = new Uint8Array(arrayBuffer);
      let eocd = -1;
      for (let i = bytes.length - 22; i >= 0 && i > bytes.length - 66000; i--) {{
        if (u32(bytes, i) === 0x06054b50) {{ eocd = i; break; }}
      }}
      if (eocd < 0) throw new Error('Invalid XLSX file.');
      const entries = u16(bytes, eocd + 10);
      let ptr = u32(bytes, eocd + 16);
      const files = {{}};
      for (let i = 0; i < entries; i++) {{
        if (u32(bytes, ptr) !== 0x02014b50) throw new Error('Invalid ZIP directory.');
        const method = u16(bytes, ptr + 10);
        const compressedSize = u32(bytes, ptr + 20);
        const uncompressedSize = u32(bytes, ptr + 24);
        const nameLen = u16(bytes, ptr + 28);
        const extraLen = u16(bytes, ptr + 30);
        const commentLen = u16(bytes, ptr + 32);
        const localOffset = u32(bytes, ptr + 42);
        const name = new TextDecoder().decode(bytes.slice(ptr + 46, ptr + 46 + nameLen));
        const localNameLen = u16(bytes, localOffset + 26);
        const localExtraLen = u16(bytes, localOffset + 28);
        const dataStart = localOffset + 30 + localNameLen + localExtraLen;
        const compressed = bytes.slice(dataStart, dataStart + compressedSize);
        let content;
        if (method === 0) content = compressed;
        else if (method === 8) content = await inflateRaw(compressed);
        else throw new Error(`Unsupported ZIP compression method ${{method}}.`);
        files[name] = content.slice(0, uncompressedSize);
        ptr += 46 + nameLen + extraLen + commentLen;
      }}
      return files;
    }}
    function xmlText(bytes) {{
      return new TextDecoder('utf-8').decode(bytes || new Uint8Array());
    }}
    function parseXml(text) {{
      const doc = new DOMParser().parseFromString(text, 'application/xml');
      if (doc.querySelector('parsererror')) throw new Error('Could not read XLSX XML.');
      return doc;
    }}
    function readSharedStrings(files) {{
      const xml = xmlText(files['xl/sharedStrings.xml']);
      if (!xml) return [];
      const doc = parseXml(xml);
      return [...doc.getElementsByTagName('si')].map(si =>
        [...si.getElementsByTagName('t')].map(t => t.textContent || '').join('')
      );
    }}
    function firstWorksheetPath(files) {{
      const workbook = parseXml(xmlText(files['xl/workbook.xml']));
      const firstSheet = workbook.getElementsByTagName('sheet')[0];
      if (!firstSheet) throw new Error('Workbook has no sheets.');
      const relId = firstSheet.getAttribute('r:id');
      const rels = parseXml(xmlText(files['xl/_rels/workbook.xml.rels']));
      const rel = [...rels.getElementsByTagName('Relationship')].find(r => r.getAttribute('Id') === relId);
      const target = rel ? rel.getAttribute('Target') : 'worksheets/sheet1.xml';
      return 'xl/' + target.replace(/^\\//, '').replace(/^xl\\//, '');
    }}
    function colIndex(cellRef) {{
      const letters = String(cellRef || '').match(/[A-Z]+/i)?.[0] || 'A';
      let n = 0;
      for (const ch of letters.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
      return n - 1;
    }}
    function readCell(c, shared) {{
      const type = c.getAttribute('t');
      if (type === 'inlineStr') return c.getElementsByTagName('t')[0]?.textContent || '';
      const v = c.getElementsByTagName('v')[0]?.textContent ?? '';
      if (type === 's') return shared[Number(v)] || '';
      if (type === 'b') return v === '1';
      if (v !== '' && /^-?\\d+(?:\\.\\d+)?$/.test(v)) return Number(v);
      return v;
    }}
    function worksheetRows(files, sheetPath, shared) {{
      const doc = parseXml(xmlText(files[sheetPath]));
      return [...doc.getElementsByTagName('row')].map(row => {{
        const values = [];
        for (const c of row.getElementsByTagName('c')) {{
          values[colIndex(c.getAttribute('r'))] = readCell(c, shared);
        }}
        return values;
      }});
    }}
    function statusComplete(status, dirGr) {{
      return cleanText(status).toUpperCase() === 'DLV' || cleanText(dirGr).toLowerCase() === 'complete';
    }}
    function finalQty(attribute, matp) {{
      const a = String(attribute || '').trim().toUpperCase();
      if (a.endsWith('-BZ')) return matp.auto ?? '';
      if (a.endsWith('-CT')) return matp.cutting ?? '';
      return '';
    }}
    function rowFromExport(values) {{
      const line = cleanText(values[0]);
      if (!line) return null;
      const item = cleanText(values[6]);
      const route = DATA.routing[item] || {{}};
      const prodDt = parseDateTime(values[9], values[10]);
      const finishDt = parseDateTime(values[51] || values[52], values[53] || values[45]);
      const lead = prodDt && finishDt ? Math.round(((prodDt - finishDt) / 3600000) * 100) / 100 : '';
      const assyQty = cleanNumber(values[12]);
      const status = cleanText(values[8]);
      const complete = statusComplete(status, values[40]);
      const progressValue = complete ? assyQty : '';
      const code = cleanText(values[3]);
      const model = cleanText(values[4]);
      const speed = DATA.speedByModel[`${{line}}|${{code}}|${{model}}`] ?? DATA.speedByLine[line] ?? '';
      const month = monthDisplay(values[2]);
      const assyOrderNo = cleanText(values[13]);
      const head = `${{line}}|${{month}}|${{cleanNumber(values[1])}}`;
      const matp = (DATA.progressMat && DATA.progressMat[`${{head}}|${{item}}`]) || {{}};
      const hpVal = DATA.progressLot ? DATA.progressLot[`${{head}}|${{assyOrderNo}}`] : undefined;
      const fin = finalQty(route.attribute2, matp);
      return {{
        line,
        seq: cleanNumber(values[1]),
        month,
        code,
        model,
        finished: '',
        orderQty: cleanNumber(values[5]),
        item,
        description: route.description || cleanText(values[7]),
        attribute: route.attribute2 || '',
        attribute1: route.attribute1 || '',
        productionOrder: cleanText(values[11]),
        speed,
        prodDate: normalizeDate(values[9]),
        prodTime: normalizeTime(values[10]),
        prodSort: prodDt ? prodDt.toISOString() : '',
        assyOrder: assyQty,
        assyOrderNo,
        status,
        remark: [cleanText(values[23]), cleanText(values[24]), cleanText(values[50])].filter(Boolean).join(' '),
        hp: hpVal === undefined ? '' : hpVal,
        fp: matp.fp ?? '',
        exp: progressValue,
        auto: matp.auto ?? '',
        cutting: matp.cutting ?? '',
        fg: fin,
        unit: route.unit || '',
        stockFg: fin,
        subcooler: fin,
        lead,
        leadRemark: cleanText(values[41]) || cleanText(values[44]),
        ts: route.ts || '',
        operation: route.operation || ''
      }};
    }}
    async function rowsFromXlsx(file) {{
      const files = await unzipXlsx(await file.arrayBuffer());
      const shared = readSharedStrings(files);
      const rows = worksheetRows(files, firstWorksheetPath(files), shared);
      if (!rows.length) throw new Error('No rows found in file.');
      const header = rows[0].map(cleanText);
      const expected = ['Line','Sequence','Production Month','Assy Material','Assy Material Description','Order Quantity','Material'];
      const ok = expected.every((name, idx) => header[idx] === name);
      if (!ok) throw new Error('File format does not match EXPORT raw data.');
      return rows.slice(1).map(rowFromExport).filter(Boolean).sort((a,b) =>
        [a.line, a.prodSort || '9999', a.seq, a.assyOrderNo, a.productionOrder, a.item].join('|')
          .localeCompare([b.line, b.prodSort || '9999', b.seq, b.assyOrderNo, b.productionOrder, b.item].join('|'), 'th', {{numeric:true}})
      );
    }}
    async function importRawData(event) {{
      const file = event.target.files?.[0];
      if (!file) return;
      els.importStatus.textContent = `Importing: ${{file.name}}`;
      try {{
        currentRows = await rowsFromXlsx(file);
        indexRows(currentRows);
        saveSelected();
        needSort = true;
        leadDirty = true;
        progressCleared = false;
        localStorage.removeItem('dailyFollowProgCleared');
        els.line.value = 'ALL';
        els.month.value = 'ALL';
        els.status.value = 'ALL';
        els.search.value = '';
        refreshFilters();
        recompute(true);
        els.importStatus.textContent = `Loaded: ${{file.name}} (${{currentRows.length.toLocaleString()}} rows)`;
      }} catch (error) {{
        console.error(error);
        els.importStatus.textContent = `Import failed: ${{error.message}}`;
      }} finally {{
        event.target.value = '';
      }}
    }}
    function seqKey(value) {{
      const n = cleanNumber(value);
      return n === '' ? '' : String(n);
    }}
    function aggregateProgress(rows) {{
      const header = rows[0].map(cleanText);
      const ix = {{}};
      header.forEach((h, i) => {{ if (!(h in ix)) ix[h] = i; }});
      const need = ['Production Line','Production Month','Sequence','Material','Assembly Order','Operation Short Text','Posted Quantity'];
      for (const name of need) {{ if (ix[name] === undefined) throw new Error('Missing column: ' + name); }}
      const opField = {{ Insert: 'fp', Brazing: 'auto', Cutting: 'cutting' }};
      const mat = {{}}; const lot = {{}};
      for (let r = 1; r < rows.length; r++) {{
        const row = rows[r];
        const line = cleanText(row[ix['Production Line']]);
        if (!line) continue;
        const head = `${{line}}|${{monthDisplay(row[ix['Production Month']])}}|${{seqKey(row[ix['Sequence']])}}`;
        const op = cleanText(row[ix['Operation Short Text']]);
        const q = Number(row[ix['Posted Quantity']]);
        const qty = Number.isFinite(q) ? q : 0;
        if (op === 'H/P bender') {{
          const k = `${{head}}|${{cleanText(row[ix['Assembly Order']])}}`;
          lot[k] = (lot[k] || 0) + qty;
        }} else if (opField[op]) {{
          const k = `${{head}}|${{cleanText(row[ix['Material']])}}`;
          if (!mat[k]) mat[k] = {{}};
          const f = opField[op];
          mat[k][f] = (mat[k][f] || 0) + qty;
        }}
      }}
      const rnd = v => Number.isInteger(v) ? v : Math.round(v * 100) / 100;
      for (const k in lot) lot[k] = rnd(lot[k]);
      for (const k in mat) for (const f in mat[k]) mat[k][f] = rnd(mat[k][f]);
      return {{ mat, lot }};
    }}
    function applyProgressToRows() {{
      for (const row of currentRows) {{
        const head = `${{row.line}}|${{row.month}}|${{seqKey(row.seq)}}`;
        const matp = (DATA.progressMat && DATA.progressMat[`${{head}}|${{row.item}}`]) || {{}};
        const hpVal = DATA.progressLot ? DATA.progressLot[`${{head}}|${{row.assyOrderNo}}`] : undefined;
        const fin = finalQty(row.attribute, matp);
        row.hp = hpVal === undefined ? '' : hpVal;
        row.fp = matp.fp ?? '';
        row.auto = matp.auto ?? '';
        row.cutting = matp.cutting ?? '';
        row.fg = fin;
        row.stockFg = fin;
        row.subcooler = fin;
      }}
    }}
    async function importStockData(event) {{
      const file = event.target.files?.[0];
      if (!file) return;
      els.importStatus.textContent = `Updating stock: ${{file.name}}`;
      try {{
        const files = await unzipXlsx(await file.arrayBuffer());
        const shared = readSharedStrings(files);
        const rows = worksheetRows(files, firstWorksheetPath(files), shared);
        if (!rows.length) throw new Error('No rows found in file.');
        const agg = aggregateProgress(rows);
        DATA.progressMat = agg.mat;
        DATA.progressLot = agg.lot;
        applyProgressToRows();
        progressCleared = false;
        localStorage.removeItem('dailyFollowProgCleared');
        leadDirty = true;
        recompute(true);
        const nMat = Object.keys(agg.mat).length, nLot = Object.keys(agg.lot).length;
        els.importStatus.textContent = `Stock updated: ${{file.name}} (${{nMat.toLocaleString()}} mat / ${{nLot.toLocaleString()}} lot)`;
      }} catch (error) {{
        console.error(error);
        els.importStatus.textContent = `Stock update failed: ${{error.message}}`;
      }} finally {{
        event.target.value = '';
      }}
    }}
    async function runZpp0059() {{
      // Ask the local companion server (serve_daily_follow.py) to drive SAP via
      // Script_0059.vbs, then apply the freshly-exported ZPP0059 progress.
      if (location.protocol === 'file:') {{
        els.importStatus.textContent = 'ZPP0059 needs the local server. Double-click Start_Daily_Follow.bat, then open the page it gives you.';
        return;
      }}
      const btn = els.zpp0059Btn;
      btn.disabled = true;
      btn.classList.add('busy');
      els.importStatus.textContent = 'ZPP0059: pulling data from SAP… (do not touch the SAP window)';
      try {{
        const resp = await fetch('api/run-zpp0059', {{ method: 'POST' }});
        const res = await resp.json().catch(() => ({{}}));
        if (!resp.ok || !res.ok) throw new Error(res.error || ('HTTP ' + resp.status));
        DATA.progressMat = res.mat || {{}};
        DATA.progressLot = res.lot || {{}};
        applyProgressToRows();
        progressCleared = false;
        localStorage.removeItem('dailyFollowProgCleared');
        leadDirty = true;
        recompute(true);
        const nMat = Object.keys(DATA.progressMat).length, nLot = Object.keys(DATA.progressLot).length;
        const fname = res.file ? ` from ${{res.file}}` : '';
        const st = res.stats || {{}};
        const dbInfo = st.new_rows !== undefined
          ? ` [DB +${{st.new_rows}} new / ${{st.skipped_rows}} dup / ${{(st.total_rows||0).toLocaleString()}} total]`
          : '';
        els.importStatus.textContent = `ZPP0059 updated${{fname}}${{dbInfo}} (${{nMat.toLocaleString()}} mat / ${{nLot.toLocaleString()}} lot)`;
      }} catch (error) {{
        console.error(error);
        els.importStatus.textContent = `ZPP0059 failed: ${{error.message}}`;
      }} finally {{
        btn.disabled = false;
        btn.classList.remove('busy');
      }}
    }}
    function clearData() {{
      currentRows = [];
      saved = {{}};
      needSort = true;
      leadDirty = true;
      progressCleared = false;
      byId = new Map();
      selected.clear();
      saveSelected();
      localStorage.removeItem('dailyFollowProgress');
      localStorage.removeItem('dailyFollowProgCleared');
      els.line.value = 'ALL';
      els.month.value = 'ALL';
      els.status.value = 'ALL';
      els.search.value = '';
      refreshFilters();
      recompute(true);
      els.importStatus.textContent = 'Data reset. Ready to update progress.';
    }}
    function clearProgress() {{
      if (!confirm('Clear all values in H/P, FP, EXP, Auto, Cutting, FG, Assy and Subcooler columns?')) return;
      for (const k of Object.keys(saved)) {{
        const field = k.slice(k.lastIndexOf('|') + 1);
        if (progressFields.has(field)) delete saved[k];
      }}
      okSet.clear();
      progressCleared = true;
      localStorage.setItem('dailyFollowProgress', JSON.stringify(saved));
      localStorage.setItem('dailyFollowOk', JSON.stringify([]));
      localStorage.setItem('dailyFollowProgCleared', '1');
      leadDirty = true;
      recompute(true);
    }}
    function visibleColumns() {{
      const visible = columns.filter(([field]) => !hidden.has(field));
      return visible.length ? visible : columns;
    }}
    function renderHeader(activeColumns) {{
      els.header.replaceChildren();
      activeColumns.forEach(([field, title]) => {{
        const th = document.createElement('th');
        th.dataset.field = field;
        if (field === '_sel') {{
          th.className = 'selcol';
          th.title = 'Select / clear all rows in view';
          const chk = document.createElement('input');
          chk.type = 'checkbox';
          chk.className = 'selall';
          chk.checked = allFilteredSelected();
          chk.indeterminate = !chk.checked && filtered.some(r => selected.has(r._id));
          chk.addEventListener('click', (e) => e.stopPropagation());
          chk.addEventListener('change', () => {{
            if (chk.checked) for (const r of filtered) {{ if (!isHairPin(r)) selected.add(r._id); }}
            else for (const r of filtered) selected.delete(r._id);
            saveSelected();
            renderWindow(true);
            updateSimReadout();
          }});
          els.selAll = chk;
          th.appendChild(chk);
          els.header.appendChild(th);
          return;
        }}
        th.appendChild(document.createTextNode(title));
        if (headerClass[field]) th.classList.add(headerClass[field]);
        th.title = '\\u00D7 to hide';
        const hide = document.createElement('span');
        hide.className = 'th-hide';
        hide.textContent = '\\u00D7';
        hide.title = 'Hide this column';
        hide.addEventListener('click', (e) => {{ e.stopPropagation(); hideColumn(field); }});
        th.appendChild(hide);
        els.header.appendChild(th);
      }});
    }}
    function recompute(resetScroll) {{
      if (needSort) {{ sortedAll = sortRows(currentRows.slice()); needSort = false; }}
      const q = els.search.value.trim().toLowerCase();
      const filtering = q || els.line.value !== 'ALL' || els.month.value !== 'ALL'
        || els.status.value !== 'ALL' || hideSkip || hideHairPin;
      filtered = filtering ? sortedAll.filter(r => passes(r, q)) : sortedAll;
      if (resetScroll && els.wrap) els.wrap.scrollTop = 0;
      render();
    }}
    function allFilteredSelected() {{
      let any = false;
      for (const r of filtered) {{
        if (isHairPin(r)) continue;
        any = true;
        if (!selected.has(r._id)) return false;
      }}
      return any;
    }}
    function saveSelected() {{
      localStorage.setItem('dailyFollowSel', JSON.stringify([...selected]));
    }}
    function restoreSelected() {{
      selected.clear();
      let ids = [];
      try {{ ids = JSON.parse(localStorage.getItem('dailyFollowSel') || '[]'); }} catch (e) {{ ids = []; }}
      for (const id of ids) {{ if (byId.has(id)) selected.add(id); }}
    }}
    function updateSimReadout() {{
      let rate = 0, qty = 0, n = 0;
      for (const id of selected) {{
        const row = byId.get(id);
        if (!row || isHairPin(row)) continue;
        rate += num(row.ts) * num(row.assyOrder);
        qty += num(row.assyOrder);
        n += 1;
      }}
      els.simBadge.textContent = (Math.round(rate * 10) / 10).toLocaleString();
      els.simSub.textContent = n.toLocaleString() + ' rows \\u00B7 ' + qty.toLocaleString() + ' pcs';
      const wh = num(els.workHour.value);
      const prod = rate > 0 ? wh / rate : 0;
      els.prodBadge.textContent = (Math.round(prod * 100) / 100).toLocaleString();
      els.prodBadge.classList.toggle('danger', prod > 0.89);
    }}
    function render() {{
      activeCols = visibleColumns();
      renderColgroup(activeCols);
      renderHeader(activeCols);
      renderHiddenBar();
      winStart = winEnd = -1;
      renderWindow(true);
      updateSimReadout();
      renderLeadStrip();
    }}
    function renderColgroup(cols) {{
      const frag = document.createDocumentFragment();
      for (const col of cols) {{
        const c = document.createElement('col');
        if (col[3]) c.style.width = col[3] + 'px';
        frag.appendChild(c);
      }}
      els.colgroup.replaceChildren(frag);
    }}
    function spacerRow(span, height) {{
      const tr = document.createElement('tr');
      tr.className = 'spacer';
      const td = document.createElement('td');
      td.colSpan = span;
      td.style.height = height + 'px';
      tr.appendChild(td);
      return tr;
    }}
    function buildRow(row) {{
      const tr = document.createElement('tr');
      tr.__row = row;
      const locked = isHairPin(row);
      const leadVal = computeLead(row);
      tr.className = locked ? 'row-locked' : (leadVal === '' ? 'row-missing' : 'row-ready');
      let ci = -1;
      for (const [field, title, klass] of activeCols) {{
        ci++;
        const td = document.createElement('td');
        td.className = klass || '';
        td.dataset.ci = ci;
        if (field === '_sel') {{
          if (!locked) {{
            const chk = document.createElement('input');
            chk.type = 'checkbox';
            chk.className = 'rowchk';
            chk.checked = selected.has(row._id);
            td.appendChild(chk);
          }}
          tr.appendChild(td);
          continue;
        }}
        let value;
        if (field === 'unitRate') {{ value = computeUnitRate(row); td.dataset.field = 'unitRate'; }}
        else if (field === 'lead') {{ value = leadVal; td.dataset.field = 'lead'; }}
        else if (field === 'scLead') {{ value = computeScLead(row); td.dataset.field = 'scLead'; }}
        else value = valueOf(row, field);
        // -BZ records end at Brazing, so the Cutting column is blocked (no value, not editable).
        const cutBlocked = field === 'cutting' && attrEndsBz(row);
        if (cutBlocked) {{ value = ''; td.classList.add('blocked'); }}
        if (field === 'status' && fmt(value)) {{
          const badge = document.createElement('span');
          const s = String(value).trim().toUpperCase();
          badge.className = 'badge ' + (s === 'DLV' || s === 'COMPLETE' ? 'ok' : 'info');
          badge.textContent = value;
          td.appendChild(badge);
        }} else {{
          td.textContent = fmt(value);
          if (!td.textContent && !cutBlocked) td.classList.add('blank');
        }}
        if ((field === 'stockFg' || field === 'subcooler') && okSet.has(keyOf(row, field))) td.classList.add('okcell');
        if (editable.has(field) && !locked && !cutBlocked) {{
          td.contentEditable = 'true';
          td.dataset.field = field;
          td.classList.add('editable');
          td.title = 'Edit value, saved locally in this browser';
          const ao = row.assyOrder;
          if (field !== 'leadRemark' && value !== '' && value !== null && value !== undefined
              && ao !== '' && ao !== null && ao !== undefined
              && Number(value) !== Number(ao)) td.classList.add('red');
        }}
        tr.appendChild(td);
      }}
      return tr;
    }}
    function renderWindow(force) {{
      const total = filtered.length;
      const viewH = els.wrap ? els.wrap.clientHeight : 800;
      const scrollTop = els.wrap ? els.wrap.scrollTop : 0;
      const buffer = 12;
      const start = Math.max(0, Math.floor(scrollTop / ROW_H) - buffer);
      const visible = Math.ceil(viewH / ROW_H) + buffer * 2;
      const end = Math.min(total, start + visible);
      if (!force && start === winStart && end === winEnd) return;
      winStart = start; winEnd = end;
      const span = activeCols.length || 1;
      const frag = document.createDocumentFragment();
      if (start > 0) frag.appendChild(spacerRow(span, start * ROW_H));
      for (let i = start; i < end; i++) {{
        const tr = buildRow(filtered[i]);
        tr.dataset.idx = i;
        frag.appendChild(tr);
      }}
      if (end < total) frag.appendChild(spacerRow(span, (total - end) * ROW_H));
      els.body.replaceChildren(frag);
      if (!measuredRowH) {{
        const sample = els.body.querySelector('tr:not(.spacer)');
        if (sample) {{
          const h = sample.getBoundingClientRect().height;
          measuredRowH = true;
          if (h > 0 && Math.abs(h - ROW_H) > 0.5) {{ ROW_H = h; winStart = winEnd = -1; renderWindow(true); }}
        }}
      }}
      highlightPending();
    }}
    function onScroll() {{
      if (scrollTicking) return;
      scrollTicking = true;
      requestAnimationFrame(() => {{ scrollTicking = false; renderWindow(false); positionFillHandle(); }});
    }}
    function positionFillHandle() {{
      const td = fillSource && fillSource.td;
      if (!td || !document.body.contains(td)) {{ els.fillHandle.style.display = 'none'; return; }}
      const r = td.getBoundingClientRect();
      const wrapR = els.wrap.getBoundingClientRect();
      const headH = els.header.getBoundingClientRect().height || 26;
      if (r.bottom <= wrapR.top + headH || r.top >= wrapR.bottom
          || r.right <= wrapR.left || r.left >= wrapR.right) {{
        els.fillHandle.style.display = 'none';
        return;
      }}
      els.fillHandle.style.display = 'block';
      els.fillHandle.style.left = (r.right - 5) + 'px';
      els.fillHandle.style.top = (r.bottom - 5) + 'px';
    }}
    function updateFillTargetFromPoint() {{
      const el = document.elementFromPoint(lastMouse.x, lastMouse.y);
      const td = el && el.closest && el.closest('td');
      if (!td) return;
      const tr = td.closest('tr');
      if (!tr || tr.classList.contains('spacer') || tr.dataset.idx === undefined) return;
      const tRow = Number(tr.dataset.idx);
      const tCi = td.dataset.ci === undefined ? fillSource.ci : Number(td.dataset.ci);
      fillRange = [
        Math.min(fillSource.idx, tRow), Math.max(fillSource.idx, tRow),
        Math.min(fillSource.ci, tCi), Math.max(fillSource.ci, tCi)
      ];
    }}
    function highlightFillRange() {{
      els.body.querySelectorAll('td.fill-preview').forEach(td => td.classList.remove('fill-preview'));
      if (!fillRange || !fillSource) return;
      const [r0, r1, c0, c1] = fillRange;
      for (const tr of els.body.children) {{
        if (tr.classList.contains('spacer') || tr.dataset.idx === undefined) continue;
        const idx = Number(tr.dataset.idx);
        if (idx < r0 || idx > r1) continue;
        for (let c = c0; c <= c1; c++) {{
          const field = activeCols[c] && activeCols[c][0];
          if (!field || !editable.has(field)) continue;
          const cell = tr.querySelector('td[data-field="' + field + '"]');
          if (cell) cell.classList.add('fill-preview');
        }}
      }}
    }}
    function fillAutoScroll() {{
      if (!fillDragging) return;
      const wrapR = els.wrap.getBoundingClientRect();
      const headH = els.header.getBoundingClientRect().height || 26;
      let dy = 0, dx = 0;
      if (lastMouse.y > wrapR.bottom - 26) dy = 22;
      else if (lastMouse.y < wrapR.top + headH + 26) dy = -22;
      if (lastMouse.x > wrapR.right - 32) dx = 26;
      else if (lastMouse.x < wrapR.left + 32) dx = -26;
      if (dy || dx) {{
        if (dy) els.wrap.scrollTop += dy;
        if (dx) els.wrap.scrollLeft += dx;
        renderWindow(false);
        updateFillTargetFromPoint();
        highlightFillRange();
      }}
    }}
    function selChanges(rect, value) {{
      const [r0, r1, c0, c1] = rect;
      const out = [];
      for (let i = r0; i <= r1; i++) {{
        const row = filtered[i];
        if (!row || isHairPin(row)) continue;
        for (let c = c0; c <= c1; c++) {{
          const field = activeCols[c] && activeCols[c][0];
          if (field && editable.has(field)) out.push({{ key: keyOf(row, field), value: value }});
        }}
      }}
      return out;
    }}
    function applyChanges(changes) {{
      const undoEntry = [];
      for (const ch of changes) {{
        const had = Object.prototype.hasOwnProperty.call(saved, ch.key);
        const prev = saved[ch.key];
        if (had && prev === ch.value) continue;
        undoEntry.push({{ key: ch.key, had: had, prev: prev }});
        saved[ch.key] = ch.value;
      }}
      if (undoEntry.length) {{
        undoStack.push(undoEntry);
        if (undoStack.length > 200) undoStack.shift();
        localStorage.setItem('dailyFollowProgress', JSON.stringify(saved));
        leadDirty = true;
      }}
      return undoEntry.length > 0;
    }}
    function undo() {{
      const entry = undoStack.pop();
      if (!entry) return;
      for (const e of entry) {{
        if (e.had) saved[e.key] = e.prev;
        else delete saved[e.key];
      }}
      localStorage.setItem('dailyFollowProgress', JSON.stringify(saved));
      leadDirty = true;
      renderWindow(true);
      renderLeadStrip();
    }}
    function highlightPending() {{
      els.body.querySelectorAll('td.fill-sel').forEach(td => td.classList.remove('fill-sel'));
      if (!pendingSel) return;
      const [r0, r1, c0, c1] = pendingSel.rect;
      for (const tr of els.body.children) {{
        if (tr.classList.contains('spacer') || tr.dataset.idx === undefined) continue;
        const idx = Number(tr.dataset.idx);
        if (idx < r0 || idx > r1) continue;
        for (let c = c0; c <= c1; c++) {{
          const field = activeCols[c] && activeCols[c][0];
          if (!field || !editable.has(field)) continue;
          const cell = tr.querySelector('td[data-field="' + field + '"]');
          if (cell) cell.classList.add('fill-sel');
        }}
      }}
    }}
    function commitPendingFill() {{
      if (!pendingSel) return;
      applyChanges(selChanges(pendingSel.rect, pendingSel.value));
      pendingSel = null;
      renderWindow(true);
      renderLeadStrip();
    }}
    function clearPendingSel() {{
      if (!pendingSel) return;
      applyChanges(selChanges(pendingSel.rect, ''));
      pendingSel = null;
      renderWindow(true);
      renderLeadStrip();
    }}
    function cancelPendingSel() {{
      pendingSel = null;
      els.body.querySelectorAll('td.fill-sel').forEach(td => td.classList.remove('fill-sel'));
    }}
    function round1(n) {{ return Math.round(n * 10) / 10; }}
    function ltColor(h) {{
      if (h < 4) return '#b91c1c';
      if (h <= 10) return '#92620a';
      if (h <= 20) return '#15643a';
      if (h <= 40) return '#1a3a9e';
      return '#4c1d95';
    }}
    function ltBg(h) {{
      if (h < 4) return '#fca5a5';
      if (h <= 10) return '#fcd34d';
      if (h <= 20) return '#6ee7a0';
      if (h <= 40) return '#93c5fd';
      return '#d8b4fe';
    }}
    function defaultAssyTarget(line) {{ return 4; }}
    function defaultScTarget(line) {{
      const L = String(line).toUpperCase();
      if (L === 'A' || L === 'B') return 4;
      if (L === 'G' || L === 'T2') return 3;
      if (L === 'T' || L === 'H') return 2;
      return 4;
    }}
    function assyTarget(line) {{
      const v = targets.assy[line];
      return (v === undefined || v === null || v === '') ? defaultAssyTarget(line) : Number(v);
    }}
    function scTarget(line) {{
      const v = targets.sc[line];
      return (v === undefined || v === null || v === '') ? defaultScTarget(line) : Number(v);
    }}
    function assyColor(h, line) {{
      const t = assyTarget(line);
      if (h < t) return '#b91c1c';
      if (h <= 10) return '#92620a';
      if (h <= 20) return '#15643a';
      if (h <= 40) return '#1a3a9e';
      return '#4c1d95';
    }}
    function assyBg(h, line) {{
      const t = assyTarget(line);
      if (h < t) return '#fca5a5';
      if (h <= 10) return '#fcd34d';
      if (h <= 20) return '#6ee7a0';
      if (h <= 40) return '#93c5fd';
      return '#d8b4fe';
    }}
    function scColor(h, line) {{
      const t = scTarget(line);
      if (h <= t) return '#b91c1c';
      if (h <= 10) return '#92620a';
      if (h <= 20) return '#15643a';
      if (h <= 40) return '#1a3a9e';
      return '#4c1d95';
    }}
    function scBg(h, line) {{
      const t = scTarget(line);
      if (h <= t) return '#fca5a5';
      if (h <= 10) return '#fcd34d';
      if (h <= 20) return '#6ee7a0';
      if (h <= 40) return '#93c5fd';
      return '#d8b4fe';
    }}
    function dashStat(label, value, cls) {{
      const box = document.createElement('div');
      box.className = 'dash-stat' + (cls ? ' ' + cls : '');
      const l = document.createElement('div'); l.className = 'l'; l.textContent = label;
      const v = document.createElement('div'); v.className = 'v'; v.textContent = value;
      box.append(l, v);
      return box;
    }}
    function computeLineSummary() {{
      if (needSort) {{ sortedAll = sortRows(currentRows.slice()); needSort = false; }}
      ensureLeadGroups();
      const byLine = new Map();
      const stoppedA = new Set();
      const stoppedS = new Set();
      const countedA = new Set();
      const countedS = new Set();
      for (const row of sortedAll) {{
        const line = row.line || '-';
        if (/^SRV/i.test(line)) continue;
        let a = byLine.get(line);
        if (!a) {{ a = {{ line: line, lead: 0, scLead: 0, pcs: 0, orders: 0, from: '', to: '', rows: [] }}; byLine.set(line, a); }}
        if (isHairPin(row)) continue;
        const gk = groupKey(row);
        if (!stoppedA.has(line) && !assyOk(row)) {{
          const lead = computeLead(row);
          if (lead === '') {{ stoppedA.add(line); }}
          else if (!countedA.has(gk)) {{
            // One leadtime per seq group (based on its min stock), not per record.
            countedA.add(gk);
            const g = leadGroups.get(gk);
            const pcs = g && g.stockMin !== null ? g.stockMin : num(valueOf(row, 'stockFg'));
            a.lead += lead;
            a.pcs += pcs;
            a.orders += 1;
            if (!a.from) a.from = row.prodDate;
            a.to = row.prodDate;
            a.rows.push({{ d: row.prodDate, item: row.item, stock: pcs, lead: lead, sc: computeScLead(row) }});
          }}
        }}
        if (!stoppedS.has(line) && !scOk(row)) {{
          const sc = computeScLead(row);
          if (sc === '') {{ stoppedS.add(line); }}
          else if (!countedS.has(gk)) {{ countedS.add(gk); a.scLead += sc; }}
        }}
      }}
      let gLead = 0, gScLead = 0, gPcs = 0, gOrders = 0;
      for (const a of byLine.values()) {{ gLead += a.lead; gScLead += a.scLead; gPcs += a.pcs; gOrders += a.orders; }}
      const lines = [...byLine.values()].sort((x, y) => String(x.line).localeCompare(String(y.line), 'en', {{ numeric: true }}));
      const maxLead = Math.max(1, ...lines.map(l => Math.max(l.lead, l.scLead)));
      return {{ lines, gLead, gScLead, gPcs, gOrders, maxLead }};
    }}
    function renderLeadStrip() {{
      if (!els.leadStrip) return;
      const {{ lines }} = computeLineSummary();
      els.leadStrip.replaceChildren();
      if (!lines.length) {{
        els.leadStrip.classList.add('empty');
        return;
      }}
      els.leadStrip.classList.remove('empty');
      for (const l of lines) {{
        const chip = document.createElement('div');
        chip.className = 'lchip';
        chip.title = 'Line ' + l.line + ' \\u00B7 ' + l.orders + ' ord \\u00B7 ' + l.pcs.toLocaleString() + ' pcs (open Dashboard for details)';
        const name = document.createElement('span');
        name.className = 'lchip-name';
        name.textContent = l.line;
        chip.appendChild(name);
        [
          {{ lab: 'ASSY \\u2265' + assyTarget(l.line), val: l.lead, bg: assyBg(l.lead, l.line), col: assyColor(l.lead, l.line) }},
          {{ lab: 'SC \\u2265' + scTarget(l.line), val: l.scLead, bg: scBg(l.scLead, l.line), col: scColor(l.scLead, l.line) }}
        ].forEach((mm) => {{
          const m = document.createElement('span');
          m.className = 'lmetric';
          m.style.background = mm.bg;
          m.style.borderColor = mm.col;
          const ml = document.createElement('span'); ml.className = 'lmetric-l'; ml.textContent = mm.lab;
          ml.style.color = mm.col;
          const mv = document.createElement('span'); mv.className = 'lmetric-v'; mv.textContent = round1(mm.val).toLocaleString();
          m.append(ml, mv);
          chip.appendChild(m);
        }});
        els.leadStrip.appendChild(chip);
      }}
    }}
    function renderDashboardBody() {{
      const {{ lines, gLead, gScLead, gPcs, gOrders, maxLead }} = computeLineSummary();
      els.dashBody.replaceChildren();
      const stats = document.createElement('div');
      stats.className = 'dash-stats';
      stats.append(
        dashStat('Total Assy LT (hr)', round1(gLead).toLocaleString(), 'assy'),
        dashStat('Total SC LT (hr)', round1(gScLead).toLocaleString(), 'sc'),
        dashStat('Lines', lines.length.toLocaleString()),
        dashStat('Orders', gOrders.toLocaleString())
      );
      els.dashBody.appendChild(stats);
      const legend = document.createElement('div');
      legend.className = 'dash-legend';
      const lgT = document.createElement('span'); lgT.className = 'dash-leg-title'; lgT.textContent = 'Target (Lead Time)';
      legend.appendChild(lgT);
      [['#e23b3b', '< 4 hr'], ['#e0a200', '4-10 hr'], ['#1f9254', '>10-20 hr'], ['#2f56c4', '>20-40 hr'], ['#7c3aed', '>40 hr']].forEach(b => {{
        const it = document.createElement('span'); it.className = 'dash-leg-item';
        const sw = document.createElement('span'); sw.className = 'dash-leg-sw'; sw.style.background = b[0];
        it.append(sw, document.createTextNode(b[1]));
        legend.appendChild(it);
      }});
      els.dashBody.appendChild(legend);
      if (!lines.length) {{
        const empty = document.createElement('div');
        empty.style.padding = '14px 2px';
        empty.style.color = 'var(--muted)';
        empty.textContent = 'No data loaded.';
        els.dashBody.appendChild(empty);
        return;
      }}
      function metricBox(cls, label, value, widthPct, sub, line) {{
        const box = document.createElement('div'); box.className = 'dmetric ' + cls;
        const lab = document.createElement('div'); lab.className = 'dmetric-label'; lab.textContent = label;
        const num = document.createElement('div'); num.className = 'dmetric-num';
        num.append(value.toLocaleString());
        const unit = document.createElement('span'); unit.textContent = ' hr'; num.appendChild(unit);
        const col = cls === 'sc' ? scColor(value, line) : assyColor(value, line);
        lab.style.color = col;
        num.style.color = '#111111';
        box.style.background = cls === 'sc' ? scBg(value, line) : assyBg(value, line);
        box.style.borderColor = col;
        box.append(lab, num);
        if (sub) {{ const s = document.createElement('div'); s.className = 'dmetric-sub'; s.textContent = sub; box.appendChild(s); }}
        return box;
      }}
      for (const l of lines) {{
        const card = document.createElement('div');
        card.className = 'dline';
        card.title = 'Click to show the rows summed for this line';
        const name = document.createElement('div'); name.className = 'dline-name';
        name.textContent = l.line;
        const assy = metricBox('assy', 'Assy LT \\u2265' + assyTarget(l.line) + 'hr', round1(l.lead), l.lead / maxLead * 100, l.orders + ' ord \\u00B7 ' + l.pcs.toLocaleString() + ' pcs', l.line);
        const sc = metricBox('sc', 'SC LT \\u2265' + scTarget(l.line) + 'hr', round1(l.scLead), l.scLead / maxLead * 100, '', l.line);
        card.append(name, assy, sc);
        const detail = document.createElement('div');
        detail.className = 'dash-detail';
        detail.hidden = true;
        for (const r of l.rows) {{
          const dr = document.createElement('div');
          dr.className = 'dash-detail-row';
          const c1 = document.createElement('span'); c1.textContent = r.d || '-';
          const c2 = document.createElement('span'); c2.textContent = r.item;
          const c3 = document.createElement('span'); c3.className = 'r'; c3.textContent = r.stock.toLocaleString() + ' pcs';
          const c4 = document.createElement('span'); c4.className = 'r'; c4.textContent = r.lead + ' hr';
          const c5 = document.createElement('span'); c5.className = 'r'; c5.textContent = (r.sc === '' ? '-' : r.sc + ' SC');
          dr.append(c1, c2, c3, c4, c5);
          detail.appendChild(dr);
        }}
        card.addEventListener('click', () => {{ detail.hidden = !detail.hidden; }});
        els.dashBody.append(card, detail);
      }}
    }}
    function openDashboard() {{ renderDashboardBody(); els.dashOverlay.hidden = false; }}
    function closeDashboard() {{ els.dashOverlay.hidden = true; }}
    function openDatabase() {{
      if (location.protocol === 'file:') {{
        alert('View Database needs the local server.\nDouble-click Start_Daily_Follow.bat and open the page it gives you.');
        return;
      }}
      els.dbOverlay.hidden = false;
      loadDatabase(els.dbSearch.value.trim());
    }}
    async function loadDatabase(q) {{
      els.dbMeta.textContent = 'Loading…';
      els.dbTableWrap.replaceChildren();
      try {{
        const url = 'api/db-data?limit=1000&q=' + encodeURIComponent(q || '');
        const resp = await fetch(url);
        const data = await resp.json();
        if (!resp.ok) throw new Error(data.error || ('HTTP ' + resp.status));
        if (!data.rows.length) {{
          const e = document.createElement('div');
          e.className = 'db-empty';
          e.textContent = q ? `No rows match "${{q}}".` : 'Database is empty. Press ZPP0059 to pull data from SAP.';
          els.dbTableWrap.appendChild(e);
          els.dbMeta.textContent = `0 of ${{(data.total||0).toLocaleString()}} rows`;
          return;
        }}
        const table = document.createElement('table');
        table.className = 'db-table';
        const thead = document.createElement('thead');
        const htr = document.createElement('tr');
        data.columns.forEach(c => {{
          const th = document.createElement('th'); th.textContent = c; htr.appendChild(th);
        }});
        thead.appendChild(htr); table.appendChild(thead);
        const tbody = document.createElement('tbody');
        for (const row of data.rows) {{
          const tr = document.createElement('tr');
          for (const cell of row) {{
            const td = document.createElement('td');
            td.textContent = cell === null ? '' : cell;
            tr.appendChild(td);
          }}
          tbody.appendChild(tr);
        }}
        table.appendChild(tbody);
        els.dbTableWrap.appendChild(table);
        const shown = data.shown.toLocaleString();
        const total = (data.total||0).toLocaleString();
        els.dbMeta.textContent = `Showing ${{shown}} of ${{total}} rows`
          + (data.shown < data.total ? ` (newest first, max ${{data.limit}})` : '');
      }} catch (error) {{
        console.error(error);
        els.dbMeta.textContent = '';
        const e = document.createElement('div');
        e.className = 'db-empty';
        e.textContent = 'Failed to load database: ' + error.message;
        els.dbTableWrap.appendChild(e);
      }}
    }}
    function targetLines() {{
      const {{ lines }} = computeLineSummary();
      const names = lines.map(l => l.line).filter(n => n && n !== '-');
      for (const def of ['A', 'B', 'G', 'H', 'T', 'T2']) {{ if (!names.includes(def)) names.push(def); }}
      return names.sort((a, b) => String(a).localeCompare(String(b), 'en', {{ numeric: true }}));
    }}
    function openTargetEditor() {{
      const grid = document.createElement('div');
      grid.className = 'target-grid';
      ['Line', 'Assy target (Hrs.)', 'SC target (Hrs.)'].forEach(t => {{
        const h = document.createElement('div'); h.className = 'th'; h.textContent = t; grid.appendChild(h);
      }});
      for (const line of targetLines()) {{
        const name = document.createElement('div'); name.className = 'tline'; name.textContent = line;
        const ai = document.createElement('input');
        ai.type = 'number'; ai.min = '0'; ai.step = '0.5'; ai.dataset.line = line; ai.dataset.kind = 'assy';
        ai.value = assyTarget(line);
        const si = document.createElement('input');
        si.type = 'number'; si.min = '0'; si.step = '0.5'; si.dataset.line = line; si.dataset.kind = 'sc';
        si.value = scTarget(line);
        grid.append(name, ai, si);
      }}
      els.targetBody.replaceChildren(grid);
      els.targetOverlay.hidden = false;
    }}
    function closeTargetEditor() {{ els.targetOverlay.hidden = true; }}
    function saveTargetEditor() {{
      const next = {{ assy: {{}}, sc: {{}} }};
      els.targetBody.querySelectorAll('input').forEach(inp => {{
        const v = inp.value.trim();
        if (v === '') return;
        next[inp.dataset.kind][inp.dataset.line] = Number(v);
      }});
      targets = next;
      saveTargets();
      renderLeadStrip();
      closeTargetEditor();
    }}
    function resetTargetEditor() {{
      targets = {{ assy: {{}}, sc: {{}} }};
      saveTargets();
      renderLeadStrip();
      openTargetEditor();
    }}
    try {{
      setup();
    }} catch (err) {{
      document.body.insertAdjacentHTML('afterbegin',
        `<div style="position:fixed;top:0;left:0;right:0;z-index:9999;background:#c00;color:#fff;padding:12px 16px;font:14px monospace;white-space:pre-wrap">
&#9888; JS Error during startup — กรุณาแจ้ง error นี้:<br>${{err.stack || err}}</div>`);
      console.error('setup() failed:', err);
    }}
  </script>
</body>
</html>
"""


def main():
    rows = build_rows()
    OUTPUT_FILE.write_text(render_html(rows), encoding="utf-8")
    print(f"Wrote {OUTPUT_FILE} with {len(rows)} rows")


if __name__ == "__main__":
    main()
