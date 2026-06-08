import io
import os
import csv
import json
import base64
from datetime import date, datetime
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from flask import Flask, render_template, request, jsonify, send_file, Response
from openpyxl import load_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB

# ═══════════════════════════════════════════════════════════════════════════════
#  STORAGE PATHS — two completely separate data dirs, one per dashboard
# ═══════════════════════════════════════════════════════════════════════════════
DATA_DIR   = Path(os.environ.get("DATA_DIR",   "/tmp/betaplan_data"))
BUILD_DIR  = Path(os.environ.get("BUILD_DIR",  "/tmp/buildstate_data"))
for d in (DATA_DIR, BUILD_DIR):
    d.mkdir(parents=True, exist_ok=True)

# Beta Plan paths
BP_SUMMARY = DATA_DIR / "latest_summary.json"
BP_ROWS    = DATA_DIR / "latest_rows.json"
BP_CHART   = DATA_DIR / "latest_chart.png"

# Build State paths
BS_SUMMARY = BUILD_DIR / "latest_summary.json"
BS_ROWS    = BUILD_DIR / "latest_rows.json"   # stored as CSV text
BS_CHART   = BUILD_DIR / "latest_chart.png"

# ═══════════════════════════════════════════════════════════════════════════════
#  BETA PLAN — constants + helpers (unchanged from v1)
# ═══════════════════════════════════════════════════════════════════════════════
WORKSTREAM_ORDER = [
    "Onboarding", "Setup", "Plan", "Prepare", "Execute",
    "LMDP App", "On-Shift", "Cross-temporal", "KTLO",
]
STATUS_COLORS = {
    "Not Started": "#9DBAE5",
    "In Progress":  "#F2C24E",
    "Blocked":      "#E26D6D",
    "Done":         "#7AB87A",
}


def parse_xlsx(file_bytes: bytes) -> list[dict]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb["Inventory"] if "Inventory" in wb.sheetnames else wb.active
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    required = {"ID", "Workstream", "Screen", "PctComplete", "Status"}
    missing = required - set(idx.keys())
    if missing:
        raise ValueError(f"Missing columns: {', '.join(sorted(missing))}")
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        rows.append({
            "id":         r[idx["ID"]],
            "workstream": str(r[idx["Workstream"]] or ""),
            "parent":     str(r[idx["Parent"]] or "") if "Parent" in idx else "",
            "screen":     str(r[idx["Screen"]] or ""),
            "owner":      str(r[idx["Owner"]] or "") if "Owner" in idx else "",
            "pct":        int(r[idx["PctComplete"]] or 0),
            "status":     str(r[idx["Status"]] or "Not Started"),
            "lastUpdate": str(r[idx["LastUpdate"]] or "") if "LastUpdate" in idx else "",
            "notes":      str(r[idx["Notes"]] or "") if "Notes" in idx else "",
        })
    return rows


def bp_make_chart(rows: list[dict], snapshot_date: str | None = None) -> bytes:
    snap = snapshot_date or date.today().isoformat()
    grouped = {ws: [] for ws in WORKSTREAM_ORDER}
    for r in rows:
        grouped.setdefault(r["workstream"], []).append(r)
    for ws in grouped:
        grouped[ws].sort(key=lambda r: (r["id"] if isinstance(r["id"], (int, float)) else 999))
    items = []
    for ws in WORKSTREAM_ORDER:
        lanes = grouped.get(ws, [])
        if not lanes:
            continue
        items.append({"kind": "header", "label": ws.upper()})
        for lane in lanes:
            label = lane["screen"]
            if lane["parent"] and lane["parent"] not in ("", ws, lane["workstream"]):
                label = f"  {lane['parent']} > {lane['screen']}"
            items.append({"kind": "lane", "label": label, "pct": lane["pct"],
                          "status": lane["status"], "owner": lane["owner"]})
    n = len(items)
    fig_h = max(8, n * 0.32)
    fig, ax = plt.subplots(figsize=(14, fig_h))
    y_positions = list(range(n))[::-1]
    for y, item in zip(y_positions, items):
        if item["kind"] == "header":
            ax.barh(y, 100, color="#222222", height=0.72, alpha=0.88)
            ax.text(50, y, item["label"], color="white", fontsize=10,
                    weight="bold", ha="center", va="center")
        else:
            color = STATUS_COLORS.get(item["status"], "#CCCCCC")
            ax.barh(y, item["pct"], color=color, height=0.62, edgecolor="#999999", linewidth=0.4)
            ax.barh(y, 100, color="none", edgecolor="#DDDDDD", height=0.62, linewidth=0.4, zorder=0)
            lx = item["pct"] + 1 if item["pct"] < 88 else item["pct"] - 6
            lc = "#333333" if item["pct"] < 88 else "white"
            ax.text(lx, y, f"{item['pct']}%", fontsize=8, va="center", color=lc, weight="bold")
            if item.get("owner"):
                ax.text(102, y, item["owner"], fontsize=7, va="center", color="#888888")
    ax.set_yticks(y_positions)
    ax.set_yticklabels([it["label"] for it in items], fontsize=8.5)
    ax.set_xlim(0, 115)
    ax.set_xlabel("Progress (%)", fontsize=10)
    ax.set_xticks([0, 25, 50, 75, 100])
    ax.set_title(f"Beta Plan 2026 — Progress Snapshot  ({snap})", fontsize=13, weight="bold", pad=14)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.spines["left"].set_color("#CCCCCC")
    ax.spines["bottom"].set_color("#CCCCCC")
    ax.tick_params(colors="#666666")
    ax.set_axisbelow(True)
    ax.grid(axis="x", color="#EEEEEE", linestyle="-", linewidth=0.5)
    legend_handles = [Patch(facecolor=c, edgecolor="#999999", label=s) for s, c in STATUS_COLORS.items()]
    ax.legend(handles=legend_handles, loc="lower right", fontsize=9, frameon=False)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=140, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def bp_summarise(rows: list[dict], uploaded_by: str = "") -> dict:
    total = len(rows)
    avg = round(sum(r["pct"] for r in rows) / total) if total else 0
    by_status: dict[str, int] = {}
    for r in rows:
        by_status[r["status"]] = by_status.get(r["status"], 0) + 1
    ws_summary = {}
    for ws in WORKSTREAM_ORDER:
        lanes = [r for r in rows if r["workstream"] == ws]
        if not lanes:
            continue
        ws_avg = round(sum(r["pct"] for r in lanes) / len(lanes))
        ws_summary[ws] = {"avg": ws_avg, "count": len(lanes), "by_status": {}}
        for r in lanes:
            s = r["status"]
            ws_summary[ws]["by_status"][s] = ws_summary[ws]["by_status"].get(s, 0) + 1
    return {
        "total": total, "avg": avg, "by_status": by_status, "workstreams": ws_summary,
        "snapshot_date": date.today().isoformat(),
        "updated_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
        "uploaded_by": uploaded_by,
    }


# ═══════════════════════════════════════════════════════════════════════════════
#  BUILD STATE — constants + helpers (mirrors build_state_chart.py exactly)
# ═══════════════════════════════════════════════════════════════════════════════
JIM_KEYS  = ["Jim_Mockup", "Jim_SP"]
TEAM_KEYS = ["Team_Design", "Team_FEdev", "Team_MW", "Team_FEwiring"]
SIZE_PTS = {"S": 1, "M": 2, "L": 3, "XL": 5}
SUB_PTS  = {"S": 1, "M": 2, "L": 3}

ENGINES = [
    ("Ledger",                  8, 90),
    ("Scheduling brain",        5, 80),
    ("Recommendation engine",   5, 40),
    ("Comms",                   5, 15),
    ("Asset system",            3, 90),
    ("LMDP analytics",          3, 75),
    ("Platform / KTLO",         3, 45),
]
RESERVE_PTS = 7


def _num(v):
    try:
        s = str(v).strip()
        return 0.0 if s in ("", "NA", "N/A") else float(s)
    except (TypeError, ValueError):
        return 0.0


def bs_component_pct(row: dict) -> tuple[float, float]:
    j_vals = [_num(row[k]) for k in JIM_KEYS if str(row.get(k, "")).strip() not in ("NA", "N/A", "")]
    t_vals = [_num(row[k]) for k in TEAM_KEYS if str(row.get(k, "")).strip() not in ("NA", "N/A", "")]
    jim  = sum(j_vals) / len(j_vals) if j_vals else 0.0
    team = sum(t_vals) / len(t_vals) if t_vals else 0.0
    return jim, team


def bs_roll_up(rows: list[dict]) -> dict:
    by_parent: dict[str, list] = {}
    for r in rows:
        if not r.get("Parent Module"):
            continue
        by_parent.setdefault(r["Parent Module"], []).append(r)
    
    out = {}
    for parent, p_rows in by_parent.items():
        by_child = {}
        for r in p_rows:
            by_child.setdefault(r.get("Child Module", ""), []).append(r)
            
        p_wsum = p_jacc = p_tacc = 0.0
        children_out = {}
        
        for child, c_rows in by_child.items():
            c_wsum = c_jacc = c_tacc = 0.0
            child_items = []
            for c in c_rows:
                sz = SIZE_PTS.get(str(c.get("ScreenSize", "M")).strip(), 2)
                sub = SUB_PTS.get(str(c.get("SubSize", "")).strip(), 1)
                w = sz * sub
                j, t = bs_component_pct(c)
                c_wsum += w; c_jacc += w * j; c_tacc += w * t
                child_items.append({
                    "label": c.get("Description", ""),
                    "weight": w,
                    "jim_pct": j,
                    "team_pct": t
                })
            
            p_wsum += c_wsum; p_jacc += c_jacc; p_tacc += c_tacc
            children_out[child] = {
                "weight": c_wsum,
                "jim_pct": c_jacc / c_wsum if c_wsum else 0,
                "team_pct": c_tacc / c_wsum if c_wsum else 0,
                "rows": child_items
            }
            
        out[parent] = {
            "size_pts": p_wsum,
            "jim_pct": p_jacc / p_wsum if p_wsum else 0,
            "team_pct": p_tacc / p_wsum if p_wsum else 0,
            "children": children_out
        }
    return out


def bs_make_chart(rows: list[dict]) -> bytes:
    screens = bs_roll_up(rows)
    items = []   # (label, indent, jd, td, pct_done, is_bar)
    
    sorted_parents = sorted(screens.items(), key=lambda kv: -(kv[1]["jim_pct"]*0.48 + kv[1]["team_pct"]*0.52))
    for mod, p_data in sorted_parents:
        jd = 0.48 * p_data["jim_pct"]
        td = 0.52 * p_data["team_pct"]
        items.append((mod, 0, jd, td, jd + td, True))
        
        for child, c_data in p_data["children"].items():
            if child:
                items.append((child, 1, 0, 0, 0, False))
            for r in c_data["rows"]:
                r_jd = 0.48 * r["jim_pct"]
                r_td = 0.52 * r["team_pct"]
                label = r["label"] if r["label"] else "Untitled"
                items.append((label, 2 if child else 1, r_jd, r_td, r_jd + r_td, True))
                
    items.append(("─── ENGINES ───", 0, 0, 0, 0, False))
    for name, pts, pct in ENGINES:
        items.append((name, 0, pct, 0, pct, True))
    items.append(("Reserve (unscoped)", 0, 0, 0, 0, True))

    n = len(items)
    fig, ax = plt.subplots(figsize=(13, max(8, n * 0.38)))
    y = list(range(n))[::-1]

    for yi, (label, indent, jd, td, pct_done, is_bar) in zip(y, items):
        if not is_bar:
            ax.axhline(yi, color="#DDDDDD", linewidth=0.6, zorder=0)
            continue
        ax.barh(yi, jd,           color="#2563EB", height=0.62, label="Jim done (DoR)")
        ax.barh(yi, td, left=jd,  color="#7AB87A", height=0.62, label="Team done (DoD)")
        remain = max(0, 100 - jd - td)
        ax.barh(yi, remain, left=jd + td, color="#E5E7EB", height=0.62, label="Remaining",
                edgecolor="#CCCCCC", linewidth=0.3)
        ax.text(100 + 1.5, yi, f"{pct_done:.0f}%", fontsize=8, va="center", color="#444444")

    ax.set_yticks(y)
    ytick_labels = []
    for label, indent, jd, td, pct_done, is_bar in items:
        ytick_labels.append(f"{'    ' * indent}{label}")
    ax.set_yticklabels(ytick_labels, fontsize=9)
    
    for i, tick_label in enumerate(ax.get_yticklabels()):
        item = items[i]
        if item[1] == 0:
            tick_label.set_weight("bold")
        if not item[5]:
            tick_label.set_color("#666666")
            
    ax.set_xlabel("Progress (%)", fontsize=10)
    ax.set_xlim(0, 105)

    tot_pts = sum(p["size_pts"] for p in screens.values()) + sum(e[1] for e in ENGINES) + RESERVE_PTS
    jd_sum = sum(p["size_pts"] * 0.48 * p["jim_pct"] / 100 for p in screens.values())
    td_sum = sum(p["size_pts"] * 0.52 * p["team_pct"] / 100 for p in screens.values())
    done_pts = jd_sum + td_sum + sum(e[1] * e[2] / 100 for e in ENGINES)
    overall_pct = done_pts / tot_pts * 100 if tot_pts else 0
    snap = date.today().isoformat()

    ax.set_title(
        f"Build State — {done_pts:.1f} / {tot_pts:.0f} pts ({overall_pct:.0f}%)  ·  {snap}\n"
        f"blue = Jim DoR  ·  green = Team DoD",
        fontsize=12, weight="bold", pad=12,
    )
    
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.spines["left"].set_color("#CCCCCC")
    ax.spines["bottom"].set_color("#CCCCCC")
    ax.tick_params(colors="#666666")
    ax.set_axisbelow(True)
    ax.grid(axis="x", color="#EEEEEE", linewidth=0.5)

    seen, handles, labels = set(), [], []
    for h, l in zip(*ax.get_legend_handles_labels()):
        if l not in seen:
            seen.add(l); handles.append(h); labels.append(l)
    ax.legend(handles=handles, labels=labels, loc="lower right", frameon=False, fontsize=9)

    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=140, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def bs_summarise(rows: list[dict], uploaded_by: str = "") -> dict:
    screens = bs_roll_up(rows)
    tot  = sum(p["size_pts"] for p in screens.values()) + sum(e[1] for e in ENGINES) + RESERVE_PTS
    jim_done  = sum(p["size_pts"] * 0.48 * p["jim_pct"] / 100 for p in screens.values())
    jim_total = sum(p["size_pts"] * 0.48 for p in screens.values())
    team_done  = sum(p["size_pts"] * 0.52 * p["team_pct"] / 100 for p in screens.values())
    team_total = sum(p["size_pts"] * 0.52 for p in screens.values())
    done = jim_done + team_done + sum(e[1] * e[2] / 100 for e in ENGINES)
    overall_pct = round(done / tot * 100) if tot else 0

    by_module = {
        mod: {
            "size_pts": p["size_pts"],
            "jim_pct":  round(p["jim_pct"]),
            "team_pct": round(p["team_pct"]),
            "overall_pct": round(p["jim_pct"]*0.48 + p["team_pct"]*0.52),
        }
        for mod, p in sorted(screens.items(), key=lambda kv: -(kv[1]["jim_pct"]*0.48 + kv[1]["team_pct"]*0.52))
    }
    return {
        "overall_pct": overall_pct,
        "total_pts":   round(tot, 1),
        "done_pts":    round(done, 1),
        "jim_pct":     round(jim_done / jim_total * 100) if jim_total else 0,
        "team_pct":    round(team_done / team_total * 100) if team_total else 0,
        "module_count": len(screens),
        "component_count": len(rows),
        "by_module": by_module,
        "snapshot_date": date.today().isoformat(),
        "updated_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
        "uploaded_by": uploaded_by,
    }


def bs_parse_csv(file_bytes: bytes) -> list[dict]:
    text = file_bytes.decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(text))
    required = {"Parent Module", "Child Module", "Description", "ScreenSize", "SubSize",
                 "Jim_Mockup", "Jim_SP", "Team_Design", "Team_FEdev", "Team_MW", "Team_FEwiring"}
    rows = list(reader)
    if not rows:
        raise ValueError("CSV is empty")
    missing = required - set(rows[0].keys())
    if missing:
        raise ValueError(f"Missing CSV columns: {', '.join(sorted(missing))}")
    return rows


def bs_parse_xlsx(file_bytes: bytes) -> list[dict]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb["Inventory"] if "Inventory" in wb.sheetnames else wb.active
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    return [dict(zip(headers, [str(c.value) if c.value is not None else "" for c in row]))
            for row in ws.iter_rows(min_row=2)
            if any(c.value for c in row)]


# ═══════════════════════════════════════════════════════════════════════════════
#  ROUTES — BETA PLAN  (all unchanged, prefix /)
# ═══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    has_data = BP_SUMMARY.exists()
    summary = chart_b64 = None
    if has_data:
        data = json.loads(BP_SUMMARY.read_text())
        chart_b64 = data.pop("chart_b64", None)
        summary = data
    return render_template("index.html", has_data=has_data, summary=summary, chart_b64=chart_b64)


@app.route("/latest-chart")
def latest_chart():
    if BP_CHART.exists():
        return send_file(BP_CHART, mimetype="image/png", download_name="Progress_Chart_latest.png")
    if not BP_ROWS.exists():
        return Response("No chart yet.", status=404, mimetype="text/plain")
    rows = json.loads(BP_ROWS.read_text())
    png  = bp_make_chart(rows)
    BP_CHART.write_bytes(png)
    return send_file(io.BytesIO(png), mimetype="image/png", download_name="Progress_Chart_latest.png")


@app.route("/latest-summary")
def latest_summary():
    if not BP_SUMMARY.exists():
        return jsonify({"error": "No data yet"}), 404
    data = json.loads(BP_SUMMARY.read_text())
    data.pop("chart_b64", None)
    return jsonify(data)


@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Please upload an .xlsx or .xls file"}), 400
    uploaded_by = request.form.get("uploaded_by", "").strip() or "Rishab"
    try:
        rows = parse_xlsx(f.read())
    except ValueError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        return jsonify({"error": f"Could not read file: {e}"}), 500
    try:
        png = bp_make_chart(rows)
    except Exception as e:
        return jsonify({"error": f"Chart generation failed: {e}"}), 500
    summary = bp_summarise(rows, uploaded_by)
    chart_b64 = base64.b64encode(png).decode()
    BP_SUMMARY.write_text(json.dumps({**summary, "chart_b64": chart_b64}))
    BP_ROWS.write_text(json.dumps(rows))
    BP_CHART.write_bytes(png)
    return jsonify({"ok": True, "summary": summary, "rows": rows})


@app.route("/chart", methods=["POST"])
def chart():
    body = request.get_json(silent=True) or {}
    rows = body.get("rows")
    if not rows:
        return jsonify({"error": "No data"}), 400
    try:
        png = bp_make_chart(rows)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    return send_file(io.BytesIO(png), mimetype="image/png",
                     download_name=f"Progress_Chart_{date.today().isoformat()}.png")


# ═══════════════════════════════════════════════════════════════════════════════
#  ROUTES — BUILD STATE  (all under /build prefix)
# ═══════════════════════════════════════════════════════════════════════════════

@app.route("/build")
def build_index():
    has_data = BS_SUMMARY.exists()
    summary = chart_b64 = None
    if has_data:
        data = json.loads(BS_SUMMARY.read_text())
        chart_b64 = data.pop("chart_b64", None)
        summary = data
    return render_template("build.html", has_data=has_data, summary=summary, chart_b64=chart_b64)


@app.route("/build/latest-chart")
def build_latest_chart():
    if BS_CHART.exists():
        return send_file(BS_CHART, mimetype="image/png", download_name="Build_State_Chart_latest.png")
    if not BS_ROWS.exists():
        return Response("No chart yet.", status=404, mimetype="text/plain")
    rows = json.loads(BS_ROWS.read_text())
    png  = bs_make_chart(rows)
    BS_CHART.write_bytes(png)
    return send_file(io.BytesIO(png), mimetype="image/png", download_name="Build_State_Chart_latest.png")


@app.route("/build/latest-summary")
def build_latest_summary():
    if not BS_SUMMARY.exists():
        return jsonify({"error": "No data yet"}), 404
    data = json.loads(BS_SUMMARY.read_text())
    data.pop("chart_b64", None)
    return jsonify(data)


@app.route("/build/upload", methods=["POST"])
def build_upload():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400
    fname = f.filename.lower()
    uploaded_by = request.form.get("uploaded_by", "").strip() or "Rishab"
    try:
        raw = f.read()
        if fname.endswith(".csv"):
            rows = bs_parse_csv(raw)
        elif fname.endswith((".xlsx", ".xls")):
            rows = bs_parse_xlsx(raw)
        else:
            return jsonify({"error": "Upload a .csv or .xlsx file"}), 400
    except ValueError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        return jsonify({"error": f"Could not read file: {e}"}), 500
    try:
        png = bs_make_chart(rows)
    except Exception as e:
        return jsonify({"error": f"Chart generation failed: {e}"}), 500
    summary   = bs_summarise(rows, uploaded_by)
    chart_b64 = base64.b64encode(png).decode()
    BS_SUMMARY.write_text(json.dumps({**summary, "chart_b64": chart_b64}))
    BS_ROWS.write_text(json.dumps(rows))
    BS_CHART.write_bytes(png)
    return jsonify({"ok": True, "summary": summary})


@app.route("/build/chart", methods=["POST"])
def build_chart():
    body = request.get_json(silent=True) or {}
    rows = body.get("rows")
    if not rows:
        return jsonify({"error": "No data"}), 400
    try:
        png = bs_make_chart(rows)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    return send_file(io.BytesIO(png), mimetype="image/png",
                     download_name=f"Build_State_Chart_{date.today().isoformat()}.png")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(debug=True, host="0.0.0.0", port=port)
