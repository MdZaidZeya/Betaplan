import io
import os
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

# ── persistent storage ────────────────────────────────────────────────────────
# Chart is stored as base64 INSIDE the summary JSON so the whole page —
# including the image — is delivered in a single request. No broken image on
# free-tier /tmp restarts, and no second HTTP round-trip for Jim.
DATA_DIR = Path(os.environ.get("DATA_DIR", "/tmp/betaplan_data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)
SUMMARY_PATH = DATA_DIR / "latest_summary.json"   # contains chart_b64 key
ROWS_PATH    = DATA_DIR / "latest_rows.json"
CHART_PATH   = DATA_DIR / "latest_chart.png"      # kept for download endpoint

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


# ── helpers ───────────────────────────────────────────────────────────────────

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


def make_chart(rows: list[dict], snapshot_date: str | None = None) -> bytes:
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
            items.append({
                "kind": "lane",
                "label": label,
                "pct": lane["pct"],
                "status": lane["status"],
                "owner": lane["owner"],
            })

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
            ax.barh(y, item["pct"], color=color, height=0.62,
                    edgecolor="#999999", linewidth=0.4)
            ax.barh(y, 100, color="none", edgecolor="#DDDDDD",
                    height=0.62, linewidth=0.4, zorder=0)
            lx = item["pct"] + 1 if item["pct"] < 88 else item["pct"] - 6
            lc = "#333333" if item["pct"] < 88 else "white"
            ax.text(lx, y, f"{item['pct']}%", fontsize=8,
                    va="center", color=lc, weight="bold")
            if item.get("owner"):
                ax.text(102, y, item["owner"], fontsize=7,
                        va="center", color="#888888")

    ax.set_yticks(y_positions)
    ax.set_yticklabels([it["label"] for it in items], fontsize=8.5)
    ax.set_xlim(0, 115)
    ax.set_xlabel("Progress (%)", fontsize=10)
    ax.set_xticks([0, 25, 50, 75, 100])
    ax.set_title(
        f"Beta Plan 2026 — Progress Snapshot  ({snap})",
        fontsize=13, weight="bold", pad=14,
    )
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.spines["left"].set_color("#CCCCCC")
    ax.spines["bottom"].set_color("#CCCCCC")
    ax.tick_params(colors="#666666")
    ax.set_axisbelow(True)
    ax.grid(axis="x", color="#EEEEEE", linestyle="-", linewidth=0.5)

    legend_handles = [
        Patch(facecolor=c, edgecolor="#999999", label=s)
        for s, c in STATUS_COLORS.items()
    ]
    ax.legend(handles=legend_handles, loc="lower right", fontsize=9, frameon=False)

    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=140, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def summarise(rows: list[dict], uploaded_by: str = "") -> dict:
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
        ws_summary[ws] = {
            "avg": ws_avg,
            "count": len(lanes),
            "by_status": {},
        }
        for r in lanes:
            s = r["status"]
            ws_summary[ws]["by_status"][s] = ws_summary[ws]["by_status"].get(s, 0) + 1

    return {
        "total": total,
        "avg": avg,
        "by_status": by_status,
        "workstreams": ws_summary,
        "snapshot_date": date.today().isoformat(),
        "updated_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
        "uploaded_by": uploaded_by,
    }


# ── routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    """Main dashboard. chart_b64 is embedded in the page — no second request."""
    has_data = SUMMARY_PATH.exists()
    summary = None
    chart_b64 = None
    if has_data:
        data = json.loads(SUMMARY_PATH.read_text())
        chart_b64 = data.pop("chart_b64", None)   # pull out before passing to template
        summary = data
    return render_template("index.html", has_data=has_data,
                           summary=summary, chart_b64=chart_b64)


@app.route("/latest-chart")
def latest_chart():
    """Download endpoint — regenerates from saved rows if PNG file is missing."""
    if CHART_PATH.exists():
        return send_file(CHART_PATH, mimetype="image/png",
                         download_name="Progress_Chart_latest.png")
    # PNG file gone (free-tier /tmp wipe) — regenerate from saved rows
    if not ROWS_PATH.exists():
        return Response("No chart yet — upload a file first.", status=404, mimetype="text/plain")
    rows = json.loads(ROWS_PATH.read_text())
    png = make_chart(rows)
    CHART_PATH.write_bytes(png)
    return send_file(io.BytesIO(png), mimetype="image/png",
                     download_name="Progress_Chart_latest.png")


@app.route("/latest-summary")
def latest_summary():
    if not SUMMARY_PATH.exists():
        return jsonify({"error": "No data yet"}), 404
    data = json.loads(SUMMARY_PATH.read_text())
    data.pop("chart_b64", None)   # don't expose huge base64 in JSON API
    return jsonify(data)


@app.route("/upload", methods=["POST"])
def upload():
    """Rishab uploads xlsx → parse → generate chart → persist → respond."""
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Please upload an .xlsx or .xls file"}), 400

    uploaded_by = request.form.get("uploaded_by", "").strip() or "Rishab"

    try:
        file_bytes = f.read()
        rows = parse_xlsx(file_bytes)
    except ValueError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        return jsonify({"error": f"Could not read file: {e}"}), 500

    try:
        png = make_chart(rows)
    except Exception as e:
        return jsonify({"error": f"Chart generation failed: {e}"}), 500

    summary = summarise(rows, uploaded_by)

    # Embed chart as base64 in summary so the homepage never needs a second request
    chart_b64 = base64.b64encode(png).decode()
    payload = {**summary, "chart_b64": chart_b64}

    SUMMARY_PATH.write_text(json.dumps(payload))
    ROWS_PATH.write_text(json.dumps(rows))
    CHART_PATH.write_bytes(png)

    return jsonify({"ok": True, "summary": summary, "rows": rows})


@app.route("/chart", methods=["POST"])
def chart():
    """Inline chart render for the upload-tab preview."""
    body = request.get_json(silent=True) or {}
    rows = body.get("rows")
    if not rows:
        return jsonify({"error": "No data"}), 400
    try:
        png = make_chart(rows)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    return send_file(io.BytesIO(png), mimetype="image/png",
                     download_name=f"Progress_Chart_{date.today().isoformat()}.png")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(debug=True, host="0.0.0.0", port=port)
