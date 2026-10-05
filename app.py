from flask import Flask, request, jsonify, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

app = Flask(__name__)

# ── Color palette ──────────────────────────────────────────────
NAVY   = "4A235A"   # Plum (headers, banners)
TEAL   = "7B4F8B"   # Mid-purple (secondary headers)
GOLD   = "C9956B"   # Warm terracotta accent
LIGHT  = "F5EFF8"   # Pale lavender tint (label cells)
WHITE  = "FBF7F4"   # Warm cream (main cell background)
GRAY   = "EDE5F0"   # Soft purple-gray (alternating rows)
DKGRAY = "4A235A"   # Plum (body text on light)

def make_border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def hdr(ws, row, col, value, bg=NAVY, fg=WHITE, bold=True, size=11):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Arial", bold=bold, color=fg, size=size)
    c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = make_border()
    return c

def cell(ws, row, col, value="", bold=False, color=DKGRAY, bg=WHITE, align="left", wrap=True):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Arial", bold=bold, color=color, size=10)
    c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    c.border = make_border()
    return c

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def freeze(ws, ref="A2"):
    ws.freeze_panes = ref


# ══════════════════════════════════════════════════════════════
#  SHEET 1 — Business Summary
# ══════════════════════════════════════════════════════════════
def build_summary(wb, d):
    ws = wb.create_sheet("Business Summary")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:B1")
    t = ws["A1"]
    t.value = "AdCraft — Advertising Plan"
    t.font = Font(name="Arial", bold=True, color=WHITE, size=16)
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    rows = [
        ("Business Type",       d.get("business_type", "")),
        ("Primary Goal",        d.get("primary_goal", "")),
        ("Target Customer",     d.get("target_customer", "")),
        ("Geography",           d.get("geography", "")),
        ("Monthly Budget",      f"${d.get('monthly_budget', 0):,.0f}"),
        ("Monthly Revenue",     f"${d.get('monthly_revenue', 0):,.0f}"),
        ("Budget Tier",         d.get("budget_tier", "")),
        ("Current Advertising", d.get("current_advertising", "None")),
    ]
    for i, (label, val) in enumerate(rows, start=2):
        cell(ws, i, 1, label, bold=True, bg=LIGHT)
        cell(ws, i, 2, val, bg=WHITE)
        ws.row_dimensions[i].height = 22

    set_col_widths(ws, [28, 50])


# ══════════════════════════════════════════════════════════════
#  SHEET 2 — Channel Plan
# ══════════════════════════════════════════════════════════════
def build_channel_plan(wb, d):
    ws = wb.create_sheet("Channel Plan")
    ws.sheet_view.showGridLines = False

    headers = ["Channel", "Monthly Budget", "% of Budget", "Primary KPI",
               "Expected Reach/Clicks", "Est. Cost Per Result", "Notes"]
    widths  = [22, 16, 14, 20, 22, 20, 40]

    for i, h in enumerate(headers, 1):
        hdr(ws, 1, i, h)
    ws.row_dimensions[1].height = 28
    freeze(ws)

    channels = d.get("channels", [])
    total_budget = d.get("monthly_budget", 0)

    for r, ch in enumerate(channels, start=2):
        bg = GRAY if r % 2 == 0 else WHITE
        budget_val = ch.get("monthly_budget", 0)
        pct = (budget_val / total_budget * 100) if total_budget else 0
        row_data = [
            ch.get("channel", ""),
            budget_val,
            f"{pct:.0f}%",
            ch.get("kpi", ""),
            ch.get("expected_reach", ""),
            ch.get("cost_per_result", ""),
            ch.get("notes", ""),
        ]
        for c_idx, val in enumerate(row_data, 1):
            cell(ws, r, c_idx, val, bg=bg,
                 align="right" if c_idx == 2 else "left")
        ws.row_dimensions[r].height = 22

    last = len(channels) + 2
    cell(ws, last, 1, "TOTAL", bold=True, bg=NAVY, color=WHITE, align="center")
    cell(ws, last, 2, sum(c.get("monthly_budget", 0) for c in channels),
         bold=True, bg=NAVY, color=WHITE, align="right")
    for col in range(3, 8):
        cell(ws, last, col, "", bg=NAVY)
    ws.row_dimensions[last].height = 24

    set_col_widths(ws, widths)


# ══════════════════════════════════════════════════════════════
#  SHEET 3 — Excluded Channels
# ══════════════════════════════════════════════════════════════
def build_excluded(wb, d):
    ws = wb.create_sheet("Excluded Channels")
    ws.sheet_view.showGridLines = False

    hdr(ws, 1, 1, "Channel", bg=TEAL)
    hdr(ws, 1, 2, "Reason Not Recommended", bg=TEAL)
    ws.row_dimensions[1].height = 28
    freeze(ws)

    excluded = d.get("excluded_channels", [])
    for r, ex in enumerate(excluded, start=2):
        bg = GRAY if r % 2 == 0 else WHITE
        cell(ws, r, 1, ex.get("channel", ""), bg=bg)
        cell(ws, r, 2, ex.get("reason", ""), bg=bg)
        ws.row_dimensions[r].height = 22

    set_col_widths(ws, [28, 60])


# ══════════════════════════════════════════════════════════════
#  SHEET 4 — Monthly Tracker
# ══════════════════════════════════════════════════════════════
def build_tracker(wb, d):
    ws = wb.create_sheet("Monthly Tracker")
    ws.sheet_view.showGridLines = False

    channels = d.get("channels", [])
    months = ["Month 1", "Month 2", "Month 3"]

    hdr(ws, 1, 1, "Channel")
    hdr(ws, 1, 2, "KPI")
    col = 3
    for m in months:
        hdr(ws, 1, col,   f"{m} — Planned")
        hdr(ws, 1, col+1, f"{m} — Actual")
        hdr(ws, 1, col+2, f"{m} — Variance")
        col += 3
    ws.row_dimensions[1].height = 28
    freeze(ws)

    for r, ch in enumerate(channels, start=2):
        bg = GRAY if r % 2 == 0 else WHITE
        cell(ws, r, 1, ch.get("channel", ""), bg=bg)
        cell(ws, r, 2, ch.get("kpi", ""), bg=bg)
        col = 3
        for _ in months:
            cell(ws, r, col,   ch.get("expected_reach", ""), bg=bg, align="right")
            cell(ws, r, col+1, "", bg=LIGHT)
            c = ws.cell(row=r, column=col+2)
            c.value = f"={get_column_letter(col+1)}{r}-{get_column_letter(col)}{r}"
            c.font = Font(name="Arial", size=10, color=DKGRAY)
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.border = make_border()
            col += 3
        ws.row_dimensions[r].height = 22

    widths = [22, 20] + [18, 18, 16] * 3
    set_col_widths(ws, widths)


# ══════════════════════════════════════════════════════════════
#  SHEET 5 — Benchmarks
# ══════════════════════════════════════════════════════════════
def build_benchmarks(wb, d):
    ws = wb.create_sheet("Benchmarks")
    ws.sheet_view.showGridLines = False

    headers = ["Channel", "Avg CTR", "Avg CPC", "Avg CPM", "Avg Conv Rate", "Notes"]
    widths  = [22, 14, 14, 14, 16, 50]
    for i, h in enumerate(headers, 1):
        hdr(ws, 1, i, h)
    ws.row_dimensions[1].height = 28
    freeze(ws)

    benchmarks = d.get("benchmarks", [])
    for r, b in enumerate(benchmarks, start=2):
        bg = GRAY if r % 2 == 0 else WHITE
        row_data = [
            b.get("channel", ""),
            b.get("avg_ctr", ""),
            b.get("avg_cpc", ""),
            b.get("avg_cpm", ""),
            b.get("avg_conv_rate", ""),
            b.get("notes", ""),
        ]
        for c_idx, val in enumerate(row_data, 1):
            cell(ws, r, c_idx, val, bg=bg)
        ws.row_dimensions[r].height = 22

    set_col_widths(ws, widths)


# ══════════════════════════════════════════════════════════════
#  SHEET 6 — 90-Day Launch Plan
# ══════════════════════════════════════════════════════════════
def build_launch_plan(wb, d):
    ws = wb.create_sheet("90-Day Launch Plan")
    ws.sheet_view.showGridLines = False

    headers = ["Week", "Phase", "Action Item", "Owner", "Status"]
    widths  = [10, 20, 50, 16, 14]
    for i, h in enumerate(headers, 1):
        hdr(ws, 1, i, h)
    ws.row_dimensions[1].height = 28
    freeze(ws)

    launch_items = d.get("launch_plan", [])
    for r, item in enumerate(launch_items, start=2):
        bg = GRAY if r % 2 == 0 else WHITE
        row_data = [
            item.get("week", ""),
            item.get("phase", ""),
            item.get("action", ""),
            item.get("owner", "You"),
            item.get("status", "To Do"),
        ]
        for c_idx, val in enumerate(row_data, 1):
            cell(ws, r, c_idx, val, bg=bg)
        ws.row_dimensions[r].height = 22

    set_col_widths(ws, widths)


# ══════════════════════════════════════════════════════════════
#  MAIN ENDPOINT
# ══════════════════════════════════════════════════════════════
@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({"error": "No JSON body received"}), 400

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        build_summary(wb, data)
        build_channel_plan(wb, data)
        build_excluded(wb, data)
        build_tracker(wb, data)
        build_benchmarks(wb, data)
        build_launch_plan(wb, data)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        biz = data.get("business_type", "AdCraft").replace(" ", "_")
        filename = f"AdCraft_Plan_{biz}.xlsx"

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "AdCraft XLSX Generator"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
