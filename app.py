"""
RALCO Field Walkthrough — Excel Export Backend
Receives job JSON from the field app, populates the EMA Sales Builder
template using openpyxl (preserving all formatting), and returns the
completed xlsx file for download.
"""

import io
import os
import re
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)  # Allow requests from GitHub Pages

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.xlsx")

# Equipment list in exact Hours sheet order (rows 2-38)
# Must match the walkthrough app's EQUIPMENT array exactly
EQUIPMENT_ORDER = [
    "swgr_main",
    "swgr_icb",
    "swgr_dob",
    "swbd",
    "swbd_disc",
    "swbd_mcb",
    "panel_dist",
    "panel_branch",
    "mcc",
    "mcc_starter",
    "mcc_mcb",
    "mcc_disc",
    "mcp",
    "vfd_mcp",
    "disc_ind",
    "mcb_ind",
    "starter_ind",
    "xfmr_oil_mv",
    "xfmr_oil_hv",
    "xfmr_dry",
    "pf_cap",
    "ltg_cont",
    "pdu",
    "wireway",
    "busduct",
    "ats",
    "mts",
    "vfd_sm",
    "vfd_lg",
    "motor_sm",
    "motor_md",
    "motor_lg",
    "comp",
    "hvac_lg",
    "hvac_sm",
    "ups_sm",
    "ups_lg",
]

EQ_COUNT = len(EQUIPMENT_ORDER)  # 37

# Service type -> which column indices get the qty written
# Columns: E=4, F=5, G=6, H=7, I=8, J=9 (0-indexed)
SVC_COLS = {
    "basic":    [4],        # E: Basic EPM
    "full":     [5],        # F: Full EPM
    "full_af":  [5, 9],     # F: Full EPM + J: AF Collection
    "full_den": [5, 6],     # F: Full EPM + G: De-Energized
    "full_all": [5, 6, 9],  # F + G + J
    "ir":       [7],        # H: IR only (legacy, kept for safety)
    "ut":       [8],        # I: UT only (legacy, kept for safety)
}


def get_building_name(job, building_id):
    """Resolve building name from id, default to Main Bldg."""
    if not building_id:
        return "Main Bldg."
    for b in job.get("buildings", []):
        if b.get("id") == building_id:
            return b.get("name", "Main Bldg.")
    return "Main Bldg."


def sanitize_filename(name):
    """Make a safe filename from customer name."""
    safe = re.sub(r"[^a-zA-Z0-9 _\-]", "_", name or "Job")
    return safe.strip("_ ") or "Job"


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "template": os.path.exists(TEMPLATE_PATH)})


@app.route("/export", methods=["POST"])
def export():
    try:
        job = request.get_json(force=True)
        if not job:
            return jsonify({"error": "No job data received"}), 400

        rooms = job.get("rooms", [])
        if not rooms:
            return jsonify({"error": "No rooms in job"}), 400

        if len(rooms) > 30:
            return jsonify({"error": "Maximum 30 rooms supported"}), 400

        # Load template fresh for each export (preserves all formatting)
        wb = load_workbook(TEMPLATE_PATH)
        ws_br = wb["Bldgs & Rooms"]
        ws_c1 = wb["Components Yr 1"]

        # ── Bldgs & Rooms sheet ─────────────────────────────────
        # Rows 2-31 (index 1-30), col A = building, col B = room name
        for idx, room in enumerate(rooms):
            row_num = idx + 2  # 1-indexed, starting at row 2
            bldg_name = get_building_name(job, room.get("buildingId"))
            ws_br.cell(row=row_num, column=1, value=bldg_name)
            ws_br.cell(row=row_num, column=2, value=room.get("name", ""))

        # ── Components Yr 1 sheet ───────────────────────────────
        # Each room block = 37 rows (one per equipment type)
        # Room 0 = rows 2-38, Room 1 = rows 39-75, etc.
        # Only write qty cols (E-J), never touch A/B/C (formula refs)
        default_svc = job.get("tier", "basic")

        for room_idx, room in enumerate(rooms):
            equipment = room.get("equipment", {})
            overrides = room.get("overrides", {})

            for eq_idx, eq_id in enumerate(EQUIPMENT_ORDER):
                qty = equipment.get(eq_id, 0)
                if not qty:
                    continue

                # Which service type applies to this asset
                svc = overrides.get(eq_id, default_svc)
                col_indices = SVC_COLS.get(svc, [4])  # default to Basic EPM col

                # Row in Components Yr 1 (1-indexed for openpyxl)
                row_num = (room_idx * EQ_COUNT) + eq_idx + 2

                for col_idx in col_indices:
                    # openpyxl uses 1-indexed columns
                    ws_c1.cell(row=row_num, column=col_idx + 1, value=qty)

        # ── Save to memory buffer ────────────────────────────────
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        customer = job.get("customer", "Job")
        filename = sanitize_filename(customer) + "_EMA_Sales_Builder.xlsx"

        return send_file(
            buffer,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )

    except Exception as e:
        app.logger.error(f"Export error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
