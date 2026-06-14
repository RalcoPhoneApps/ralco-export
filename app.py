"""
RALCO Field Walkthrough — Export + Cloud Storage Backend v2
Handles Excel export and cloud job storage with PostgreSQL.
"""

import io
import os
import re
import json
import uuid
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.xlsx")
# Fix DATABASE_URL for psycopg2 compatibility
raw_url = os.environ.get("DATABASE_URL", "")
if raw_url.startswith("postgres://"):
    raw_url = raw_url.replace("postgres://", "postgresql://", 1)
DATABASE_URL = raw_url or None

# ── Database setup ────────────────────────────────────────────
def get_db():
    import psycopg2
    conn = psycopg2.connect(DATABASE_URL, sslmode='require')
    return conn

def init_db():
    """Create tables if they don't exist."""
    if not DATABASE_URL:
        print("WARNING: No DATABASE_URL set — cloud storage disabled")
        return
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS jobs (
                id TEXT PRIMARY KEY,
                customer TEXT,
                device_id TEXT,
                device_name TEXT,
                rep TEXT,
                job_date TEXT,
                data JSONB NOT NULL,
                created_at TIMESTAMP DEFAULT NOW(),
                updated_at TIMESTAMP DEFAULT NOW()
            )
        """)
        conn.commit()
        cur.close()
        conn.close()
        print("Database initialized successfully")
    except Exception as e:
        print(f"Database init error: {e}")

# Initialize on startup
with app.app_context():
    init_db()

# ── Equipment + service mapping ───────────────────────────────
EQUIPMENT_ORDER = [
    "swgr_main","swgr_icb","swgr_dob","swbd","swbd_disc","swbd_mcb",
    "panel_dist","panel_branch","mcc","mcc_starter","mcc_mcb","mcc_disc",
    "mcp","vfd_mcp","disc_ind","mcb_ind","starter_ind","xfmr_oil_mv",
    "xfmr_oil_hv","xfmr_dry","pf_cap","ltg_cont","pdu","wireway","busduct",
    "ats","mts","vfd_sm","vfd_lg","motor_sm","motor_md","motor_lg",
    "comp","hvac_lg","hvac_sm","ups_sm","ups_lg",
]
EQ_COUNT = len(EQUIPMENT_ORDER)

SVC_COLS = {
    "basic":    [4],
    "full":     [5],
    "full_af":  [5, 9],
    "full_den": [5, 6],
    "full_all": [5, 6, 9],
    "ir":       [7],
    "ut":       [8],
}

def get_building_name(job, building_id):
    if not building_id:
        return "Main Bldg."
    for b in job.get("buildings", []):
        if b.get("id") == building_id:
            return b.get("name", "Main Bldg.")
    return "Main Bldg."

def sanitize_filename(name):
    safe = re.sub(r"[^a-zA-Z0-9 _\-]", "_", name or "Job")
    return safe.strip("_ ") or "Job"

def build_excel(job):
    """Populate template with job data, return BytesIO buffer."""
    wb = load_workbook(TEMPLATE_PATH)
    ws_br = wb["Bldgs & Rooms"]
    ws_c1 = wb["Components Yr 1"]

    rooms = job.get("rooms", [])
    for idx, room in enumerate(rooms):
        row_num = idx + 2
        bldg_name = get_building_name(job, room.get("buildingId"))
        ws_br.cell(row=row_num, column=1, value=bldg_name)
        ws_br.cell(row=row_num, column=2, value=room.get("name", ""))

    default_svc = job.get("tier", "basic")
    for room_idx, room in enumerate(rooms):
        equipment = room.get("equipment", {})
        overrides = room.get("overrides", {})
        for eq_idx, eq_id in enumerate(EQUIPMENT_ORDER):
            qty = equipment.get(eq_id, 0)
            if not qty:
                continue
            svc = overrides.get(eq_id, default_svc)
            col_indices = SVC_COLS.get(svc, [4])
            row_num = (room_idx * EQ_COUNT) + eq_idx + 2
            for col_idx in col_indices:
                ws_c1.cell(row=row_num, column=col_idx + 1, value=qty)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ── Health check ──────────────────────────────────────────────
@app.route("/health", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "template": os.path.exists(TEMPLATE_PATH),
        "database": DATABASE_URL is not None
    })

# ── Export ────────────────────────────────────────────────────
@app.route("/export", methods=["POST"])
def export():
    try:
        job = request.get_json(force=True)
        if not job:
            return jsonify({"error": "No job data received"}), 400
        if not job.get("rooms"):
            return jsonify({"error": "No rooms in job"}), 400
        if len(job.get("rooms", [])) > 30:
            return jsonify({"error": "Maximum 30 rooms supported"}), 400

        buffer = build_excel(job)
        filename = sanitize_filename(job.get("customer", "Job")) + "_EMA_Sales_Builder.xlsx"

        return send_file(
            buffer,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )
    except Exception as e:
        app.logger.error(f"Export error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

# ── Save job ──────────────────────────────────────────────────
@app.route("/jobs/save", methods=["POST"])
def save_job():
    if not DATABASE_URL:
        return jsonify({"error": "Database not configured"}), 503
    try:
        payload = request.get_json(force=True)
        job = payload.get("job")
        device_id = payload.get("deviceId", "unknown")
        device_name = payload.get("deviceName", "Unnamed Device")

        if not job or not job.get("id"):
            return jsonify({"error": "Invalid job data"}), 400

        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO jobs (id, customer, device_id, device_name, rep, job_date, data, updated_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, NOW())
            ON CONFLICT (id) DO UPDATE SET
                customer = EXCLUDED.customer,
                device_id = EXCLUDED.device_id,
                device_name = EXCLUDED.device_name,
                rep = EXCLUDED.rep,
                job_date = EXCLUDED.job_date,
                data = EXCLUDED.data,
                updated_at = NOW()
        """, (
            job["id"],
            job.get("customer", ""),
            device_id,
            device_name,
            job.get("rep", ""),
            job.get("date", ""),
            json.dumps(job)
        ))
        conn.commit()
        cur.close()
        conn.close()
        return jsonify({"status": "saved", "id": job["id"]})
    except Exception as e:
        app.logger.error(f"Save error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

# ── List all jobs ─────────────────────────────────────────────
@app.route("/jobs", methods=["GET"])
def list_jobs():
    if not DATABASE_URL:
        return jsonify({"error": "Database not configured"}), 503
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            SELECT id, customer, device_name, rep, job_date, updated_at
            FROM jobs
            ORDER BY updated_at DESC
        """)
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return jsonify([{
            "id": r[0],
            "customer": r[1],
            "deviceName": r[2],
            "rep": r[3],
            "date": r[4],
            "updatedAt": r[5].isoformat() if r[5] else None
        } for r in rows])
    except Exception as e:
        app.logger.error(f"List error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

# ── Load single job ───────────────────────────────────────────
@app.route("/jobs/<job_id>", methods=["GET"])
def load_job(job_id):
    if not DATABASE_URL:
        return jsonify({"error": "Database not configured"}), 503
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT data FROM jobs WHERE id = %s", (job_id,))
        row = cur.fetchone()
        cur.close()
        conn.close()
        if not row:
            return jsonify({"error": "Job not found"}), 404
        return jsonify(row[0])
    except Exception as e:
        app.logger.error(f"Load error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

# ── Delete job ────────────────────────────────────────────────
@app.route("/jobs/<job_id>", methods=["DELETE"])
def delete_job_cloud(job_id):
    if not DATABASE_URL:
        return jsonify({"error": "Database not configured"}), 503
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("DELETE FROM jobs WHERE id = %s", (job_id,))
        conn.commit()
        cur.close()
        conn.close()
        return jsonify({"status": "deleted", "id": job_id})
    except Exception as e:
        app.logger.error(f"Delete error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)


# ── Tag counter ───────────────────────────────────────────────
_tag_counter = {'value': None}

@app.route("/tag-counter", methods=["GET"])
def get_tag_counter():
    if not DATABASE_URL:
        return jsonify({"counter": _tag_counter['value']})
    try:
        conn = get_db(); cur = conn.cursor()
        cur.execute("SELECT value FROM meta WHERE key='tag_counter'")
        row = cur.fetchone()
        cur.close(); conn.close()
        return jsonify({"counter": row[0] if row else None})
    except Exception as e:
        return jsonify({"counter": None})

@app.route("/tag-counter", methods=["POST"])
def set_tag_counter():
    data = request.get_json(force=True)
    counter = data.get('counter')
    if not counter: return jsonify({"status":"ignored"}), 200
    _tag_counter['value'] = counter
    if not DATABASE_URL: return jsonify({"status":"ok"})
    try:
        conn = get_db(); cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value BIGINT)
        """)
        cur.execute("""
            INSERT INTO meta (key, value) VALUES ('tag_counter', %s)
            ON CONFLICT (key) DO UPDATE SET value = GREATEST(meta.value, EXCLUDED.value)
        """, (int(counter),))
        conn.commit(); cur.close(); conn.close()
        return jsonify({"status":"ok"})
    except Exception as e:
        return jsonify({"status":"error", "error":str(e)}), 500

# ── EPM Master Export ─────────────────────────────────────────
EPM_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "epm_template.xlsx")

@app.route("/export/epm", methods=["POST"])
def export_epm():
    try:
        payload = request.get_json(force=True)
        job = payload.get("job", {})
        assets = payload.get("assets", [])
        if not assets:
            return jsonify({"error": "No assets to export"}), 400

        template = EPM_TEMPLATE_PATH if os.path.exists(EPM_TEMPLATE_PATH) else TEMPLATE_PATH
        wb = load_workbook(template)

        # Write Equipment Information sheet
        if "Equipment Information" in wb.sheetnames:
            ws_ei = wb["Equipment Information"]
            for idx, a in enumerate(assets, start=2):
                ws_ei.cell(row=idx, column=1, value=a.get("tag",""))
                ws_ei.cell(row=idx, column=2, value=a.get("name",""))
                ws_ei.cell(row=idx, column=3, value=a.get("location",""))
                ws_ei.cell(row=idx, column=4, value=a.get("eqType",""))
                ws_ei.cell(row=idx, column=5, value=a.get("manufacturer",""))
                ws_ei.cell(row=idx, column=6, value=a.get("model",""))
                ws_ei.cell(row=idx, column=7, value=a.get("serial",""))
                ws_ei.cell(row=idx, column=8, value=a.get("voltage",""))
                ws_ei.cell(row=idx, column=9, value=a.get("criticality",""))
                ws_ei.cell(row=idx, column=10, value=a.get("maintStatus",""))
                ws_ei.cell(row=idx, column=13, value=a.get("notes",""))

        # Write Code Violations sheet
        if "Code Violations" in wb.sheetnames:
            ws_cv = wb["Code Violations"]
            for idx, a in enumerate(assets, start=2):
                ws_cv.cell(row=idx, column=1, value=a.get("tag",""))
                ws_cv.cell(row=idx, column=2, value=a.get("name",""))
                ws_cv.cell(row=idx, column=3, value=a.get("location",""))
                violations = a.get("codeViolations", [])
                for vi, v in enumerate(violations[:5]):
                    if not v.get("article"): continue
                    col_base = 4 + (vi * 3)
                    ws_cv.cell(row=idx, column=col_base,   value=v.get("article",""))
                    ws_cv.cell(row=idx, column=col_base+1, value=v.get("sev",""))
                    ws_cv.cell(row=idx, column=col_base+2, value=v.get("corrected","No"))

        # Write Infrared sheet
        if "Infrared" in wb.sheetnames:
            ws_ir = wb["Infrared"]
            for idx, a in enumerate(assets, start=2):
                ir = a.get("ir", {})
                ws_ir.cell(row=idx, column=1, value=a.get("tag",""))
                ws_ir.cell(row=idx, column=2, value=a.get("name",""))
                ws_ir.cell(row=idx, column=3, value=a.get("location",""))
                ws_ir.cell(row=idx, column=4, value=ir.get("phase",""))
                ws_ir.cell(row=idx, column=5, value=ir.get("component",""))
                ws_ir.cell(row=idx, column=6, value=ir.get("result",""))
                ws_ir.cell(row=idx, column=7, value=ir.get("notes",""))

        # Write Ultrasonic sheet
        if "Ultrasonic" in wb.sheetnames:
            ws_ut = wb["Ultrasonic"]
            for idx, a in enumerate(assets, start=2):
                ut = a.get("ut", {})
                ws_ut.cell(row=idx, column=1, value=a.get("tag",""))
                ws_ut.cell(row=idx, column=2, value=a.get("name",""))
                ws_ut.cell(row=idx, column=3, value=a.get("location",""))
                ws_ut.cell(row=idx, column=4, value=ut.get("component",""))
                ws_ut.cell(row=idx, column=5, value=ut.get("failureMode",""))
                ws_ut.cell(row=idx, column=6, value=ut.get("db",""))
                ws_ut.cell(row=idx, column=7, value=ut.get("result",""))
                ws_ut.cell(row=idx, column=8, value=ut.get("notes",""))

        # Write Energized Data sheet
        if "Energized Data" in wb.sheetnames:
            ws_en = wb["Energized Data"]
            for idx, a in enumerate(assets, start=2):
                en = a.get("energized", {})
                ws_en.cell(row=idx, column=1, value=a.get("tag",""))
                ws_en.cell(row=idx, column=2, value=a.get("name",""))
                ws_en.cell(row=idx, column=3, value=a.get("location",""))
                ws_en.cell(row=idx, column=4, value=en.get("phase",""))
                ws_en.cell(row=idx, column=5, value=en.get("wire",""))
                ws_en.cell(row=idx, column=6, value=en.get("vab",""))
                ws_en.cell(row=idx, column=7, value=en.get("vbc",""))
                ws_en.cell(row=idx, column=8, value=en.get("vca",""))
                ws_en.cell(row=idx, column=10, value=en.get("aa",""))
                ws_en.cell(row=idx, column=11, value=en.get("ab",""))
                ws_en.cell(row=idx, column=12, value=en.get("ac",""))
                ws_en.cell(row=idx, column=14, value=en.get("ratedAmps",""))
                ws_en.cell(row=idx, column=15, value=en.get("ocpd",""))
                ws_en.cell(row=idx, column=18, value=en.get("analysis","N/A"))
                ws_en.cell(row=idx, column=19, value=en.get("summary","PASS"))

        # Write Insulation sheet
        if "Insulation" in wb.sheetnames:
            ws_ins = wb["Insulation"]
            for idx, a in enumerate(assets, start=2):
                ins = a.get("insulation", {})
                ws_ins.cell(row=idx, column=1, value=a.get("tag",""))
                ws_ins.cell(row=idx, column=2, value=a.get("name",""))
                ws_ins.cell(row=idx, column=3, value=a.get("location",""))
                ws_ins.cell(row=idx, column=4, value=ins.get("testV",""))
                ws_ins.cell(row=idx, column=5, value=ins.get("phPh",""))
                ws_ins.cell(row=idx, column=6, value=ins.get("phG",""))
                ws_ins.cell(row=idx, column=7, value=ins.get("pi",""))
                ws_ins.cell(row=idx, column=8, value=ins.get("temp",""))
                ws_ins.cell(row=idx, column=9, value=ins.get("result","N/A"))
                ws_ins.cell(row=idx, column=10, value=ins.get("notes",""))

        # Write DLRO sheet
        if "DLRO" in wb.sheetnames:
            ws_dl = wb["DLRO"]
            for idx, a in enumerate(assets, start=2):
                dl = a.get("dlro", {})
                ws_dl.cell(row=idx, column=1, value=a.get("tag",""))
                ws_dl.cell(row=idx, column=2, value=a.get("name",""))
                ws_dl.cell(row=idx, column=3, value=a.get("location",""))
                ws_dl.cell(row=idx, column=4, value=dl.get("phA",""))
                ws_dl.cell(row=idx, column=5, value=dl.get("phB",""))
                ws_dl.cell(row=idx, column=6, value=dl.get("phC",""))
                ws_dl.cell(row=idx, column=7, value=dl.get("dev",""))
                ws_dl.cell(row=idx, column=8, value=dl.get("result","N/A"))
                ws_dl.cell(row=idx, column=9, value=dl.get("notes",""))

        # Write Torque sheet
        if "Torque" in wb.sheetnames:
            ws_tq = wb["Torque"]
            for idx, a in enumerate(assets, start=2):
                tq = a.get("torque", {})
                ws_tq.cell(row=idx, column=1, value=a.get("tag",""))
                ws_tq.cell(row=idx, column=2, value=a.get("name",""))
                ws_tq.cell(row=idx, column=3, value=a.get("location",""))
                ws_tq.cell(row=idx, column=4, value=tq.get("phase",""))
                ws_tq.cell(row=idx, column=5, value=tq.get("connType",""))
                ws_tq.cell(row=idx, column=6, value=tq.get("spec",""))
                ws_tq.cell(row=idx, column=7, value=tq.get("verified",""))
                ws_tq.cell(row=idx, column=8, value=tq.get("result","N/A"))
                ws_tq.cell(row=idx, column=9, value=tq.get("notes",""))

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        customer = job.get("customer", "Job")
        filename = sanitize_filename(customer) + "_EPM_Master.xlsx"
        return send_file(buffer,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True, download_name=filename)

    except Exception as e:
        app.logger.error(f"EPM export error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500
