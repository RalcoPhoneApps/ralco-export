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
