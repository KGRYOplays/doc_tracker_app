import os
import pandas as pd
import qrcode
import random
import string
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, jsonify, session

import db_manager

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'barcode-routing-secret-2024-change-in-prod')

MASTER_DB  = 'master_db.xlsx'
QR_FOLDER  = os.path.join('static', 'qr_generated')

# Global cache for master database (loaded once on startup)
MASTER_DB_CACHE = None

if not os.path.exists('static'):
    os.makedirs('static')
if not os.path.exists(QR_FOLDER):
    os.makedirs(QR_FOLDER)

# Initialize database on startup
db_manager.init_db()


# ─────────────────────────────────────────────
#  AUTH DECORATORS
# ─────────────────────────────────────────────
def require_admin(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('admin_verified'):
            return jsonify({'status': 'error', 'message': 'Admin PIN required.'}), 403
        return f(*args, **kwargs)
    return decorated


# ─────────────────────────────────────────────
#  HELPER FUNCTIONS
# ─────────────────────────────────────────────
def generate_6_digit_code():
    chars = string.ascii_uppercase + string.digits
    chars = chars.replace('O', '').replace('0', '').replace('I', '').replace('1', '')
    return ''.join(random.choices(chars, k=6))


def load_master_db():
    global MASTER_DB_CACHE
    if MASTER_DB_CACHE is not None:
        return MASTER_DB_CACHE
    try:
        MASTER_DB_CACHE = pd.read_excel(MASTER_DB)
        if not all(col in MASTER_DB_CACHE.columns for col in ['Department', 'School/Office', 'Employee_Name']):
            raise FileNotFoundError("Wrong columns in master_db.xlsx")
        return MASTER_DB_CACHE
    except Exception as e:
        print(f"Error loading master database: {e}")
        data = {
            'Department': ['Division Office', 'Elementary Schools', 'Secondary Schools', 'Senior High Schools'],
            'School/Office': ['Main Office', 'Riverside Elementary', 'Westside High', 'North Academy'],
            'Employee_Name': ['John Doe', 'Jane Smith', 'Mike Ross', 'Sarah Lee']
        }
        df = pd.DataFrame(data)
        df.to_excel(MASTER_DB, index=False)
        MASTER_DB_CACHE = df
        return df


def save_code_lookup(code, department, school, employees, doc_type):
    conn = db_manager.get_db_connection()
    conn.execute(
        "INSERT OR REPLACE INTO code_lookup (code, department, school_office, employees, doc_type, generated_at) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (code, department, school, ','.join(employees), doc_type, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    )
    conn.commit()
    conn.close()
    db_manager.trigger_excel_export()


def get_code_lookup(code):
    conn = db_manager.get_db_connection()
    row = conn.execute("SELECT * FROM code_lookup WHERE code = ?", (code,)).fetchone()
    conn.close()
    if not row:
        return None
    return {
        'department': row['department'],
        'school':     row['school_office'],
        'employees':  row['employees'].split(','),
        'doc_type':   row['doc_type']
    }


# ─────────────────────────────────────────────
#  PAGE ROUTES
# ─────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/scanner')
def scanner_page():
    return render_template('scanner.html')


# ─────────────────────────────────────────────
#  AUTH ENDPOINTS
# ─────────────────────────────────────────────
@app.route('/api/verify_pin', methods=['POST'])
def verify_pin():
    try:
        data      = request.json
        pin_type  = data.get('pin_type')   # 'scanner' or 'admin'
        pin       = data.get('pin', '').strip()

        if pin_type == 'scanner':
            stored = db_manager.get_setting('scanner_pin', 'scanner123')
            if pin == stored:
                session['scanner_verified'] = True
                session.permanent = True
                return jsonify({'status': 'success', 'message': 'Scanner access granted.'})
            return jsonify({'status': 'error', 'message': 'Incorrect scanner PIN.'})

        elif pin_type == 'admin':
            stored = db_manager.get_setting('admin_pin', 'admin123')
            if pin == stored:
                session['admin_verified'] = True
                session.permanent = True
                return jsonify({'status': 'success', 'message': 'Admin access granted.'})
            return jsonify({'status': 'error', 'message': 'Incorrect admin PIN.'})

        return jsonify({'status': 'error', 'message': 'Invalid pin_type.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/check_auth', methods=['GET'])
def check_auth():
    return jsonify({
        'scanner_verified': bool(session.get('scanner_verified')),
        'admin_verified':   bool(session.get('admin_verified'))
    })


@app.route('/api/logout', methods=['POST'])
def logout():
    pin_type = (request.json or {}).get('pin_type', 'all')
    if pin_type == 'scanner':
        session.pop('scanner_verified', None)
    elif pin_type == 'admin':
        session.pop('admin_verified', None)
    else:
        session.clear()
    return jsonify({'status': 'success'})


# ─────────────────────────────────────────────
#  MASTER DATA APIs
# ─────────────────────────────────────────────
@app.route('/api/get_departments')
def get_departments():
    df = load_master_db()
    return jsonify(sorted(df['Department'].dropna().unique().tolist()))


@app.route('/api/get_schools', methods=['POST'])
def get_schools():
    df   = load_master_db()
    dept = request.json.get('department')
    filtered = df[df['Department'] == dept]
    return jsonify(sorted(filtered['School/Office'].dropna().unique().tolist()))


@app.route('/api/get_employees', methods=['POST'])
def get_employees():
    data = request.json
    df   = load_master_db()
    filtered = df[
        (df['Department'] == data.get('department')) &
        (df['School/Office'] == data.get('school'))
    ]
    return jsonify(sorted(filtered['Employee_Name'].dropna().unique().tolist()))


@app.route('/api/get_all_schools', methods=['GET'])
def get_all_schools():
    """All unique schools/offices for the scanner receiving-office picker."""
    try:
        df      = load_master_db()
        schools = sorted(df['School/Office'].dropna().unique().tolist())
        return jsonify({'status': 'success', 'schools': schools})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/get_all_school_employees', methods=['POST'])
def get_all_school_employees():
    """Return ALL employees belonging to a specific school (for Add-All feature)."""
    try:
        data   = request.json
        dept   = data.get('department', '').strip()
        school = data.get('school', '').strip()
        df     = load_master_db()

        if dept and school:
            filtered = df[(df['Department'] == dept) & (df['School/Office'] == school)]
        elif school:
            filtered = df[df['School/Office'] == school]
        else:
            return jsonify({'status': 'error', 'message': 'School/Office is required.'})

        employees = sorted(filtered['Employee_Name'].dropna().unique().tolist())
        return jsonify({'status': 'success', 'employees': employees, 'count': len(employees)})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ─────────────────────────────────────────────
#  CODE GENERATION
# ─────────────────────────────────────────────
@app.route('/generate_code', methods=['POST'])
def generate_code():
    try:
        data       = request.json
        code       = generate_6_digit_code()
        code_type  = data.get('code_type', 'barcode')
        department = data.get('department')
        school     = data.get('school')
        employees  = data.get('employees')
        doc_type   = data.get('doc_type')

        save_code_lookup(code, department, school, employees, doc_type)

        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        filename  = f"code_{timestamp}"
        filepath  = os.path.join(QR_FOLDER, filename)

        if code_type == 'barcode':
            try:
                from barcode import Code128
                from barcode.writer import ImageWriter
                barcode_obj = Code128(code, writer=ImageWriter())
                barcode_obj.save(filepath)
            except Exception as e:
                print(f"Barcode lib failed: {e}. Falling back to QR.")
                code_type = 'qr'

        if code_type == 'qr':
            qr = qrcode.QRCode(version=1, box_size=10, border=4)
            qr.add_data(code)
            qr.make(fit=True)
            img = qr.make_image(fill="black", back_color="white")
            img.save(f"{filepath}.png")

        return jsonify({
            'status':        'success',
            'image_url':     f"/static/qr_generated/{filename}.png",
            'code_display':  code,
            'generated_type': code_type
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ─────────────────────────────────────────────
#  SCAN LOGGING
# ─────────────────────────────────────────────
@app.route('/log_scan', methods=['POST'])
def log_scan():
    try:
        data             = request.json
        code             = data.get('scanned_data', '').strip().upper()
        receiving_office = data.get('receiving_office', '').strip()
        receiver_name    = data.get('receiver_name', '').strip()

        if not receiving_office:
            return jsonify({'status': 'error', 'message': 'Please select a Receiving Office!'})

        lookup = get_code_lookup(code)
        if not lookup:
            return jsonify({'status': 'error', 'message': f'Barcode "{code}" not found in database!'})

        employees = [e.strip() for e in lookup['employees'] if e.strip()]
        if not employees:
            return jsonify({'status': 'error', 'message': 'No employees linked to this barcode!'})

        conn = db_manager.get_db_connection()

        # ── 1. ACCIDENTAL DOUBLE-SCAN CHECK ──────────────────────────────────
        for emp in employees:
            row = conn.execute(
                "SELECT * FROM routing_records WHERE code = ? AND employee = ?", (code, emp)
            ).fetchone()
            if row:
                office_cols = [c for c in row.keys() if c.startswith('receiving_office_')]
                stages = sorted({int(c.split('_')[-1]) for c in office_cols if c.split('_')[-1].isdigit()})
                if not stages:
                    stages = list(range(1, 11))

                latest_office = latest_ts = None
                for i in sorted(stages, reverse=True):
                    key = f'receiving_office_{i}'
                    if key in row.keys() and row[key]:
                        latest_office = row[key]
                        latest_ts     = row[f'timestamp_{i}']
                        break

                if latest_office and latest_office.strip().lower() == receiving_office.lower():
                    conn.close()
                    return jsonify({
                        'status':  'warning',
                        'message': f'Duplicate scan! Already received at "{latest_office}" on {latest_ts}.'
                    })

        # ── 2. LOG THE ROUTING STAMP ──────────────────────────────────────────
        current_ts       = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        updated_rows_ids = []
        slot_updated     = 1

        for emp in employees:
            row = conn.execute(
                "SELECT * FROM routing_records WHERE code = ? AND employee = ?", (code, emp)
            ).fetchone()

            if not row:
                # First scan — ensure stage-1 columns exist, then insert
                db_manager.ensure_routing_columns(conn, 1)
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO routing_records "
                    "(department, school_office, employee, code, receiving_office_1, receiver_name_1, timestamp_1) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (lookup['department'], lookup['school'], emp, code,
                     receiving_office, receiver_name, current_ts)
                )
                conn.commit()
                updated_rows_ids.append(cursor.lastrowid)
                slot_updated = 1

            else:
                record_id   = row['id']
                office_cols = [c for c in row.keys() if c.startswith('receiving_office_')]
                stages = sorted({int(c.split('_')[-1]) for c in office_cols if c.split('_')[-1].isdigit()})
                if not stages:
                    stages = list(range(1, 11))

                empty_slot = None
                for i in stages:
                    key = f'receiving_office_{i}'
                    if key in row.keys() and not row[key]:
                        empty_slot = i
                        break

                if empty_slot is None:
                    # All slots full — expand schema dynamically
                    next_slot = max(stages) + 1
                    db_manager.ensure_routing_columns(conn, next_slot)
                    empty_slot = next_slot

                conn.execute(
                    f"UPDATE routing_records "
                    f"SET receiving_office_{empty_slot}=?, receiver_name_{empty_slot}=?, timestamp_{empty_slot}=? "
                    f"WHERE id=?",
                    (receiving_office, receiver_name, current_ts, record_id)
                )
                conn.commit()
                updated_rows_ids.append(record_id)
                slot_updated = empty_slot

        conn.close()

        for r_id in updated_rows_ids:
            db_manager.queue_sync(r_id)
        db_manager.trigger_excel_export()

        return jsonify({
            'status':    'success',
            'message':   f'Tracked {len(employees)} employee(s) at "{receiving_office}"!',
            'data':      employees,
            'school':    lookup['school'],
            'doc_type':  lookup['doc_type'],
            'timestamp': current_ts,
            'slot':      slot_updated
        })

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ─────────────────────────────────────────────
#  ROUTING RECORDS — PAGINATED
# ─────────────────────────────────────────────
@app.route('/api/get_routing_records', methods=['GET'])
def get_routing_records():
    try:
        page     = int(request.args.get('page', 1))
        per_page = int(request.args.get('per_page', 50))
        search   = request.args.get('search', '').strip()
        sort_by  = request.args.get('sort_by', 'id')
        order    = request.args.get('order', 'DESC').upper()

        # Sanitise sort params
        allowed_sort = {'id', 'department', 'school_office', 'employee', 'code'}
        if sort_by not in allowed_sort:
            sort_by = 'id'
        if order not in ('ASC', 'DESC'):
            order = 'DESC'

        conn = db_manager.get_db_connection()

        where_sql = ""
        params    = []
        if search:
            where_sql = ("WHERE (employee LIKE ? OR code LIKE ? "
                         "OR department LIKE ? OR school_office LIKE ?)")
            pat    = f"%{search}%"
            params = [pat, pat, pat, pat]

        total = conn.execute(
            f"SELECT COUNT(*) FROM routing_records {where_sql}", params
        ).fetchone()[0]

        offset = (page - 1) * per_page
        rows = conn.execute(
            f"SELECT * FROM routing_records {where_sql} "
            f"ORDER BY {sort_by} {order} LIMIT ? OFFSET ?",
            params + [per_page, offset]
        ).fetchall()

        conn.close()

        records     = [dict(r) for r in rows]
        total_pages = max(1, (total + per_page - 1) // per_page)

        return jsonify({
            'status':  'success',
            'records': records,
            'pagination': {
                'page':        page,
                'per_page':    per_page,
                'total':       total,
                'total_pages': total_pages
            }
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ─────────────────────────────────────────────
#  EDIT / DELETE RECORDS (Admin only)
# ─────────────────────────────────────────────
@app.route('/api/update_record/<int:record_id>', methods=['PUT'])
@require_admin
def update_record(record_id):
    try:
        data = request.json
        conn = db_manager.get_db_connection()

        # Fetch valid DB columns
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(routing_records)")
        db_cols = {row[1] for row in cursor.fetchall()}

        set_parts, vals = [], []
        for field, value in data.items():
            if field in db_cols and field != 'id':
                set_parts.append(f"{field} = ?")
                vals.append(value)

        if not set_parts:
            conn.close()
            return jsonify({'status': 'error', 'message': 'No valid fields to update.'})

        vals.append(record_id)
        conn.execute(f"UPDATE routing_records SET {', '.join(set_parts)} WHERE id = ?", vals)
        conn.commit()
        conn.close()

        db_manager.trigger_excel_export()
        return jsonify({'status': 'success', 'message': 'Record updated.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/delete_record/<int:record_id>', methods=['DELETE'])
@require_admin
def delete_record(record_id):
    try:
        conn = db_manager.get_db_connection()
        conn.execute("DELETE FROM routing_records WHERE id = ?", (record_id,))
        conn.commit()
        conn.close()
        db_manager.trigger_excel_export()
        return jsonify({'status': 'success', 'message': 'Record deleted.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ─────────────────────────────────────────────
#  SETTINGS (Admin only)
# ─────────────────────────────────────────────
@app.route('/api/get_settings', methods=['GET'])
@require_admin
def get_settings():
    try:
        return jsonify({
            'status':              'success',
            'gsheets_enabled':     db_manager.get_setting('gsheets_enabled', 'False') == 'True',
            'gsheets_id':          db_manager.get_setting('gsheets_id', ''),
            'gsheets_credentials': db_manager.get_setting('gsheets_credentials', ''),
            'scanner_pin':         db_manager.get_setting('scanner_pin', 'scanner123'),
            'admin_pin':           db_manager.get_setting('admin_pin', 'admin123'),
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/save_settings', methods=['POST'])
@require_admin
def save_settings():
    try:
        data = request.json
        db_manager.save_setting('gsheets_enabled', 'True' if data.get('gsheets_enabled') else 'False')
        db_manager.save_setting('gsheets_id',          data.get('gsheets_id', '').strip())
        db_manager.save_setting('gsheets_credentials', data.get('gsheets_credentials', '').strip())

        if data.get('scanner_pin', '').strip():
            db_manager.save_setting('scanner_pin', data['scanner_pin'].strip())
        if data.get('admin_pin', '').strip():
            db_manager.save_setting('admin_pin', data['admin_pin'].strip())

        return jsonify({'status': 'success', 'message': 'Settings saved!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/test_sheets', methods=['POST'])
def test_sheets():
    try:
        data     = request.json
        sheet_id = data.get('gsheets_id', '').strip()
        creds    = data.get('gsheets_credentials', '').strip()
        if not sheet_id or not creds:
            return jsonify({'status': 'error', 'message': 'Missing Sheet ID or credentials JSON.'})
        result = db_manager.test_google_sheets_connection(sheet_id, creds)
        return jsonify(result)
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/reload_excel', methods=['POST'])
def reload_excel():
    """Force reload data from Excel - preserves manual Excel edits."""
    try:
        count = db_manager.force_import_from_excel()
        return jsonify({'status': 'success', 'message': f'Loaded {count} new records from Excel!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/gsheet_status', methods=['GET'])
def gsheet_status():
    """Quick-test the currently saved Google Sheets connection."""
    try:
        enabled  = db_manager.get_setting('gsheets_enabled', 'False')
        sheet_id = db_manager.get_setting('gsheets_id', '')
        creds    = db_manager.get_setting('gsheets_credentials', '')
        if enabled != 'True' or not sheet_id or not creds:
            return jsonify({'status': 'disabled', 'message': 'Google Sheets sync is not configured.'})
        result = db_manager.test_google_sheets_connection(sheet_id, creds)
        return jsonify(result)
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ─────────────────────────────────────────────
#  STARTUP
# ─────────────────────────────────────────────
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)