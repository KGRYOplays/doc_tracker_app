import os
import io
import time
import sqlite3
import pandas as pd
import qrcode
import random
import string
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, jsonify, session, send_file, redirect, url_for
from flask_cors import CORS
from PIL import Image, ImageDraw, ImageFont
import hashlib
import json as json_lib
import threading
import bcrypt
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

import db_manager

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY')
if not app.secret_key:
    app.secret_key = 'dev-fallback-key-do-not-use-in-production'
    print("WARNING: SECRET_KEY env var not set. Using insecure fallback. Set SECRET_KEY in production.")
CORS(app)

limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["60 per minute"],
    storage_uri="memory://"
)

MASTER_DB  = 'master_db.xlsx'
QR_FOLDER  = os.path.join('static', 'qr_generated')

if not os.path.exists('static'):
    os.makedirs('static')
if not os.path.exists(QR_FOLDER):
    os.makedirs(QR_FOLDER)

# Initialize database on startup
db_manager.init_db()

# Seed default SDO Manila logo if none set
if not db_manager.get_setting('logo_url', ''):
    logo_path = 'static/uploads/sdo_manila_logo.png'
    if os.path.exists(logo_path):
        db_manager.save_setting('logo_url', '/static/uploads/sdo_manila_logo.png')

# Serialize concurrent scan writes so SELECT → INSERT/UPDATE is atomic
scan_lock = threading.Lock()

# ── Graceful shutdown: drain GSheets sync queue before exit ──
import signal as _signal
import atexit as _atexit

def _drain_sync_queue():
    if hasattr(db_manager, 'gs_sync_queue'):
        print("Shutdown: draining GSheets sync queue...")
        try:
            db_manager.gs_sync_queue.join()
            print("Shutdown: queue drained.")
        except Exception as e:
            print(f"Shutdown: error draining queue: {e}")

_signal.signal(_signal.SIGTERM, lambda *_: (_drain_sync_queue(), exit(0)))
_atexit.register(_drain_sync_queue)
print("Graceful shutdown handler registered.")


# ─────────────────────────────────────────────
#  AUTH DECORATORS
# ─────────────────────────────────────────────
def require_admin(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if session.get('user_role') != 'Admin':
            return jsonify({'status': 'error', 'message': 'Admin privileges required.'}), 403
        return f(*args, **kwargs)
    return decorated

def require_role(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if session.get('user_role') not in roles:
                return jsonify({'status': 'error', 'message': f'Access restricted. Requires roles: {", ".join(roles)}'}), 403
            return f(*args, **kwargs)
        return decorated
    return decorator

def require_login(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('username'):
            return jsonify({'status': 'error', 'message': 'Login required.'}), 401
        return f(*args, **kwargs)
    return decorated


# ─────────────────────────────────────────────
#  HELPER FUNCTIONS
# ─────────────────────────────────────────────
def generate_6_digit_code():
    chars = string.ascii_uppercase + string.digits
    chars = chars.replace('O', '').replace('0', '').replace('I', '').replace('1', '')
    return ''.join(random.choices(chars, k=6))


def _get_gsheets_config():
    """Return (is_enabled, sheet_id, creds) tuple. Creds are decrypted.
    Returns (False, None, None) during maintenance mode to block sync pushes."""
    if db_manager.get_setting('maintenance_mode') == 'True':
        return False, None, None
    enabled = db_manager.get_setting('gsheets_enabled') == 'True'
    sheet_id = db_manager.get_setting('gsheets_id')
    creds = db_manager.get_decrypted_setting('gsheets_credentials')
    return enabled and bool(sheet_id) and bool(creds), sheet_id, creds


def _check_maintenance():
    """Check if maintenance mode is active and the current user is not admin.
    Returns a (blocked, message_or_None) tuple.
    If blocked, the caller should return the message as a 403 response."""
    if db_manager.get_setting('maintenance_mode') == 'True':
        role = session.get('user_role', '')
        if role.lower() != 'admin':
            return True, 'System is in maintenance mode. Only admins can perform this action.'
    return False, None


# Master data is now stored in the SQLite `master_data` table.
# See db_manager.py for helper functions: get_all_departments(),
# get_schools_for_department(), get_employees_for_school(), etc.


def _resolve_employee_ids(conn, employee_names, school):
    """Look up (or create on the fly) employee IDs for a list of names + school.
    Returns list of (name, employee_id) tuples."""
    ids = []
    s = conn.execute("SELECT id FROM schools WHERE name = ?", (school,)).fetchone()
    if not s:
        # School not in schools table yet — create it
        conn.execute("INSERT OR IGNORE INTO schools (name) VALUES (?)", (school,))
        s = conn.execute("SELECT id FROM schools WHERE name = ?", (school,)).fetchone()
    school_id = s['id']
    for name in employee_names:
        e = conn.execute(
            "SELECT id FROM employees WHERE name = ? AND school_id = ?",
            (name, school_id)
        ).fetchone()
        if not e:
            conn.execute(
                "INSERT OR IGNORE INTO employees (name, school_id) VALUES (?, ?)",
                (name, school_id)
            )
            e = conn.execute(
                "SELECT id FROM employees WHERE name = ? AND school_id = ?",
                (name, school_id)
            ).fetchone()
        if e:
            ids.append((name, e['id']))
    return ids


def save_code_lookup(code, department, school, employees, doc_type):
    gen_time = datetime.now().strftime("%m/%d/%Y")
    employees = [db_manager.reorder_name(e) for e in employees]
    
    # ── 1. SQLite-first: write to local DB first (fast, always succeeds) ──
    import time as _time
    import sqlite3 as _sqlite3
    for _attempt in range(3):
        try:
            conn = db_manager.get_db_connection()
            conn.execute(
                "INSERT OR REPLACE INTO code_lookup (code, department, school_office, employees, doc_type, generated_at) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (code, department, school, '|||'.join(employees), doc_type, gen_time)
            )
            # Dual-write to code_employees
            emp_ids = _resolve_employee_ids(conn, employees, school)
            for _name, eid in emp_ids:
                conn.execute(
                    "INSERT OR IGNORE INTO code_employees (code, employee_id) VALUES (?, ?)",
                    (code, eid)
                )
            conn.commit()
            conn.close()
            break
        except _sqlite3.OperationalError as _e:
            if 'locked' in str(_e) and _attempt < 2:
                _time.sleep(1 * (_attempt + 1))
                continue
            raise
    db_manager.trigger_excel_export()
    
    # ── 2. Enqueue for async GSheets push (non-blocking, zero API cost) ──
    db_manager.enqueue_sync('code_lookup', code=code)


def save_code_lookup_batch(code_dicts):
    """Batch save multiple codes — insert into SQLite first, then push to GSheets."""
    gen_time = datetime.now().strftime("%m/%d/%Y")
    
    # ── 1. SQLite-first: insert all new codes into local DB ──
    import time as _time
    import sqlite3 as _sqlite3
    for _attempt in range(3):
        try:
            conn = db_manager.get_db_connection()
            conn.executemany(
                "INSERT OR REPLACE INTO code_lookup (code, department, school_office, employees, doc_type, generated_at) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                [(
                    cd['code'], cd['department'], cd['school_office'],
                    cd['employees'], cd['doc_type'], gen_time
                ) for cd in code_dicts]
            )
            # Dual-write to code_employees
            for cd in code_dicts:
                emp_names = [e.strip() for e in cd['employees'].split('|||') if e.strip()]
                emp_ids = _resolve_employee_ids(conn, emp_names, cd['school_office'])
                for _name, eid in emp_ids:
                    conn.execute(
                        "INSERT OR IGNORE INTO code_employees (code, employee_id) VALUES (?, ?)",
                        (cd['code'], eid)
                    )
            conn.commit()
            conn.close()
            break
        except _sqlite3.OperationalError as _e:
            if 'locked' in str(_e) and _attempt < 2:
                _time.sleep(1 * (_attempt + 1))
                continue
            raise
    db_manager.trigger_excel_export()
    
    # ── 2. Enqueue all new codes for async GSheets push ──
    for cd in code_dicts:
        db_manager.enqueue_sync('code_lookup', code=cd['code'])


def get_code_lookup(code):
    conn = db_manager.get_db_connection()
    row = conn.execute("SELECT * FROM code_lookup WHERE code = ?", (code,)).fetchone()
    if not row:
        conn.close()
        return None
    
    # Try RDBMS first: read from code_employees + employees
    emp_rows = conn.execute("""
        SELECT e.name FROM code_employees ce
        JOIN employees e ON e.id = ce.employee_id
        WHERE ce.code = ?
    """, (code,)).fetchall()
    
    if emp_rows:
        employees_list = [db_manager.reorder_name(r['name']) for r in emp_rows]
    else:
        # Fallback to ||| splitting for backward compat (pre-RDBMS data)
        employees_str = row['employees']
        if '|||' in employees_str:
            employees_list = employees_str.split('|||')
        else:
            employees_list = employees_str.split(',')
        employees_list = [db_manager.reorder_name(e) for e in employees_list]
    
    conn.close()
    return {
        'department': row['department'],
        'school':     row['school_office'],
        'employees':  employees_list,
        'doc_type':   row['doc_type']
    }


# ─────────────────────────────────────────────
#  BARCODE / LABEL GENERATION HELPERS
# ─────────────────────────────────────────────
def generate_barcode_image(code):
    """Generate a barcode PIL image, 1.5 in wide x 0.75 in tall at 300 DPI."""
    from barcode import Code128
    from barcode.writer import ImageWriter
    DPI = 300
    target_w = int(1.5 * DPI)
    target_h = int(0.75 * DPI)

    writer_options = {
        'module_width': 0.35,
        'module_height': 15.0,
        'font_size': 10,
        'text_distance': 3,
        'quiet_zone': 1.0,
        'write_text': True,
        'background': 'white',
        'foreground': 'black',
    }
    writer = ImageWriter()
    barcode_obj = Code128(code, writer=writer)
    buf = io.BytesIO()
    barcode_obj.write(buf, writer_options)
    buf.seek(0)
    img = Image.open(buf).convert('RGB')
    img = img.resize((target_w, target_h), Image.LANCZOS)

    border_px = 6
    boxed_w = target_w + border_px * 2
    boxed_h = target_h + border_px * 2 + 24
    boxed = Image.new('RGB', (boxed_w, boxed_h), 'white')
    draw = ImageDraw.Draw(boxed)
    draw.rectangle([0, 0, boxed_w - 1, boxed_h - 1], outline='black', width=2)
    paste_x = (boxed_w - target_w) // 2
    paste_y = border_px
    boxed.paste(img, (paste_x, paste_y))
    try:
        font = ImageFont.truetype("arial.ttf", 16)
    except:
        font = ImageFont.load_default()
    text = f"{code}"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_x = (boxed_w - (bbox[2] - bbox[0])) // 2
    text_y = boxed_h - 20
    draw.text((text_x, text_y), text, fill='black', font=font)
    return boxed


def generate_qr_image(code):
    """Generate a QR code PIL image, 1.5 in x 1.5 in square at 300 DPI."""
    qr = qrcode.QRCode(version=1, box_size=10, border=2)
    qr.add_data(code)
    qr.make(fit=True)
    img = qr.make_image(fill="black", back_color="white").convert('RGB')
    DPI = 300
    target = int(1.5 * DPI)
    img = img.resize((target, target), Image.LANCZOS)

    border_px = 6
    boxed_w = target + border_px * 2
    boxed_h = target + border_px * 2 + 24
    boxed = Image.new('RGB', (boxed_w, boxed_h), 'white')
    draw = ImageDraw.Draw(boxed)
    draw.rectangle([0, 0, boxed_w - 1, boxed_h - 1], outline='black', width=2)
    paste_x = (boxed_w - target) // 2
    paste_y = border_px
    boxed.paste(img, (paste_x, paste_y))
    try:
        font = ImageFont.truetype("arial.ttf", 16)
    except:
        font = ImageFont.load_default()
    text = f"{code}"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_x = (boxed_w - (bbox[2] - bbox[0])) // 2
    text_y = boxed_h - 20
    draw.text((text_x, text_y), text, fill='black', font=font)
    return boxed


def create_label_pdf(codes, dpi=300):
    """
    Create a PDF with 4 columns x 5 rows of combined QR + barcode labels
    on 8.5 x 13 (Legal) paper. Each cell shows the QR (1.5" sq) on top
    and the Code128 barcode (1.5" x 0.75") below.
    """
    from reportlab.lib.pagesizes import legal
    from reportlab.lib.units import inch
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    PAGE_W, PAGE_H = legal  # 8.5 x 13 in points
    COLS, ROWS = 4, 5
    MARGIN = 0.4 * inch
    GAP = 0.15 * inch

    usable_w = PAGE_W - 2 * MARGIN
    usable_h = PAGE_H - 2 * MARGIN
    cell_w = (usable_w - (COLS - 1) * GAP) / COLS
    cell_h = (usable_h - (ROWS - 1) * GAP) / ROWS
    pad_x = (usable_w - (COLS * cell_w + (COLS - 1) * GAP)) / 2
    pad_y = (usable_h - (ROWS * cell_h + (ROWS - 1) * GAP)) / 2

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=legal)

    for idx, item in enumerate(codes):
        code = item.get('code', '')

        if idx > 0 and idx % (COLS * ROWS) == 0:
            c.showPage()

        pos = idx % (COLS * ROWS)
        col = pos % COLS
        row = pos // COLS

        cx = MARGIN + pad_x + col * (cell_w + GAP)
        cy = MARGIN + pad_y + row * (cell_h + GAP)

        qr_pil = generate_qr_image(code)
        bc_pil = generate_barcode_image(code)

        # Scale QR to fit in upper ~60% of cell
        qr_max_w = cell_w * 0.85
        qr_max_h = cell_h * 0.55
        qr_scale = min(qr_max_w / qr_pil.width, qr_max_h / qr_pil.height)
        qr_disp_w = qr_pil.width * qr_scale
        qr_disp_h = qr_pil.height * qr_scale
        qr_x = cx + (cell_w - qr_disp_w) / 2
        qr_y = cy + cell_h * 0.42 + (cell_h * 0.58 - qr_disp_h) / 2

        # Scale barcode to fit in lower ~40% of cell
        bc_max_w = cell_w * 0.85
        bc_max_h = cell_h * 0.35
        bc_scale = min(bc_max_w / bc_pil.width, bc_max_h / bc_pil.height)
        bc_disp_w = bc_pil.width * bc_scale
        bc_disp_h = bc_pil.height * bc_scale
        bc_x = cx + (cell_w - bc_disp_w) / 2
        bc_y = cy + (cell_h * 0.42 - bc_disp_h) / 2

        def _draw(pil_img, x, y, w, h):
            tmp = io.BytesIO()
            pil_img.save(tmp, format='PNG')
            tmp.seek(0)
            c.drawImage(ImageReader(tmp), x, y, width=w, height=h)

        _draw(qr_pil, qr_x, qr_y, qr_disp_w, qr_disp_h)
        _draw(bc_pil, bc_x, bc_y, bc_disp_w, bc_disp_h)

        # Dashed cell border
        c.setStrokeColorRGB(0.7, 0.7, 0.7)
        c.setLineWidth(0.5)
        c.rect(cx, cy, cell_w, cell_h)
        c.setStrokeColorRGB(0, 0, 0)
        c.setLineWidth(1)

    c.save()
    buf.seek(0)
    return buf


def create_document_page_pdf(code):
    """
    Create a PDF with QR + barcode placed at the TOP RIGHT corner of a document.
    The page looks like a formal document with header area.
    """
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
    from reportlab.lib.colors import HexColor
    
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    width, height = letter  # 612 x 792 points
    
    # ── Generate both images ──
    qr_img = generate_qr_image(code)
    bc_img = generate_barcode_image(code)
    gap = 0.1 * inch
    
    # ── Place both at TOP RIGHT corner ──
    img_w_inches = 1.5
    img_w = img_w_inches * inch
    qr_aspect = qr_img.height / qr_img.width
    bc_aspect = bc_img.height / bc_img.width
    qr_h = img_w * qr_aspect
    bc_h = img_w * bc_aspect
    total_h = qr_h + gap + bc_h
    
    margin = 0.5 * inch
    x = width - margin - img_w
    y = height - margin - total_h
    
    def _draw(pil_img, draw_x, draw_y, dw, dh):
        tmp = io.BytesIO()
        pil_img.save(tmp, format='PNG')
        tmp.seek(0)
        c.drawImage(ImageReader(tmp), draw_x, draw_y, width=dw, height=dh)
    
    _draw(qr_img, x, y + bc_h + gap, img_w, qr_h)
    _draw(bc_img, x, y, img_w, bc_h)
    
    # ── Document header / title ──
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin, height - 0.8 * inch, "DOCUMENT TRACKING SLIP")
    
    # ── Document info lines ──
    c.setFont("Helvetica", 11)
    y_pos = height - 1.3 * inch
    c.drawString(margin, y_pos, f"Tracking Code: {code}")
    
    # Look up info if available
    lookup = get_code_lookup(code)
    if lookup:
        y_pos -= 18
        c.drawString(margin, y_pos, f"Department: {lookup['department']}")
        y_pos -= 18
        c.drawString(margin, y_pos, f"School/Office: {lookup['school']}")
        y_pos -= 18
        employees_str = ', '.join(lookup['employees'])
        c.drawString(margin, y_pos, f"Employee(s): {employees_str}")
        y_pos -= 18
        c.drawString(margin, y_pos, f"Document Type: {lookup['doc_type']}")
    
    # ── Separator line ──
    y_pos -= 20
    c.setStrokeColor(HexColor('#6366f1'))
    c.setLineWidth(1.5)
    c.line(margin, y_pos, width - margin, y_pos)
    
    # ── Routing table header ──
    y_pos -= 25
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margin, y_pos, "ROUTING HISTORY")
    
    c.setFont("Helvetica-Bold", 9)
    y_pos -= 18
    table_margin = margin + 10
    col_w = (width - 2 * table_margin) / 5
    col_headers = ['#', 'Receiving Office', 'Receiver', 'Date/Time', 'Status']
    
    # Draw table headers
    c.setFillColor(HexColor('#1e1b4b'))
    for i, hdr in enumerate(col_headers):
        c.drawString(table_margin + i * col_w + 5, y_pos, hdr)
    
    # Draw header underline
    y_pos -= 3
    c.setStrokeColorRGB(0.8, 0.8, 0.8)
    c.line(table_margin, y_pos, width - table_margin, y_pos)
    y_pos -= 12
    
    # ── Fetch routing data ──
    try:
        conn = db_manager.get_db_connection()
        employees = lookup['employees'] if lookup else []
        
        # Collect routing records for this code
        all_routes = {}
        for emp in employees:
            row = conn.execute(
                "SELECT * FROM routing_records WHERE code = ? AND employee = ?", (code, emp)
            ).fetchone()
            if row:
                office_cols = [c for c in row.keys() if c.startswith('receiving_office_')]
                stages = sorted({int(c.split('_')[-1]) for c in office_cols if c.split('_')[-1].isdigit()})
                for i in stages:
                    off_key = f'receiving_office_{i}'
                    rec_key = f'receiver_name_{i}'
                    ts_key  = f'timestamp_{i}'
                    office = row[off_key] if off_key in row.keys() else ''
                    receiver = row[rec_key] if rec_key in row.keys() else ''
                    ts = row[ts_key] if ts_key in row.keys() else ''
                    if office:
                        if office not in all_routes:
                            all_routes[office] = {'receiver': receiver, 'timestamp': ts, 'stage': i}
        
        conn.close()
        
        # Draw routing entries
        c.setFont("Helvetica", 9)
        c.setFillColorRGB(0.1, 0.1, 0.1)
        for idx, (office, data) in enumerate(all_routes.items()):
            if y_pos < 40:  # Don't go below page margin
                c.showPage()
                y_pos = height - margin
                c.setFont("Helvetica", 9)
            
            row_data = [str(idx + 1), office, data['receiver'], data['timestamp'], '✓ Received']
            for i, val in enumerate(row_data):
                c.drawString(table_margin + i * col_w + 5, y_pos, str(val))
            y_pos -= 16
            
            if idx < len(all_routes) - 1:
                c.setStrokeColorRGB(0.9, 0.9, 0.9)
                c.setLineWidth(0.5)
                c.line(table_margin, y_pos + 8, width - table_margin, y_pos + 8)
                c.setStrokeColorRGB(0, 0, 0)
                c.setLineWidth(1)
        
        if not all_routes:
            c.setFont("Helvetica-Oblique", 10)
            c.setFillColorRGB(0.5, 0.5, 0.5)
            c.drawString(table_margin, y_pos, "No routing history yet. Scan the barcode at receiving offices to track the document.")
            c.setFillColorRGB(0.1, 0.1, 0.1)
            
    except Exception as e:
        print(f"Error fetching routing data: {e}")
    
    c.save()
    buf.seek(0)
    return buf


def create_barcode_overlay_pdf(code, position='top-right', page_size='legal'):
    """
    Create a minimal PDF with QR + barcode at the specified position.
    No headers, no routing table - just the labels on a blank page.
    Useful for printing onto an already-printed document.
    
    Supported positions: top-right, top-left, bottom-right, bottom-left
    Supported page sizes: legal (8.5x13), letter (8.5x11)
    """
    from reportlab.lib.pagesizes import letter, legal
    from reportlab.lib.units import inch
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
    
    if page_size == 'legal':
        pagesize = legal
    else:
        pagesize = letter
    
    width, height = pagesize
    
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=pagesize)
    
    qr_img = generate_qr_image(code)
    bc_img = generate_barcode_image(code)
    gap = 0.1 * inch
    
    img_w_inches = 1.5
    img_w = img_w_inches * inch
    qr_aspect = qr_img.height / qr_img.width
    bc_aspect = bc_img.height / bc_img.width
    qr_h = img_w * qr_aspect
    bc_h = img_w * bc_aspect
    total_h = qr_h + gap + bc_h
    
    margin = 0.5 * inch
    
    if position == 'top-right':
        x = width - margin - img_w
        y = height - margin - total_h
    elif position == 'top-left':
        x = margin
        y = height - margin - total_h
    elif position == 'bottom-right':
        x = width - margin - img_w
        y = margin
    elif position == 'bottom-left':
        x = margin
        y = margin
    else:
        x = width - margin - img_w
        y = height - margin - total_h
    
    def _draw(pil_img, draw_x, draw_y, dw, dh):
        tmp = io.BytesIO()
        pil_img.save(tmp, format='PNG')
        tmp.seek(0)
        c.drawImage(ImageReader(tmp), draw_x, draw_y, width=dw, height=dh)
    
    _draw(qr_img, x, y + bc_h + gap, img_w, qr_h)
    _draw(bc_img, x, y, img_w, bc_h)
    
    c.setFont("Helvetica", 8)
    code_label = f"CODE: {code}"
    code_label_w = c.stringWidth(code_label, "Helvetica", 8)
    label_x = x + (img_w - code_label_w) / 2
    label_y = y - 12
    c.drawString(label_x, label_y, code_label)
    
    c.save()
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
#  PAGE ROUTES
# ─────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/scanner')
def scanner_page():
    return redirect(url_for('index'))


# ─────────────────────────────────────────────
#  AUTH ENDPOINTS
# ─────────────────────────────────────────────

@app.route('/api/login', methods=['POST'])
@limiter.limit("5 per minute")
def login():
    try:
        data = request.json
        username = data.get('username', '').strip().lower()
        password = data.get('password', '').strip()
        
        if not username or not password:
            return jsonify({'status': 'error', 'message': 'Username and passkey are required.'}), 400
            
        conn = db_manager.get_db_connection()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        
        if not user:
            conn.close()
            return jsonify({'status': 'error', 'message': 'Invalid username or passkey.'}), 401
            
        pw_bytes = password.encode()
        stored_hash = user['password_hash']
        
        if stored_hash.startswith('$2b$') or stored_hash.startswith('$2a$'):
            if not bcrypt.checkpw(pw_bytes, stored_hash.encode()):
                conn.close()
                return jsonify({'status': 'error', 'message': 'Invalid username or passkey.'}), 401
        elif stored_hash == hashlib.sha256(pw_bytes).hexdigest():
            new_hash = bcrypt.hashpw(pw_bytes, bcrypt.gensalt()).decode()
            conn.execute("UPDATE users SET password_hash = ? WHERE username = ?", (new_hash, username))
            conn.commit()
        else:
            conn.close()
            return jsonify({'status': 'error', 'message': 'Invalid username or passkey.'}), 401
            
        if user['status'] != 'Approved':
            conn.close()
            return jsonify({'status': 'error', 'message': 'Your account is pending admin approval.'}), 403
            
        supervised = (user['supervised_schools'] or '').strip()
        emp_name = (user['employee_name'] or '').strip()
        session['username'] = user['username']
        session['user_role'] = user['role']
        session['school_office'] = user['school_office'] or ''
        session['supervised_schools'] = supervised
        session['email'] = user['email'] or ''
        session['employee_name'] = emp_name
        session.permanent = True
        conn.close()
        
        return jsonify({
            'status': 'success',
            'message': f'Logged in successfully as {user["role"]}!',
            'requires_password_change': bool(user['requires_password_change']),
            'user': {
                'username': user['username'],
                'email': user['email'] or '',
                'role': user['role'],
                'school_office': user['school_office'] or '',
                'supervised_schools': supervised,
                'employee_name': emp_name
            }
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/register', methods=['POST'])
@limiter.limit("3 per minute")
def register():
    try:
        data = request.json
        username = data.get('username', '').strip().lower()
        role_request = data.get('role', 'User').strip()
        school_office = data.get('school_office', '').strip()
        email = data.get('email', '').strip()
        employee_name = data.get('employee_name', '').strip()
        
        # Validate username: no spaces allowed
        if ' ' in username:
            return jsonify({'status': 'error', 'message': 'Username must not contain spaces.'}), 400
            
        if not username:
            return jsonify({'status': 'error', 'message': 'Username is required.'}), 400
        
        if not school_office:
            return jsonify({'status': 'error', 'message': 'School/Office is required.'}), 400
            
        # Validate email domain
        if not email.endswith('@deped.gov.ph'):
            return jsonify({'status': 'error', 'message': 'Only @deped.gov.ph email addresses are allowed.'}), 400
            
        if role_request not in ['User', 'Supervisor', 'Admin']:
            role_request = 'User'
            
        conn = db_manager.get_db_connection()
        existing = conn.execute("SELECT username FROM users WHERE username = ?", (username,)).fetchone()
        
        if existing:
            conn.close()
            return jsonify({'status': 'error', 'message': 'Username already exists.'}), 400
        
        import secrets
        password = secrets.token_urlsafe(8)
        pw_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        conn.execute(
            "INSERT INTO users (username, password_hash, role, status, school_office, email, requires_password_change, employee_name) "
            "VALUES (?, ?, ?, 'Pending', ?, ?, 1, ?)",
            (username, pw_hash, role_request, school_office, email, employee_name)
        )
        conn.commit()
        conn.close()
        
        # Enqueue for async GSheets push
        db_manager.enqueue_sync('users', code=username)
            
        return jsonify({'status': 'success', 'message': 'Registration request submitted! Please wait for Admin approval.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ── Guest login (no credentials needed, read-only records view) ──
@app.route('/api/guest_login', methods=['POST'])
def guest_login():
    try:
        data = request.json or {}
        school = data.get('school_office', '').strip()
        session['username'] = 'guest'
        session['user_role'] = 'Guest'
        session['school_office'] = school
        session.permanent = True
        return jsonify({
            'status': 'success',
            'message': 'Logged in as Guest!' + (f' Viewing: {school}' if school else ''),
            'user': {
                'username': 'guest',
                'role': 'Guest',
                'school_office': school
            }
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/check_auth', methods=['GET'])
def check_auth():
    return jsonify({
        'logged_in': bool(session.get('username')),
        'username': session.get('username', ''),
        'email': session.get('email', ''),
        'role': session.get('user_role', 'Viewer'),
        'school_office': session.get('school_office', ''),
        'supervised_schools': session.get('supervised_schools', ''),
        'employee_name': session.get('employee_name', '')
    })

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'status': 'success'})

# ── Update user profile (school_office) ──
@app.route('/api/update_profile', methods=['POST'])
@require_login
def update_profile():
    try:
        data = request.json
        school_office = data.get('school_office', '').strip()
        username = session['username']
        
        conn = db_manager.get_db_connection()
        conn.execute("UPDATE users SET school_office = ? WHERE username = ?", (school_office, username))
        conn.commit()
        conn.close()
        
        session['school_office'] = school_office
        
        return jsonify({'status': 'success', 'message': 'Profile updated!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/get_user_profile', methods=['GET'])
@require_login
def get_user_profile():
    try:
        conn = db_manager.get_db_connection()
        user = conn.execute(
            "SELECT username, email, role, status, school_office, supervised_schools, employee_name FROM users WHERE username = ?",
            (session['username'],)
        ).fetchone()
        conn.close()
        if user:
            return jsonify({'status': 'success', 'user': dict(user)})
        return jsonify({'status': 'error', 'message': 'User not found'}), 404
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ─────────────────────────────────────────────
#  ADMIN USER MANAGEMENT APIs
# ─────────────────────────────────────────────
@app.route('/api/admin/users', methods=['GET'])
@require_admin
def admin_get_users():
    try:
        conn = db_manager.get_db_connection()
        rows = conn.execute("SELECT username, role, status, school_office, supervised_schools, employee_name FROM users WHERE username != 'admin'").fetchall()
        conn.close()
        users = [dict(r) for r in rows]
        return jsonify({'status': 'success', 'users': users})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/admin/approve_user/<username>', methods=['POST'])
@require_admin
def admin_approve_user(username):
    try:
        username = username.strip().lower()
        data = request.json or {}
        assigned_role = data.get('role', 'User')
        
        conn = db_manager.get_db_connection()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        if not user:
            conn.close()
            return jsonify({'status': 'error', 'message': 'User not found.'}), 404
            
        conn.execute(
            "UPDATE users SET status = 'Approved', role = ? WHERE username = ?",
            (assigned_role, username)
        )
        conn.commit()
        conn.close()
        
        # Enqueue for async GSheets push
        db_manager.enqueue_sync('users', code=username)
            
        return jsonify({'status': 'success', 'message': f'Account "{username}" approved as {assigned_role}!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/admin/reject_user/<username>', methods=['POST'])
@require_admin
def admin_reject_user(username):
    try:
        username = username.strip().lower()
        
        # Protect the admin account from being rejected/deleted
        if username == 'admin':
            return jsonify({'status': 'error', 'message': 'The Administrator account cannot be deleted.'}), 400
            
        conn = db_manager.get_db_connection()
        conn.execute("DELETE FROM users WHERE username = ?", (username,))
        conn.commit()
        conn.close()
        
        # Enqueue for async GSheets push
        db_manager.enqueue_sync('users', code=username)
            
        return jsonify({'status': 'success', 'message': f'Account request "{username}" rejected.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ── Request profile change (pending admin approval) ──
@app.route('/api/request_profile_change', methods=['POST'])
@require_login
def request_profile_change():
    try:
        data = request.json
        current_password = data.get('current_password', '')
        changes = data.get('changes', {})
        username = session['username']

        if not current_password:
            return jsonify({'status': 'error', 'message': 'Current password is required.'}), 400

        # Verify current password
        conn = db_manager.get_db_connection()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        if not user:
            conn.close()
            return jsonify({'status': 'error', 'message': 'User not found.'}), 404

        stored_hash = user['password_hash']
        pw_bytes = current_password.encode()
        if stored_hash.startswith('$2b$') or stored_hash.startswith('$2a$'):
            if not bcrypt.checkpw(pw_bytes, stored_hash.encode()):
                conn.close()
                return jsonify({'status': 'error', 'message': 'Current password is incorrect.'}), 401
        elif stored_hash == hashlib.sha256(pw_bytes).hexdigest():
            pass
        else:
            conn.close()
            return jsonify({'status': 'error', 'message': 'Current password is incorrect.'}), 401

        applied_immediately = []
        pending_fields = {}

        # ── Password: applied immediately ──
        if 'password' in changes:
            new_pw = changes['password'].strip()
            if new_pw:
                if len(new_pw) < 8:
                    conn.close()
                    return jsonify({'status': 'error', 'message': 'New password must be at least 8 characters.'}), 400
                new_hash = bcrypt.hashpw(new_pw.encode(), bcrypt.gensalt()).decode()
                conn.execute("UPDATE users SET password_hash = ?, requires_password_change = 0 WHERE username = ?", (new_hash, username))
                applied_immediately.append('password')

        # ── Username: validate uniqueness ──
        if 'username' in changes:
            new_username = changes['username'].strip().lower()
            if new_username and new_username != username:
                existing = conn.execute("SELECT username FROM users WHERE username = ?", (new_username,)).fetchone()
                if existing:
                    conn.close()
                    return jsonify({'status': 'error', 'message': 'Username already taken.'}), 400
                pending_fields['username'] = new_username
                pending_fields['_current_username'] = username

        # ── Other fields → pending ──
        for field in ['employee_name', 'email']:
            if field in changes:
                val = changes[field]
                if val is not None and (not isinstance(val, str) or val.strip()):
                    pending_fields[field] = val.strip() if isinstance(val, str) else val

        pending_id = None
        if pending_fields:
            # Add current username so apply knows which row to update (even if username changes)
            if '_current_username' not in pending_fields:
                pending_fields['_current_username'] = username
            pending_id = db_manager.create_pending_change(username, pending_fields)

        conn.close()

        # Enqueue for async GSheets push (password change)
        if 'password' in applied_immediately:
            db_manager.enqueue_sync('users', code=username)

        msg_parts = []
        if applied_immediately:
            msg_parts.append(f"{', '.join(applied_immediately)} updated successfully")
        if pending_id:
            msg_parts.append("other changes submitted for admin approval")
        msg = '. '.join(msg_parts) + '.' if msg_parts else 'No changes to apply.'

        return jsonify({
            'status': 'success',
            'message': msg,
            'applied_immediately': applied_immediately,
            'pending_id': pending_id
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/admin/pending_changes', methods=['GET'])
@require_admin
def admin_get_pending_changes():
    try:
        changes = db_manager.get_pending_changes()
        return jsonify({'status': 'success', 'changes': changes})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/admin/approve_change/<int:change_id>', methods=['POST'])
@require_admin
def admin_approve_change(change_id):
    try:
        result = db_manager.apply_pending_change(change_id)
        if result['status'] == 'error':
            return jsonify({'status': 'error', 'message': result['message']}), 400
        return jsonify({'status': 'success', 'message': 'Change approved and applied.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/admin/reject_change/<int:change_id>', methods=['POST'])
@require_admin
def admin_reject_change(change_id):
    try:
        db_manager.reject_pending_change(change_id)
        return jsonify({'status': 'success', 'message': 'Change rejected.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ── Change password (for first login or regular) ──
@app.route('/api/change_password', methods=['POST'])
@require_login
def change_password():
    try:
        data = request.json
        current_password = data.get('current_password', '').strip()
        new_password = data.get('new_password', '').strip()
        username = session['username']
        
        if not current_password or not new_password:
            return jsonify({'status': 'error', 'message': 'Current and new password are required.'}), 400
            
        if len(new_password) < 8:
            return jsonify({'status': 'error', 'message': 'New passkey must be at least 8 characters.'}), 400

        import re
        pw = new_password
        rules = []
        if not re.search(r'[A-Z]', pw): rules.append('uppercase letter')
        if not re.search(r'[a-z]', pw): rules.append('lowercase letter')
        if not re.search(r'[0-9]', pw): rules.append('number')
        if not re.search(r'[!@#$%^&*()_+\-=\[\]{}|;:\'",.<>?/]', pw): rules.append('special character')
        if rules:
            return jsonify({'status': 'error', 'message': f'Passkey must contain at least one {", ".join(rules)}.'}), 400
            
        conn = db_manager.get_db_connection()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        
        if not user:
            conn.close()
            return jsonify({'status': 'error', 'message': 'User not found.'}), 404
        
        pw_bytes = current_password.encode()
        stored_hash = user['password_hash']
        
        if stored_hash.startswith('$2b$') or stored_hash.startswith('$2a$'):
            if not bcrypt.checkpw(pw_bytes, stored_hash.encode()):
                conn.close()
                return jsonify({'status': 'error', 'message': 'Current passkey is incorrect.'}), 400
        elif stored_hash != hashlib.sha256(pw_bytes).hexdigest():
            conn.close()
            return jsonify({'status': 'error', 'message': 'Current passkey is incorrect.'}), 400
            
        new_hash = bcrypt.hashpw(new_password.encode(), bcrypt.gensalt()).decode()
        conn.execute("UPDATE users SET password_hash = ?, requires_password_change = 0 WHERE username = ?", (new_hash, username))
        conn.commit()
        conn.close()
        
        return jsonify({'status': 'success', 'message': 'Passkey changed successfully!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/change_email', methods=['POST'])
@require_login
def change_email():
    try:
        data = request.json
        current_password = data.get('current_password', '').strip()
        new_email = data.get('new_email', '').strip().lower()
        username = session['username']

        if not current_password or not new_email:
            return jsonify({'status': 'error', 'message': 'Current password and new email are required.'}), 400

        if not new_email.endswith('@deped.gov.ph'):
            return jsonify({'status': 'error', 'message': 'Only @deped.gov.ph email addresses are allowed.'}), 400

        conn = db_manager.get_db_connection()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()

        if not user:
            conn.close()
            return jsonify({'status': 'error', 'message': 'User not found.'}), 404

        cur_hash = hashlib.sha256(current_password.encode()).hexdigest()
        if user['password_hash'] != cur_hash:
            conn.close()
            return jsonify({'status': 'error', 'message': 'Current password is incorrect.'}), 400

        conn.execute("UPDATE users SET email = ? WHERE username = ?", (new_email, username))
        conn.commit()
        conn.close()

        session['email'] = new_email

        return jsonify({'status': 'success', 'message': 'Email changed successfully!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/admin/reset_passkey', methods=['POST'])
@require_admin
def admin_reset_passkey():
    try:
        data = request.json
        target_username = data.get('username', '').strip().lower()
        new_password = data.get('new_password', '').strip()
        
        if not target_username or not new_password:
            return jsonify({'status': 'error', 'message': 'Username and new password are required.'}), 400
            
        if len(new_password) < 8:
            return jsonify({'status': 'error', 'message': 'New passkey must be at least 8 characters.'}), 400

        import re
        pw = new_password
        rules = []
        if not re.search(r'[A-Z]', pw): rules.append('uppercase letter')
        if not re.search(r'[a-z]', pw): rules.append('lowercase letter')
        if not re.search(r'[0-9]', pw): rules.append('number')
        if not re.search(r'[!@#$%^&*()_+\-=\[\]{}|;:\'",.<>?/]', pw): rules.append('special character')
        if rules:
            return jsonify({'status': 'error', 'message': f'Passkey must contain at least one {", ".join(rules)}.'}), 400
            
        conn = db_manager.get_db_connection()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (target_username,)).fetchone()
        if not user:
            conn.close()
            return jsonify({'status': 'error', 'message': 'User not found.'}), 404
            
        new_hash = bcrypt.hashpw(new_password.encode(), bcrypt.gensalt()).decode()
        conn.execute("UPDATE users SET password_hash = ?, requires_password_change = 1 WHERE username = ?", (new_hash, target_username))
        conn.commit()
        
        conn.close()
        
        # Enqueue for async GSheets push
        db_manager.enqueue_sync('users', code=target_username)
            
        return jsonify({'status': 'success', 'message': f'Passkey for "{target_username}" reset successfully! User will need to change on next login.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ─────────────────────────────────────────────
#  MASTER DATA APIs (from SQLite master_data table)
# ─────────────────────────────────────────────
@app.route('/api/get_departments')
def get_departments():
    try:
        depts = db_manager.get_all_departments()
        return jsonify(depts)
    except Exception as e:
        return jsonify([])


@app.route('/api/get_schools', methods=['POST'])
def get_schools():
    try:
        dept = request.json.get('department', '')
        schools = db_manager.get_schools_for_department(dept)
        return jsonify(schools)
    except Exception as e:
        return jsonify([])


@app.route('/api/get_employees', methods=['POST'])
def get_employees():
    try:
        data = request.json
        employees = db_manager.get_employees_for_school(data.get('department', ''), data.get('school', ''))
        return jsonify(employees)
    except Exception as e:
        return jsonify([])


@app.route('/api/get_all_schools', methods=['GET'])
def get_all_schools():
    """All unique schools/offices for the scanner receiving-office picker."""
    try:
        schools = db_manager.get_all_schools()
        return jsonify({'status': 'success', 'schools': schools})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/get_employees_suggest', methods=['POST'])
def get_employees_suggest():
    """Return employees split into same-school first, then cross-school suggestions.
    Used by the unified employee picker in the Generate tab."""
    try:
        data = request.json
        department = data.get('department', '')
        school = data.get('school', '')
        result = db_manager.get_employees_with_suggestions(department, school)
        return jsonify({'status': 'success', **result})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/get_all_school_employees', methods=['POST'])
def get_all_school_employees():
    """Return ALL employees belonging to a specific school (for Add-All feature)."""
    try:
        data   = request.json
        school = data.get('school', '').strip()
        if not school:
            return jsonify({'status': 'error', 'message': 'School/Office is required.'})
        employees = db_manager.get_employees_by_school(school)
        return jsonify({'status': 'success', 'employees': employees, 'count': len(employees)})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ─────────────────────────────────────────────
#  ADMIN: MASTER DATA MANAGEMENT (from SQLite)
# ─────────────────────────────────────────────
@app.route('/api/admin/master_data', methods=['GET'])
@require_admin
def admin_get_master_data():
    """Return all master_data entries for admin editing."""
    try:
        data = db_manager.get_all_master_data()
        return jsonify({'status': 'success', 'data': data})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/admin/master_data/add', methods=['POST'])
@require_admin
def admin_add_master_data():
    """Add employee(s) to master_data. Body: { "department": "...", "school": "...", "employees": ["Name1", "Name2"] }"""
    try:
        data = request.json
        dept = data.get('department', '').strip()
        school = data.get('school', '').strip()
        employees = data.get('employees', [])
        if not dept or not school or not employees:
            return jsonify({'status': 'error', 'message': 'Department, school, and employees list are required.'}), 400
        count = db_manager.add_master_entries(dept, school, employees)
        return jsonify({'status': 'success', 'message': f'Added {count} employee(s) to master data.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/admin/master_data/delete/<int:entry_id>', methods=['DELETE'])
@require_admin
def admin_delete_master_data(entry_id):
    """Delete a single master_data entry."""
    try:
        db_manager.delete_master_entry(entry_id)
        return jsonify({'status': 'success', 'message': 'Entry deleted.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ─────────────────────────────────────────────
#  ADMIN: ALL GENERATED CODES
# ─────────────────────────────────────────────
@app.route('/api/admin/all_codes', methods=['GET'])
@require_admin
def admin_all_codes():
    """Return all entries from code_lookup, paginated."""
    try:
        page = int(request.args.get('page', 1))
        per_page = int(request.args.get('per_page', 50))
        search = request.args.get('search', '').strip()
        sort_by = request.args.get('sort_by', 'generated_at')
        order = request.args.get('order', 'DESC').upper()

        allowed_sort = {'code', 'department', 'school_office', 'doc_type', 'generated_at'}
        if sort_by not in allowed_sort:
            sort_by = 'generated_at'
        if order not in ('ASC', 'DESC'):
            order = 'DESC'

        conn = db_manager.get_db_connection()
        where_sql = ""
        params = []
        if search:
            where_sql = "WHERE (code LIKE ? OR department LIKE ? OR school_office LIKE ? OR doc_type LIKE ? OR employees LIKE ?)"
            pat = f"%{search}%"
            params = [pat, pat, pat, pat, pat]

        total = conn.execute(f"SELECT COUNT(*) FROM code_lookup {where_sql}", params).fetchone()[0]
        offset = (page - 1) * per_page
        rows = conn.execute(
            f"SELECT * FROM code_lookup {where_sql} ORDER BY {sort_by} {order} LIMIT ? OFFSET ?",
            params + [per_page, offset]
        ).fetchall()

        # Batch-fetch employee names from RDBMS for all codes on this page
        codes = [dict(r) for r in rows]
        code_list = [c['code'] for c in codes]
        emp_map = {}
        if code_list:
            placeholders = ','.join(['?'] * len(code_list))
            emp_rows = conn.execute(f"""
                SELECT ce.code, e.name FROM code_employees ce
                JOIN employees e ON e.id = ce.employee_id
                WHERE ce.code IN ({placeholders})
            """, code_list).fetchall()
            for r in emp_rows:
                emp_map.setdefault(r['code'], []).append(r['name'])
        for c in codes:
            if c['code'] in emp_map:
                c['employees_list'] = emp_map[c['code']]
            else:
                emp_str = c.get('employees', '')
                c['employees_list'] = emp_str.split('|||') if '|||' in emp_str else emp_str.split(',')

        conn.close()

        total_pages = max(1, (total + per_page - 1) // per_page)
        return jsonify({
            'status': 'success',
            'codes': codes,
            'pagination': {
                'page': page,
                'per_page': per_page,
                'total': total,
                'total_pages': total_pages
            }
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ─────────────────────────────────────────────
#  CODE GENERATION
# ─────────────────────────────────────────────
@app.route('/generate_code', methods=['POST'])
@require_login
def generate_code():
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    try:
        data       = request.json
        code       = generate_6_digit_code()
        department = data.get('department')
        school     = data.get('school')
        employees  = data.get('employees')
        doc_type   = data.get('doc_type')

        # ── ROLE-BASED SCHOOL RESTRICTIONS ──
        user_role     = session.get('user_role', '')
        user_school   = session.get('school_office', '')

        if user_role == 'User' and user_school:
            # Force to assigned school
            conn = db_manager.get_db_connection()
            row = conn.execute(
                "SELECT department FROM schools WHERE name = ? LIMIT 1",
                (user_school,)
            ).fetchone()
            conn.close()
            if row:
                department = row['department']
            school = user_school

        # ── CUSTOM NAMES ARE MONITORING-ONLY: ensure they are NOT inserted into master_data ──
        # save_code_lookup() stores names in code_lookup.employees; it does NOT touch master_data.
        # The only place that adds to master_data is the admin API (/api/admin/master_data/add)
        # or the initial Excel seed. So no extra guard needed here — names are monitoring-only by design.

        save_code_lookup(code, department, school, employees, doc_type)

        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        basename  = f"code_{timestamp}"

        # Generate QR (1.5" x 1.5") and Barcode (1.5" x 0.75")
        qr_img = generate_qr_image(code)
        qr_path = os.path.join(QR_FOLDER, f"{basename}_qr.png")
        qr_img.save(qr_path)

        bc_img = generate_barcode_image(code)
        bc_path = os.path.join(QR_FOLDER, f"{basename}_barcode.png")
        bc_img.save(bc_path)

        # Also save a combined preview image (QR on top, barcode below)
        combined_w = max(qr_img.width, bc_img.width)
        combined_h = qr_img.height + bc_img.height
        combined = Image.new('RGB', (combined_w, combined_h), 'white')
        combined.paste(qr_img, ((combined_w - qr_img.width) // 2, 0))
        combined.paste(bc_img, ((combined_w - bc_img.width) // 2, qr_img.height))
        combined_path = os.path.join(QR_FOLDER, f"{basename}.png")
        combined.save(combined_path)

        return jsonify({
            'status':         'success',
            'code_display':   code,
            'image_url':      f"/static/qr_generated/{basename}.png",
            'qr_image_url':   f"/static/qr_generated/{basename}_qr.png",
            'barcode_image_url': f"/static/qr_generated/{basename}_barcode.png",
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ─────────────────────────────────────────────
#  BATCH CODE GENERATION
# ─────────────────────────────────────────────
@app.route('/api/generate_batch', methods=['POST'])
@require_login
def generate_batch():
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    try:
        data = request.json
        generations = data.get('generations', [])
        if not generations:
            return jsonify({'status': 'error', 'message': 'No generations provided.'}), 400

        results = []
        code_dicts = []
        user_role = session.get('user_role', '')
        user_school = session.get('school_office', '')
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')

        for idx, gen in enumerate(generations):
            code = generate_6_digit_code()
            department = gen.get('department')
            school = gen.get('school')
            employees = gen.get('employees', [])
            doc_type = gen.get('doc_type')

            if user_role == 'User' and user_school:
                conn = db_manager.get_db_connection()
                row = conn.execute(
                    "SELECT department FROM master_data WHERE school_office = ? LIMIT 1",
                    (user_school,)
                ).fetchone()
                conn.close()
                if row:
                    department = row['department']
                school = user_school

            employees = [db_manager.reorder_name(e) for e in employees]
            employees_str = '|||'.join(employees)
            code_dicts.append({
                'code': code,
                'department': department or '',
                'school_office': school or '',
                'employees': employees_str,
                'doc_type': doc_type or '',
                'generated_at': timestamp
            })

            basename = f"code_{timestamp}_{idx}"
            results.append({
                'code': code,
                'department': department,
                'school': school,
                'employees': employees,
                'doc_type': doc_type,
                'qr_image_url': None,
                'barcode_image_url': None
            })

        save_code_lookup_batch(code_dicts)

        ts = datetime.now().strftime('%Y%m%d%H%M%S')
        for idx, r in enumerate(results):
            basename = f"code_{ts}_{idx}"
            qr_img = generate_qr_image(r['code'])
            qr_path = os.path.join(QR_FOLDER, f"{basename}_qr.png")
            qr_img.save(qr_path)

            bc_img = generate_barcode_image(r['code'])
            bc_path = os.path.join(QR_FOLDER, f"{basename}_barcode.png")
            bc_img.save(bc_path)

            combined_w = max(qr_img.width, bc_img.width)
            combined_h = qr_img.height + bc_img.height
            combined = Image.new('RGB', (combined_w, combined_h), 'white')
            combined.paste(qr_img, ((combined_w - qr_img.width) // 2, 0))
            combined.paste(bc_img, ((combined_w - bc_img.width) // 2, qr_img.height))
            combined_path = os.path.join(QR_FOLDER, f"{basename}.png")
            combined.save(combined_path)

            results[idx]['qr_image_url'] = f"/static/qr_generated/{basename}_qr.png"
            results[idx]['barcode_image_url'] = f"/static/qr_generated/{basename}_barcode.png"
            results[idx]['image_url'] = f"/static/qr_generated/{basename}.png"

        return jsonify({'status': 'success', 'results': results, 'count': len(results)})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ─────────────────────────────────────────────
#  BARCODE LABEL / DOCUMENT PDF ENDPOINTS
# ─────────────────────────────────────────────

@app.route('/api/label_image/<code>')
def label_image(code):
    """Serve a combined PNG image (QR top, barcode bottom) for the given code."""
    try:
        qr_img = generate_qr_image(code.upper())
        bc_img = generate_barcode_image(code.upper())
        combined_w = max(qr_img.width, bc_img.width)
        combined_h = qr_img.height + bc_img.height
        combined = Image.new('RGB', (combined_w, combined_h), 'white')
        combined.paste(qr_img, ((combined_w - qr_img.width) // 2, 0))
        combined.paste(bc_img, ((combined_w - bc_img.width) // 2, qr_img.height))
        buf = io.BytesIO()
        combined.save(buf, format='PNG')
        buf.seek(0)
        return send_file(buf, mimetype='image/png')
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/document_page', methods=['POST'])
def document_page():
    """
    Generate a PDF with QR + barcode at TOP RIGHT corner.
    Body: { "code": "..." }
    """
    try:
        data = request.json
        code = data.get('code', '').strip().upper()
        
        if not code:
            return jsonify({'status': 'error', 'message': 'Code is required.'}), 400
        
        pdf_buf = create_document_page_pdf(code)
        return send_file(
            pdf_buf,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"barcode_{code}.pdf"
        )
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/batch_labels', methods=['POST'])
def batch_labels():
    """
    Generate a PDF for batch printing (QR + barcode per code).
    Body: { "codes": [{"code": "..."}, ...] }
    """
    try:
        data = request.json
        codes = data.get('codes', [])
        
        if not codes:
            return jsonify({'status': 'error', 'message': 'No codes provided.'}), 400
        
        pdf_buf = create_label_pdf(codes)
        return send_file(
            pdf_buf,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='barcode_labels.pdf'
        )
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/barcode_overlay', methods=['POST'])
def barcode_overlay():
    """
    Generate a minimal overlay PDF with QR + barcode at the specified position.
    Body: { "code": "...", "position": "top-right|top-left|bottom-right|bottom-left", "page_size": "legal|letter" }
    """
    try:
        data = request.json
        code = data.get('code', '').strip().upper()
        position = data.get('position', 'top-right')
        page_size = data.get('page_size', 'legal')
        
        if not code:
            return jsonify({'status': 'error', 'message': 'Code is required.'}), 400
        
        pdf_buf = create_barcode_overlay_pdf(code, position, page_size)
        return send_file(
            pdf_buf,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"barcode_overlay_{code}.pdf"
        )
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ─────────────────────────────────────────────
#  SCAN LOGGING
# ─────────────────────────────────────────────
@app.route('/log_scan', methods=['POST'])
@require_login
def log_scan():
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403

    data             = request.json
    code             = data.get('scanned_data', '').strip().upper()
    receiving_office = data.get('receiving_office', '').strip()
    receiver_name    = data.get('receiver_name', '').strip()

    # ── ROLE-BASED SCHOOL RESTRICTIONS ──
    user_role     = session.get('user_role', '')
    user_school   = session.get('school_office', '')
    if user_role == 'User' and user_school:
        receiving_office = user_school

    if not receiving_office:
        return jsonify({'status': 'error', 'message': 'Please select a Receiving Office!'})

    lookup = get_code_lookup(code)
    if not lookup:
        return jsonify({'status': 'error', 'message': f'Barcode "{code}" not found in database!'})

    employees = [e.strip() for e in lookup['employees'] if e.strip()]
    if not employees:
        return jsonify({'status': 'error', 'message': 'No employees linked to this barcode!'})

    current_ts       = datetime.now().strftime("%m/%d/%Y")
    updated_rows_ids = []
    slot_updated     = 1

    with scan_lock:
        try:
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
                    latest_receiver = None
                    for i in sorted(stages, reverse=True):
                        key = f'receiving_office_{i}'
                        if key in row.keys() and row[key]:
                            latest_office = row[key]
                            receiver_key = f'receiver_name_{i}'
                            latest_receiver = row[receiver_key] if receiver_key in row.keys() else ''
                            latest_ts     = row[f'timestamp_{i}']
                            break

                    if (latest_office and latest_office.strip().lower() == receiving_office.lower() and 
                        latest_receiver and latest_receiver.strip().lower() == receiver_name.lower()):
                        conn.close()
                        return jsonify({
                            'status':  'warning',
                            'message': f'Duplicate scan! Already received at "{latest_office}" by "{latest_receiver}" on {latest_ts}.'
                        })

            # ── 2. LOG THE ROUTING STAMP ──────────────────────────────────────────
            # Resolve employee IDs once
            emp_id_map = {}
            emp_ids_list = _resolve_employee_ids(conn, employees, lookup['school'])
            for name, eid in emp_ids_list:
                emp_id_map[name] = eid

            for emp in employees:
                row = conn.execute(
                    "SELECT * FROM routing_records WHERE code = ? AND employee = ?", (code, emp)
                ).fetchone()

                if not row:
                    db_manager.ensure_routing_columns(conn, 1)
                    cursor = conn.cursor()
                    emp_id = emp_id_map.get(emp)
                    cursor.execute(
                        "INSERT INTO routing_records "
                        "(department, school_office, employee, employee_id, code, receiving_office_1, receiver_name_1, timestamp_1) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                        (lookup['department'], lookup['school'], emp, emp_id, code,
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

                # Auto-set status to 'released' when scanned at RECORDS SERVICES
                # unless admin manually set it to 'with corrections'
                if receiving_office.strip().upper() == 'RECORDS SERVICES':
                    prev_status = row['status'] if row else ''
                    if prev_status != 'with corrections':
                        conn.execute(
                            "UPDATE routing_records SET status = 'released' "
                            "WHERE code = ? AND employee = ?",
                            (code, emp)
                        )
                        conn.commit()

            conn.close()
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)})

    # ── Post-processing (no lock needed) ──────────────────────────────────
    if updated_rows_ids:
        db_manager.queue_sync_batch(updated_rows_ids)
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


# ─────────────────────────────────────────────
#  ROUTING RECORDS — PAGINATED
# ─────────────────────────────────────────────
@app.route('/api/get_routing_records', methods=['GET'])
@require_login
def get_routing_records():
    try:
        page     = int(request.args.get('page', 1))
        per_page = int(request.args.get('per_page', 50))
        search   = request.args.get('search', '').strip()
        sort_by  = request.args.get('sort_by', 'last_activity')
        order    = request.args.get('order', 'DESC').upper()

        # Build a dynamic "last_activity" expression: MAX of all timestamp columns
        # converted to YYYYMMDD so cross-year comparisons work correctly.
        def _sortable_ts(col):
            return (
                f"CASE WHEN {col} != '' AND {col} IS NOT NULL "
                f"THEN SUBSTR({col}, 7, 4) || SUBSTR({col}, 1, 2) || SUBSTR({col}, 4, 2) "
                f"ELSE '00000000' END"
            )
        ts_exprs = ', '.join(
            _sortable_ts(f"timestamp_{i}") for i in range(1, 11)
        )
        last_activity_expr = f"MAX({ts_exprs})"

        # Sanitise sort params — 'last_activity' is a virtual computed sort
        allowed_sort = {'id', 'department', 'school_office', 'employee', 'code', 'last_activity', 'status', 'doc_type'}
        if sort_by not in allowed_sort:
            sort_by = 'last_activity'
        if order not in ('ASC', 'DESC'):
            order = 'DESC'

        conn = db_manager.get_db_connection()

        where_sql = ""
        params    = []

        # ── ROLE-BASED FILTERING ──
        user_role   = session.get('user_role', '')
        user_school = session.get('school_office', '')

        if user_role == 'User' and user_school:
            where_sql = "WHERE school_office = ?"
            params.append(user_school)
            if search:
                where_sql += " AND (employee LIKE ? OR code LIKE ? OR department LIKE ? OR school_office LIKE ?)"
                pat = f"%{search}%"
                params += [pat, pat, pat, pat]

        elif user_role == 'Guest' and user_school:
            where_sql = "WHERE school_office = ?"
            params.append(user_school)
            if search:
                where_sql += " AND (employee LIKE ? OR code LIKE ? OR department LIKE ? OR school_office LIKE ?)"
                pat = f"%{search}%"
                params += [pat, pat, pat, pat]

        elif search:
            where_sql = ("WHERE (employee LIKE ? OR code LIKE ? "
                         "OR department LIKE ? OR school_office LIKE ?)")
            pat    = f"%{search}%"
            params = [pat, pat, pat, pat]

        total = conn.execute(
            f"SELECT COUNT(*) FROM routing_records {where_sql}", params
        ).fetchone()[0]

        offset = (page - 1) * per_page

        # Determine ORDER BY clause
        if sort_by == 'last_activity':
            order_clause = f"{last_activity_expr} {order}, id DESC"
        else:
            order_clause = f"{sort_by} {order}"

        rows = conn.execute(
            f"SELECT * FROM routing_records {where_sql} "
            f"ORDER BY {order_clause} LIMIT ? OFFSET ?",
            params + [per_page, offset]
        ).fetchall()

        # Batch-fetch doc_type for all codes in the result set
        codes_in_page = list({dict(r)['code'] for r in rows if dict(r).get('code')})
        code_doc_type_map = {}
        if codes_in_page:
            placeholders = ','.join(['?'] * len(codes_in_page))
            lookup_rows = conn.execute(
                f"SELECT code, doc_type FROM code_lookup WHERE code IN ({placeholders})", codes_in_page
            ).fetchall()
            code_doc_type_map = {row['code']: row['doc_type'] or '' for row in lookup_rows}

        conn.close()

        records = []
        for r in rows:
            d = dict(r)
            # Inject doc_type from code_lookup if not already present in the record
            if not d.get('doc_type') and d.get('code'):
                d['doc_type'] = code_doc_type_map.get(d['code'], '')
            # Compute last_activity (most recent non-empty timestamp) for display
            timestamps = [d.get(f'timestamp_{i}', '') or '' for i in range(1, 11)]
            non_empty = [t for t in timestamps if t.strip()]
            if non_empty:
                def _ts_key(t):
                    try:
                        return t[6:10] + t[0:2] + t[3:5]
                    except (IndexError, TypeError):
                        return t
                d['last_activity'] = max(non_empty, key=_ts_key)
            else:
                d['last_activity'] = ''
            records.append(d)

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
#  ROUTING RECORDS — GROUPED (server-side group pagination)
# ─────────────────────────────────────────────
@app.route('/api/get_routing_grouped', methods=['GET'])
@require_login
def get_routing_grouped():
    try:
        page     = int(request.args.get('page', 1))
        per_page = int(request.args.get('per_page', 75))
        search   = request.args.get('search', '').strip()
        sort_by  = request.args.get('sort_by', 'last_activity')
        order    = request.args.get('order', 'DESC').upper()

        allowed_sort = {'id', 'school_office', 'employee', 'code', 'doc_type', 'status', 'last_activity'}
        if sort_by not in allowed_sort:
            sort_by = 'last_activity'
        if order not in ('ASC', 'DESC'):
            order = 'DESC'

        conn = db_manager.get_db_connection()

        where_sql = ""
        params    = []

        user_role   = session.get('user_role', '')
        user_school = session.get('school_office', '')

        if user_role == 'User' and user_school:
            where_sql = "WHERE r.school_office = ?"
            params.append(user_school)
            if search:
                where_sql += " AND (r.employee LIKE ? OR r.code LIKE ? OR r.department LIKE ? OR r.school_office LIKE ?)"
                pat = f"%{search}%"
                params += [pat, pat, pat, pat]

        elif user_role == 'Guest' and user_school:
            where_sql = "WHERE r.school_office = ?"
            params.append(user_school)
            if search:
                where_sql += " AND (r.employee LIKE ? OR r.code LIKE ? OR r.department LIKE ? OR r.school_office LIKE ?)"
                pat = f"%{search}%"
                params += [pat, pat, pat, pat]

        elif search:
            where_sql = "WHERE (r.employee LIKE ? OR r.code LIKE ? OR r.department LIKE ? OR r.school_office LIKE ?)"
            pat    = f"%{search}%"
            params = [pat, pat, pat, pat]

        total = conn.execute(
            f"SELECT COUNT(DISTINCT r.code) FROM routing_records r {where_sql}", params
        ).fetchone()[0]

        offset = (page - 1) * per_page

        if sort_by == 'last_activity':
            order_clause = "MAX(r.id) DESC"
        elif sort_by == 'employee':
            order_clause = f"MIN(r.employee) {order}"
        elif sort_by == 'school_office':
            order_clause = f"r.school_office {order}"
        elif sort_by == 'code':
            order_clause = f"r.code {order}"
        elif sort_by == 'status':
            order_clause = f"r.status {order}"
        elif sort_by == 'doc_type':
            order_clause = f"cl.doc_type {order}"
        else:
            order_clause = "MAX(r.id) DESC"

        groups = conn.execute(f"""
            SELECT r.code, r.school_office, r.status,
                   COUNT(*) as employee_count,
                   MIN(r.employee) as first_employee,
                   MIN(r.id) as first_id,
                   MAX(r.id) as last_id
            FROM routing_records r
            LEFT JOIN code_lookup cl ON cl.code = r.code
            {where_sql}
            GROUP BY r.code
            ORDER BY {order_clause}
            LIMIT ? OFFSET ?
        """, params + [per_page, offset]).fetchall()

        codes_in_page = [g['code'] for g in groups]

        doc_type_map = {}
        if codes_in_page:
            placeholders = ','.join(['?'] * len(codes_in_page))
            lookup_rows = conn.execute(
                f"SELECT code, doc_type FROM code_lookup WHERE code IN ({placeholders})", codes_in_page
            ).fetchall()
            doc_type_map = {row['code']: row['doc_type'] or '' for row in lookup_rows}

        latest_map = {}
        if codes_in_page:
            placeholders = ','.join(['?'] * len(codes_in_page))
            latest_rows = conn.execute(f"""
                SELECT r1.* FROM routing_records r1
                INNER JOIN (
                    SELECT code, MAX(id) as max_id FROM routing_records
                    WHERE code IN ({placeholders})
                    GROUP BY code
                ) r2 ON r1.code = r2.code AND r1.id = r2.max_id
            """, codes_in_page).fetchall()
            for row in latest_rows:
                d = dict(row)
                code = d['code']
                latest_ts = ''
                latest_office = ''
                latest_receiver = ''
                for i in range(10, 0, -1):
                    ts = d.get(f'timestamp_{i}', '') or ''
                    if ts.strip():
                        latest_ts = ts
                        latest_office = d.get(f'receiving_office_{i}', '') or '-'
                        latest_receiver = d.get(f'receiver_name_{i}', '') or '-'
                        break
                latest_map[code] = {
                    'office': latest_office,
                    'receiver': latest_receiver,
                    'ts': latest_ts
                }

        conn.close()

        result = []
        for g in groups:
            d = dict(g)
            code = d['code']
            lm = latest_map.get(code, {})
            result.append({
                'code': code,
                'school_office': d['school_office'] or '',
                'doc_type': doc_type_map.get(code, '') or '',
                'status': d['status'] or 'for signature',
                'employee_count': d['employee_count'],
                'first_employee': d['first_employee'] or '',
                'extras': d['employee_count'] - 1,
                'summaryOffice': lm.get('office', ''),
                'summaryReceiver': lm.get('receiver', ''),
                'summaryTs': lm.get('ts', '')
            })

        total_pages = max(1, (total + per_page - 1) // per_page)

        return jsonify({
            'status': 'success',
            'groups': result,
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
#  CODE LOOKUP — for routing slip
# ─────────────────────────────────────────────
@app.route('/api/get_code_lookup/<path:code>', methods=['GET'])
@require_login
def get_code_lookup_api(code):
    try:
        lookup = get_code_lookup(code)
        if not lookup:
            return jsonify({'status': 'error', 'message': 'Code not found'}), 404
        return jsonify({'status': 'success', 'data': lookup})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ─────────────────────────────────────────────
#  ROUTING RECORDS — BY CODE (for expanding groups)
# ─────────────────────────────────────────────
@app.route('/api/get_records_by_code/<path:code>', methods=['GET'])
@require_login
def get_records_by_code(code):
    try:
        conn = db_manager.get_db_connection()

        rows = conn.execute(
            "SELECT * FROM routing_records WHERE code = ? ORDER BY id ASC", (code,)
        ).fetchall()

        dtype_row = conn.execute(
            "SELECT doc_type FROM code_lookup WHERE code = ?", (code,)
        ).fetchone()
        doc_type = dtype_row['doc_type'] if dtype_row else ''

        conn.close()

        records = []
        for r in rows:
            d = dict(r)
            if not d.get('doc_type'):
                d['doc_type'] = doc_type
            records.append(d)

        return jsonify({'status': 'success', 'records': records})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/get_known_routing_values', methods=['GET'])
@require_login
def get_known_routing_values():
    """Distinct receiving offices and receiver names from routing history (for autosuggest)."""
    try:
        conn = db_manager.get_db_connection()
        offices = set()
        receivers = set()
        for i in range(1, 11):
            for col, store in [('receiving_office', offices), ('receiver_name', receivers)]:
                rows = conn.execute(
                    f"SELECT DISTINCT {col}_{i} AS v FROM routing_records "
                    f"WHERE {col}_{i} IS NOT NULL AND trim({col}_{i}) != ''"
                ).fetchall()
                for r in rows:
                    store.add(r['v'])
        conn.close()
        return jsonify({
            'status': 'success',
            'offices': sorted(offices),
            'receivers': sorted(receivers)
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/schools_codes', methods=['GET'])
@require_login
def api_schools_codes():
    try:
        conn = db_manager.get_db_connection()
        rows = conn.execute("SELECT id, name, department, sta_code FROM schools ORDER BY name").fetchall()
        conn.close()
        return jsonify({'status': 'success', 'schools': [dict(r) for r in rows]})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/employees_codes', methods=['GET'])
@require_login
def api_employees_codes():
    try:
        conn = db_manager.get_db_connection()
        rows = conn.execute("""
            SELECT e.id, e.name, e.school_id, s.name AS school_name, e.employee_no, e.notes
            FROM employees e
            JOIN schools s ON s.id = e.school_id
            ORDER BY e.name
        """).fetchall()
        conn.close()
        return jsonify({'status': 'success', 'employees': [dict(r) for r in rows]})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ─────────────────────────────────────────────
#  REFERENCE DATA CRUD (Admin only)
# ─────────────────────────────────────────────
@app.route('/api/add_school', methods=['POST'])
@require_login
@require_role('Admin')
def api_add_school():
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    for _attempt in range(3):
        conn = None
        try:
            data = request.json
            name = (data.get('name') or '').strip()
            department = (data.get('department') or '').strip()
            sta_code = (data.get('sta_code') or '').strip() or None
            if not name:
                return jsonify({'status': 'error', 'message': 'School name is required.'})
            conn = db_manager.get_db_connection()
            existing = conn.execute("SELECT id FROM schools WHERE name = ?", (name,)).fetchone()
            if existing:
                return jsonify({'status': 'error', 'message': 'A school with this name already exists.'})
            conn.execute("INSERT INTO schools (name, department, sta_code) VALUES (?, ?, ?)",
                         (name, department, sta_code))
            conn.commit()
            new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.close()
            return jsonify({'status': 'success', 'message': f'School "{name}" added.', 'id': new_id})
        except sqlite3.OperationalError as _e:
            if 'locked' in str(_e) and _attempt < 2:
                time.sleep(1 * (_attempt + 1))
                continue
            return jsonify({'status': 'error', 'message': 'Database is busy, please try again.'})
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/edit_school', methods=['POST'])
@require_login
@require_role('Admin')
def api_edit_school():
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    for _attempt in range(3):
        conn = None
        try:
            data = request.json
            school_id = data.get('id')
            if not school_id:
                return jsonify({'status': 'error', 'message': 'School ID is required.'})
            name = (data.get('name') or '').strip()
            department = (data.get('department') or '').strip()
            sta_code = (data.get('sta_code') or '').strip() or None
            if not name:
                return jsonify({'status': 'error', 'message': 'School name is required.'})
            conn = db_manager.get_db_connection()
            duplicate = conn.execute("SELECT id FROM schools WHERE name = ? AND id != ?", (name, school_id)).fetchone()
            if duplicate:
                return jsonify({'status': 'error', 'message': 'Another school with this name already exists.'})
            conn.execute("UPDATE schools SET name = ?, department = ?, sta_code = ? WHERE id = ?",
                         (name, department, sta_code, school_id))
            conn.commit()
            conn.close()
            return jsonify({'status': 'success', 'message': 'School updated.'})
        except sqlite3.OperationalError as _e:
            if 'locked' in str(_e) and _attempt < 2:
                time.sleep(1 * (_attempt + 1))
                continue
            return jsonify({'status': 'error', 'message': 'Database is busy, please try again.'})
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/delete_school', methods=['POST'])
@require_login
@require_role('Admin')
def api_delete_school():
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    for _attempt in range(3):
        conn = None
        try:
            data = request.json
            school_id = data.get('id')
            if not school_id:
                return jsonify({'status': 'error', 'message': 'School ID is required.'})
            conn = db_manager.get_db_connection()
            emp_count = conn.execute("SELECT COUNT(*) FROM employees WHERE school_id = ?", (school_id,)).fetchone()[0]
            if emp_count > 0:
                return jsonify({'status': 'error', 'message': f'Cannot delete: {emp_count} employee(s) still assigned to this school. Reassign them first.'})
            conn.execute("DELETE FROM schools WHERE id = ?", (school_id,))
            conn.commit()
            conn.close()
            return jsonify({'status': 'success', 'message': 'School deleted.'})
        except sqlite3.OperationalError as _e:
            if 'locked' in str(_e) and _attempt < 2:
                time.sleep(1 * (_attempt + 1))
                continue
            return jsonify({'status': 'error', 'message': 'Database is busy, please try again.'})
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/add_employee', methods=['POST'])
@require_login
@require_role('Admin')
def api_add_employee():
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    for _attempt in range(3):
        conn = None
        try:
            data = request.json
            name = (data.get('name') or '').strip()
            school_id = data.get('school_id')
            employee_no = (data.get('employee_no') or '').strip() or None
            notes = (data.get('notes') or '').strip() or None
            if not name:
                return jsonify({'status': 'error', 'message': 'Employee name is required.'})
            if not school_id:
                return jsonify({'status': 'error', 'message': 'School is required.'})
            conn = db_manager.get_db_connection()
            school = conn.execute("SELECT id FROM schools WHERE id = ?", (school_id,)).fetchone()
            if not school:
                return jsonify({'status': 'error', 'message': 'Selected school does not exist.'})
            existing = conn.execute("SELECT id FROM employees WHERE name = ? AND school_id = ?", (name, school_id)).fetchone()
            if existing:
                return jsonify({'status': 'error', 'message': 'This employee already exists at this school.'})
            conn.execute("INSERT INTO employees (name, school_id, employee_no, notes) VALUES (?, ?, ?, ?)",
                         (name, school_id, employee_no, notes))
            conn.commit()
            new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.close()
            return jsonify({'status': 'success', 'message': f'Employee "{name}" added.', 'id': new_id})
        except sqlite3.OperationalError as _e:
            if 'locked' in str(_e) and _attempt < 2:
                time.sleep(1 * (_attempt + 1))
                continue
            return jsonify({'status': 'error', 'message': 'Database is busy, please try again.'})
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/edit_employee', methods=['POST'])
@require_login
@require_role('Admin')
def api_edit_employee():
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    for _attempt in range(3):
        conn = None
        try:
            data = request.json
            emp_id = data.get('id')
            if not emp_id:
                return jsonify({'status': 'error', 'message': 'Employee ID is required.'})
            name = (data.get('name') or '').strip()
            school_id = data.get('school_id')
            employee_no = (data.get('employee_no') or '').strip() or None
            notes = (data.get('notes') or '').strip() or None
            if not name:
                return jsonify({'status': 'error', 'message': 'Employee name is required.'})
            if not school_id:
                return jsonify({'status': 'error', 'message': 'School is required.'})
            conn = db_manager.get_db_connection()
            school = conn.execute("SELECT id FROM schools WHERE id = ?", (school_id,)).fetchone()
            if not school:
                return jsonify({'status': 'error', 'message': 'Selected school does not exist.'})
            duplicate = conn.execute("SELECT id FROM employees WHERE name = ? AND school_id = ? AND id != ?",
                                     (name, school_id, emp_id)).fetchone()
            if duplicate:
                return jsonify({'status': 'error', 'message': 'Another employee with this name already exists at this school.'})
            conn.execute("UPDATE employees SET name = ?, school_id = ?, employee_no = ?, notes = ? WHERE id = ?",
                         (name, school_id, employee_no, notes, emp_id))
            conn.commit()
            conn.close()
            return jsonify({'status': 'success', 'message': 'Employee updated.'})
        except sqlite3.OperationalError as _e:
            if 'locked' in str(_e) and _attempt < 2:
                time.sleep(1 * (_attempt + 1))
                continue
            return jsonify({'status': 'error', 'message': 'Database is busy, please try again.'})
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/delete_employee', methods=['POST'])
@require_login
@require_role('Admin')
def api_delete_employee():
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    for _attempt in range(3):
        conn = None
        try:
            data = request.json
            emp_id = data.get('id')
            confirmed = data.get('confirm', False)
            if not emp_id:
                return jsonify({'status': 'error', 'message': 'Employee ID is required.'})
            conn = db_manager.get_db_connection()
            emp = conn.execute("SELECT name FROM employees WHERE id = ?", (emp_id,)).fetchone()
            if not emp:
                return jsonify({'status': 'error', 'message': 'Employee not found.'})
            emp_name = emp['name']
            code_count = conn.execute("SELECT COUNT(*) FROM code_employees WHERE employee_id = ?", (emp_id,)).fetchone()[0]
            if code_count > 0 and not confirmed:
                return jsonify({'status': 'warning', 'message': f'This employee is linked to {code_count} document code(s). Delete anyway?', 'code_count': code_count, 'confirm_required': True})
            if confirmed:
                conn.execute("DELETE FROM code_employees WHERE employee_id = ?", (emp_id,))
            conn.execute("DELETE FROM employees WHERE id = ?", (emp_id,))
            conn.commit()
            conn.close()
            return jsonify({'status': 'success', 'message': f'Employee "{emp_name}" deleted.'})
        except sqlite3.OperationalError as _e:
            if 'locked' in str(_e) and _attempt < 2:
                time.sleep(1 * (_attempt + 1))
                continue
            return jsonify({'status': 'error', 'message': 'Database is busy, please try again.'})
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)}), 500


# ─────────────────────────────────────────────
#  EDIT / DELETE RECORDS (Admin / Supervisor)
# ─────────────────────────────────────────────
@app.route('/api/update_record/<int:record_id>', methods=['PUT'])
@require_login
@require_role('Admin', 'Supervisor')
def update_record(record_id):
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403

    for _attempt in range(3):
        conn = None
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
                return jsonify({'status': 'error', 'message': 'No valid fields to update.'})

            vals.append(record_id)
            conn.execute(f"UPDATE routing_records SET {', '.join(set_parts)} WHERE id = ?", vals)
            conn.commit()
        except sqlite3.OperationalError as _e:
            if 'locked' in str(_e) and _attempt < 2:
                time.sleep(1 * (_attempt + 1))
                continue
            return jsonify({'status': 'error', 'message': 'Database is busy, please try again.'})
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)})
        finally:
            if conn:
                try:
                    conn.close()
                except Exception:
                    pass

        # Post-processing (outside retry, runs once on success)
        db_manager.enqueue_sync('routing_records', ref_id=record_id)
        db_manager.trigger_excel_export()
        return jsonify({'status': 'success', 'message': 'Record updated.'})


@app.route('/api/update_records_by_code/<path:code>', methods=['PUT'])
@require_login
@require_role('Admin', 'Supervisor')
def update_records_by_code(code):
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    try:
        data = request.json
        conn = db_manager.get_db_connection()

        if 'doc_type' in data and data['doc_type']:
            conn.execute("UPDATE code_lookup SET doc_type = ? WHERE code = ?", (data['doc_type'], code))

        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(routing_records)")
        db_cols = {row[1] for row in cursor.fetchall()}

        set_parts, vals = [], []
        for field, value in data.items():
            if field == 'doc_type':
                continue
            if field in db_cols and field != 'id':
                set_parts.append(f"{field} = ?")
                vals.append(value)

        if set_parts:
            vals.append(code)
            conn.execute(f"UPDATE routing_records SET {', '.join(set_parts)} WHERE code = ?", vals)

        # Auto-set status to 'released' if any receiving_office is RECORDS SERVICES
        for i in range(1, 11):
            office_key = f'receiving_office_{i}'
            if data.get(office_key, '').strip().upper() == 'RECORDS SERVICES':
                conn.execute(
                    "UPDATE routing_records SET status = 'released' "
                    "WHERE code = ? AND "
                    "LOWER(status) != 'with corrections'",
                    (code,)
                )
                break

        conn.commit()

        affected = conn.execute("SELECT id FROM routing_records WHERE code = ?", (code,)).fetchall()
        affected_ids = [r['id'] for r in affected]
        conn.close()

        if affected_ids:
            db_manager.enqueue_sync_batch('routing_records', affected_ids)
        db_manager.trigger_excel_export()

        return jsonify({'status': 'success', 'message': f'Updated {len(affected_ids)} record(s).'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/delete_record/<int:record_id>', methods=['DELETE'])
@require_admin
def delete_record(record_id):
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    try:
        conn = db_manager.get_db_connection()
        row = conn.execute("SELECT code, employee FROM routing_records WHERE id = ?", (record_id,)).fetchone()
        code = row['code'] if row else None
        employee = row['employee'] if row else None
        conn.execute("DELETE FROM routing_records WHERE id = ?", (record_id,))
        conn.commit()
        conn.close()

        # Synchronously delete from GSheets so periodic pull doesn't re-import
        _sync_delete_from_gsheets(code, employee)

        db_manager.trigger_excel_export()
        return jsonify({'status': 'success', 'message': 'Record deleted.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


def _sync_delete_from_gsheets(code, employee=None):
    """Helper to synchronously delete a code from GSheets via the db_manager.
    Prevents re-import by periodic pull."""
    if not code:
        return
    sheet_id = db_manager.get_setting('gsheets_id')
    creds = db_manager.get_decrypted_setting('gsheets_credentials')
    if not sheet_id or not creds:
        return
    try:
        db_manager.synchronous_delete_from_gsheets(sheet_id, creds, code)
    except Exception as e:
        print(f"synchronous delete from GSheets failed for {code}: {e}")


@app.route('/api/delete_code/<path:code>', methods=['DELETE'])
@require_admin
def delete_code(code):
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    try:
        conn = db_manager.get_db_connection()
        affected = conn.execute(
            "SELECT id, code, employee FROM routing_records WHERE code = ?", (code,)
        ).fetchall()
        affected_ids = [r['id'] for r in affected]

        conn.execute("DELETE FROM routing_records WHERE code = ?", (code,))
        conn.execute("DELETE FROM code_lookup WHERE code = ?", (code,))
        conn.execute("DELETE FROM code_employees WHERE code = ?", (code,))
        conn.commit()
        conn.close()

        # Synchronously delete from GSheets
        _sync_delete_from_gsheets(code)

        db_manager.trigger_excel_export()

        return jsonify({
            'status': 'success',
            'message': f'Deleted code {code} and {len(affected_ids)} routing record(s).'
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/batch_delete', methods=['POST'])
@require_admin
def batch_delete():
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    try:
        data = request.get_json(force=True) or {}
        codes = data.get('codes', [])
        record_ids = data.get('record_ids', [])

        if not codes and not record_ids:
            return jsonify({'status': 'error', 'message': 'No items to delete.'}), 400

        conn = db_manager.get_db_connection()
        total = 0

        for rid in record_ids:
            row = conn.execute("SELECT code, employee FROM routing_records WHERE id = ?", (rid,)).fetchone()
            rcode = row['code'] if row else None
            remp = row['employee'] if row else None
            conn.execute("DELETE FROM routing_records WHERE id = ?", (rid,))
            _sync_delete_from_gsheets(rcode, remp)
            total += 1

        for code in codes:
            affected = conn.execute(
                "SELECT id, code, employee FROM routing_records WHERE code = ?", (code,)
            ).fetchall()
            conn.execute("DELETE FROM routing_records WHERE code = ?", (code,))
            conn.execute("DELETE FROM code_lookup WHERE code = ?", (code,))
            conn.execute("DELETE FROM code_employees WHERE code = ?", (code,))
            _sync_delete_from_gsheets(code)
            total += 1 + len(affected)

        conn.commit()
        conn.close()
        db_manager.trigger_excel_export()

        return jsonify({
            'status': 'success',
            'message': f'Deleted {total} item(s).'
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ─────────────────────────────────────────────
#  FORCE SYNC
# ─────────────────────────────────────────────
@app.route('/api/force_sync', methods=['POST'])
@require_admin
def force_sync():
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    try:
        data = request.get_json(force=True) or {}
        direction = data.get('direction', 'pull').strip().lower()
        if direction not in ('pull', 'push'):
            return jsonify({'status': 'error', 'message': 'direction must be "pull" or "push"'}), 400
        if direction == 'pull':
            count, errors = db_manager.force_pull_from_gsheets()
        else:
            count, errors = db_manager.force_push_to_gsheets()
            errors = errors if isinstance(errors, list) else []
        return jsonify({
            'status': 'success' if not errors else 'partial',
            'message': f"Force {direction} complete. Synced {count} records." if not errors else f"Force {direction} completed with {len(errors)} error(s).",
            'count': count,
            'errors': errors
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ─────────────────────────────────────────────
#  REVIVE CODE
# ─────────────────────────────────────────────
@app.route('/api/revive_code', methods=['POST'])
@require_admin
def revive_code():
    blocked, msg = _check_maintenance()
    if blocked:
        return jsonify({'status': 'error', 'message': msg}), 403
    try:
        data = request.get_json(force=True) or {}
        code = (data.get('code') or '').strip()
        doc_type = (data.get('doc_type') or '').strip()
        school = (data.get('school') or '').strip()
        employee_names = data.get('employees', [])
        if not isinstance(employee_names, list):
            employee_names = [employee_names]
        if not code or not doc_type:
            return jsonify({'status': 'error', 'message': 'code and doc_type are required.'}), 400
        conn = db_manager.get_db_connection()
        conn.execute("DELETE FROM routing_records WHERE code = ?", (code,))
        conn.execute("DELETE FROM code_employees WHERE code = ?", (code,))
        existing = conn.execute("SELECT code FROM code_lookup WHERE code = ?", (code,)).fetchone()
        today = datetime.now().strftime('%Y-%m-%d')
        if existing:
            conn.execute("UPDATE code_lookup SET doc_type = ?, date_generated = ? WHERE code = ?", (doc_type, today, code))
        else:
            conn.execute("INSERT INTO code_lookup (code, doc_type, date_generated) VALUES (?, ?, ?)", (code, doc_type, today))
        for emp_name in employee_names:
            emp_name = emp_name.strip()
            if not emp_name:
                continue
            emp = conn.execute("SELECT id FROM employees WHERE name = ?", (emp_name,)).fetchone()
            if emp:
                employee_id = emp['id']
            else:
                school_id = None
                if school:
                    school_row = conn.execute("SELECT id FROM schools WHERE name = ?", (school,)).fetchone()
                    if school_row:
                        school_id = school_row['id']
                conn.execute("INSERT INTO employees (name, school_id) VALUES (?, ?)", (emp_name, school_id))
                employee_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute("INSERT OR IGNORE INTO code_employees (code, employee_id, employee_name) VALUES (?, ?, ?)", (code, employee_id, emp_name))
        conn.commit()
        conn.close()
        sheet_id = db_manager.get_setting('gsheets_id')
        creds = db_manager.get_decrypted_setting('gsheets_credentials')
        if sheet_id and creds:
            try:
                db_manager.synchronous_delete_from_gsheets(sheet_id, creds, code)
            except Exception as e:
                print(f"revive_code: GSheets delete error: {e}")
            db_manager.enqueue_sync('code_lookup', ref_id=code, code=code, operation='upsert')
            for emp_name in employee_names:
                db_manager.enqueue_sync('code_employees', ref_id=code, code=code, employee=emp_name, operation='upsert')
        db_manager.trigger_excel_export()
        return jsonify({
            'status': 'success',
            'message': f'Code {code} revived with {len(employee_names)} employee(s).'
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ─────────────────────────────────────────────
#  SETTINGS (Admin only)
# ─────────────────────────────────────────────
@app.route('/api/get_settings', methods=['GET'])
@require_login
def get_settings():
    try:
        # Never expose the actual credentials to the frontend
        has_creds = bool(db_manager.get_setting('gsheets_credentials', ''))
        # Show whether credentials come from environment variables (persistent) or SQLite (wiped on Render)
        creds_from_env = bool(os.environ.get('GSHEETS_CREDENTIALS'))
        id_from_env = bool(os.environ.get('GSHEETS_ID'))
        user_role = session.get('user_role', '')
        is_admin = user_role == 'Admin'
        return jsonify({
            'status':               'success',
            'gsheets_enabled':      db_manager.get_setting('gsheets_enabled', 'False') == 'True',
            'gsheets_id':           db_manager.get_setting('gsheets_id', ''),
            'gsheets_configured':   has_creds,
            'gsheets_from_env':     creds_from_env and id_from_env,  # True = survives Render wipes
            'scanner_pin':          db_manager.get_setting('scanner_pin', 'scanner123') if is_admin else '',
            'admin_pin':            db_manager.get_setting('admin_pin', 'admin123') if is_admin else '',
            'gsheets_pull_enabled': db_manager.get_setting('gsheets_pull_enabled', 'True') == 'True',
            'gsheets_pull_interval': db_manager.get_setting('gsheets_pull_interval', '5'),
            'logo_url': db_manager.get_setting('logo_url', ''),
            'header_line_1': db_manager.get_setting('header_line_1', 'Republic of the Philippines'),
            'header_line_2': db_manager.get_setting('header_line_2', 'Department of Education - NCR'),
            'header_line_3': db_manager.get_setting('header_line_3', 'Schools Division Office of Manila'),
            'system_title': db_manager.get_setting('system_title', 'Document Tracking System'),
            'system_title_font': db_manager.get_setting('system_title_font', ''),
            'custom_font_url': db_manager.get_setting('custom_font_url', ''),
            'custom_font_format': db_manager.get_setting('custom_font_format', ''),
            'title_letter_spacing': db_manager.get_setting('title_letter_spacing', '0'),
            'title_bg_url': db_manager.get_setting('title_bg_url', ''),
            'title_bg_opacity': db_manager.get_setting('title_bg_opacity', '7'),
            'title_glow_enabled': db_manager.get_setting('title_glow_enabled', 'True') == 'True',
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/upload_logo', methods=['POST'])
@require_admin
def upload_logo():
    try:
        # Handle logo removal
        if request.is_json and request.json.get('remove'):
            current_logo = db_manager.get_setting('logo_url', '')
            if current_logo:
                file_path = os.path.join('static', 'uploads', os.path.basename(current_logo))
                if os.path.exists(file_path):
                    os.remove(file_path)
            db_manager.save_setting('logo_url', '')
            return jsonify({'status': 'success', 'message': 'Logo removed.'})

        # Handle logo upload
        if 'logo' not in request.files:
            return jsonify({'status': 'error', 'message': 'No file provided.'}), 400
        file = request.files['logo']
        if file.filename == '':
            return jsonify({'status': 'error', 'message': 'No file selected.'}), 400

        # Validate file type
        allowed = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
        ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
        if ext not in allowed:
            return jsonify({'status': 'error', 'message': f'Invalid file type: .{ext}. Allowed: {", ".join(allowed)}'}), 400

        # Save to static/uploads/
        upload_dir = os.path.join('static', 'uploads')
        if not os.path.exists(upload_dir):
            os.makedirs(upload_dir)

        from werkzeug.utils import secure_filename
        safe_name = secure_filename(f'logo_{datetime.now().strftime("%Y%m%d%H%M%S")}.{ext}')
        file_path = os.path.join(upload_dir, safe_name)
        file.save(file_path)

        logo_url = f'/static/uploads/{safe_name}'
        db_manager.save_setting('logo_url', logo_url)

        return jsonify({'status': 'success', 'message': 'Logo uploaded.', 'logo_url': logo_url})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/upload_font', methods=['POST'])
@require_admin
def upload_font():
    try:
        if 'font' not in request.files:
            return jsonify({'status': 'error', 'message': 'No file provided.'}), 400
        file = request.files['font']
        if file.filename == '':
            return jsonify({'status': 'error', 'message': 'No file selected.'}), 400

        allowed = {'ttf', 'woff', 'woff2', 'otf'}
        ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
        if ext not in allowed:
            return jsonify({'status': 'error', 'message': f'Invalid font type: .{ext}. Allowed: {", ".join(allowed)}'}), 400

        font_dir = os.path.join('static', 'uploads', 'fonts')
        if not os.path.exists(font_dir):
            os.makedirs(font_dir)

        from werkzeug.utils import secure_filename
        safe_name = secure_filename(f'custom_font_{datetime.now().strftime("%Y%m%d%H%M%S")}.{ext}')
        file_path = os.path.join(font_dir, safe_name)
        file.save(file_path)

        font_url = f'/static/uploads/fonts/{safe_name}'
        db_manager.save_setting('custom_font_url', font_url)
        db_manager.save_setting('custom_font_format', ext)

        return jsonify({'status': 'success', 'message': 'Font uploaded!', 'font_url': font_url})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/remove_font', methods=['POST'])
@require_admin
def remove_font():
    try:
        current_font = db_manager.get_setting('custom_font_url', '')
        if current_font:
            file_path = os.path.join('static', 'uploads', 'fonts', os.path.basename(current_font))
            if os.path.exists(file_path):
                os.remove(file_path)
        db_manager.save_setting('custom_font_url', '')
        db_manager.save_setting('custom_font_format', '')
        return jsonify({'status': 'success', 'message': 'Custom font removed.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/upload_title_bg', methods=['POST'])
@require_login
def upload_title_bg():
    try:
        if request.is_json and request.json.get('remove'):
            current = db_manager.get_setting('title_bg_url', '')
            if current:
                fp = os.path.join('static', 'uploads', os.path.basename(current))
                if os.path.exists(fp): os.remove(fp)
            db_manager.save_setting('title_bg_url', '')
            return jsonify({'status': 'success', 'message': 'Background removed.'})

        if 'image' not in request.files:
            return jsonify({'status': 'error', 'message': 'No file provided.'}), 400
        file = request.files['image']
        if file.filename == '':
            return jsonify({'status': 'error', 'message': 'No file selected.'}), 400

        allowed = {'png', 'jpg', 'jpeg', 'gif', 'webp', 'svg'}
        ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
        if ext not in allowed:
            return jsonify({'status': 'error', 'message': f'Invalid type: .{ext}'}), 400

        upload_dir = os.path.join('static', 'uploads')
        if not os.path.exists(upload_dir): os.makedirs(upload_dir)
        from werkzeug.utils import secure_filename
        safe = secure_filename(f'title_bg_{datetime.now().strftime("%Y%m%d%H%M%S")}.{ext}')
        fp = os.path.join(upload_dir, safe)
        file.save(fp)
        url = f'/static/uploads/{safe}'
        db_manager.save_setting('title_bg_url', url)
        return jsonify({'status': 'success', 'message': 'Background uploaded!', 'url': url})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/save_settings', methods=['POST'])
@require_admin
def save_settings():
    try:
        data = request.json
        db_manager.save_setting('gsheets_enabled', 'True' if data.get('gsheets_enabled') else 'False')

        if data.get('header_line_1', '').strip():
            db_manager.save_setting('header_line_1', data['header_line_1'].strip())
        if data.get('header_line_2', '').strip():
            db_manager.save_setting('header_line_2', data['header_line_2'].strip())
        if data.get('header_line_3', '').strip():
            db_manager.save_setting('header_line_3', data['header_line_3'].strip())
        if data.get('system_title', '').strip():
            db_manager.save_setting('system_title', data['system_title'].strip())

        if 'system_title_font' in data:
            db_manager.save_setting('system_title_font', data['system_title_font'].strip())
        if 'title_letter_spacing' in data:
            db_manager.save_setting('title_letter_spacing', data['title_letter_spacing'].strip())
        if 'title_bg_opacity' in data:
            db_manager.save_setting('title_bg_opacity', data['title_bg_opacity'].strip())
        if 'title_glow_enabled' in data:
            db_manager.save_setting('title_glow_enabled', str(data['title_glow_enabled']))

        if data.get('scanner_pin', '').strip():
            db_manager.save_setting('scanner_pin', data['scanner_pin'].strip())
        if data.get('admin_pin', '').strip():
            db_manager.save_setting('admin_pin', data['admin_pin'].strip())

        if 'gsheets_pull_enabled' in data:
            db_manager.save_setting('gsheets_pull_enabled',
                                    'True' if data['gsheets_pull_enabled'] else 'False')

        if data.get('gsheets_pull_interval', '').strip():
            val = data['gsheets_pull_interval'].strip()
            try:
                parsed = int(val)
                if parsed >= 1:
                    db_manager.save_setting('gsheets_pull_interval', str(parsed))
                    db_manager.reset_periodic_pull_timer()
            except ValueError:
                pass

        return jsonify({'status': 'success', 'message': 'Settings saved!'})
    except ValueError as ve:
        return jsonify({'status': 'error', 'message': str(ve)})
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

@app.route('/api/rebuild_code_employees', methods=['POST'])
@require_admin
def rebuild_code_employees():
    """Rebuild code_employees links from code_lookup against current employees/schools."""
    try:
        ce_count, rr_count, unmatched = db_manager.rebuild_code_employees()
        msg = f'Rebuilt {ce_count} code-employee links, relinked {rr_count} routing records.'
        if unmatched:
            msg += f' {unmatched} unmatched names skipped (already logged).'
        return jsonify({'status': 'success', 'message': msg, 'code_employees': ce_count, 'routing_records': rr_count, 'unmatched': unmatched})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/replace_employees_schools', methods=['POST'])
@require_admin
def replace_employees_schools():
    """Replace schools + employees tables from GSheets worksheets,
    auto-insert missing employees from code_lookup, rebuild links."""
    try:
        result = db_manager.replace_schools_employees_from_gsheets()
        msg = (
            f"Schools: {result['schools_inserted']} inserted. "
            f"Employees: {result['employees_inserted']} from GSheets + {result['missing_inserted']} auto-inserted. "
            f"Code-employee links: {result['code_employees']} rebuilt. "
            f"Routing records: {result['routing_records']} relinked."
        )
        if result['unmatched']:
            msg += f" {result['unmatched']} names still unmatched."
        return jsonify({'status': 'success', 'message': msg, **result})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/gsheet_status', methods=['GET'])
@require_login
def gsheet_status():
    """Quick-test the currently saved Google Sheets connection."""
    try:
        enabled  = db_manager.get_setting('gsheets_enabled', 'False')
        sheet_id = db_manager.get_setting('gsheets_id', '')
        # Use decrypted credentials for the connection test
        creds    = db_manager.get_decrypted_setting('gsheets_credentials')
        if enabled != 'True' or not sheet_id or not creds:
            return jsonify({'status': 'disabled', 'message': 'Google Sheets sync is not configured.'})
        result = db_manager.test_google_sheets_connection(sheet_id, creds)
        return jsonify(result)
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/gsheets_last_pull', methods=['GET'])
@require_login
def gsheets_last_pull():
    return jsonify(db_manager.get_last_pull_info())


@app.route('/api/flush_sync', methods=['POST'])
def flush_sync():
    """Block until the GSheets sync queue is empty — used by beforeunload sendBeacon."""
    try:
        db_manager.gs_sync_queue.join()
        return '', 204
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/bulk_sync', methods=['POST'])
@require_admin
def bulk_sync():
    """Push ALL existing data (routing records, codes, users, master_data) to Google Sheets at once."""
    try:
        if db_manager.get_setting('maintenance_mode') == 'True':
            return jsonify({'status': 'error', 'message': 'Bulk sync is disabled during maintenance mode.'}), 409
        count, errors = db_manager.bulk_sync_to_gsheets()
        if errors and count == 0:
            return jsonify({
                'status': 'error',
                'message': f'Bulk sync failed: {errors[0][:200]}',
                'errors': errors[:10]
            })
        msg = f'Synced {count} routing records + codes, users, and master data to Google Sheets.'
        if errors:
            msg += f' {len(errors)} had errors (showing first 5): ' + '; '.join(errors[:5])
        return jsonify({
            'status': 'success' if not errors else 'partial',
            'message': msg,
            'synced': count,
            'errors': errors[:10]
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/sync/pull_from_gsheets', methods=['POST'])
@require_admin
def pull_from_gsheets():
    """Manually trigger a full Google Sheets → SQLite restore.
    Useful for recovering data after a wipe or for debugging sync issues."""
    try:
        if db_manager.get_setting('maintenance_mode') == 'True':
            return jsonify({'status': 'error', 'message': 'Use "Turn Off Maintenance → Import Changes" instead. Direct pull is blocked during maintenance.'}), 409
        count, errors = db_manager.pull_all_from_gsheets()
        if errors and count == 0:
            return jsonify({
                'status': 'error',
                'message': f'Pull from GSheets failed: {errors[0][:300]}',
                'errors': errors
            })
        msg = f'Restored {count} routing records from Google Sheets (+ codes, users, master data).'
        if errors:
            msg += ' Warnings: ' + '; '.join(errors[:5])
        return jsonify({
            'status': 'success' if not errors else 'partial',
            'message': msg,
            'restored': count,
            'errors': errors[:10]
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/sync/reference', methods=['GET'])
@require_login
def sync_reference():
    """Pull reference data (code_lookup, master_data) from GSheets.
    Used by the scanner tab on load so GSheets edits are available before scanning.
    Lightweight upsert — does not wipe SQLite."""
    if not session.get('user_role'):
        return jsonify({'status': 'error', 'message': 'Not authenticated'}), 401
    success, error = db_manager.pull_reference_from_gsheets()
    if success:
        return jsonify({'status': 'success', 'message': 'Reference data synced from Google Sheets.'})
    return jsonify({'status': 'warning', 'message': f'Reference sync skipped or failed: {error}'}), 200


@app.route('/api/admin/maintenance/toggle', methods=['POST'])
@require_admin
def maintenance_toggle():
    """Toggle maintenance mode on/off. When ON, only admins can scan/generate,
    and the periodic pull + background sync worker are paused so the admin
    can safely edit data in Google Sheets without surprise syncs.

    Accepts optional {"gsheets_enabled": false} to disable sync atomically
    (avoids the two-fetch race that caused 'database is locked').

    When turning OFF, the request may include {"pull_gsheets": true} to first
    pull GSheets → SQLite (importing any edits made directly in GSheets)."""
    try:
        data = request.get_json(silent=True) or {}
        current = db_manager.get_setting('maintenance_mode', 'False')
        new_val = 'False' if current == 'True' else 'True'

        # Atomically disable sync when turning maintenance ON
        if new_val == 'True' and 'gsheets_enabled' in data:
            db_manager.save_setting('gsheets_enabled',
                                    'True' if data['gsheets_enabled'] else 'False')

        if new_val == 'False' and data.get('pull_gsheets'):
            result, errors = db_manager.pull_all_from_gsheets()
            if errors:
                return jsonify({'status': 'error', 'message': f'GSheets import failed: {errors[0]}'}), 500
            db_manager.update_last_pull_info("ok", "GSheets imported before exiting maintenance mode.")

        if new_val == 'False':
            db_manager.reset_periodic_pull_timer()

        db_manager.save_setting('maintenance_mode', new_val)

        if new_val == 'True':
            db_manager.update_last_pull_info("paused", "Maintenance mode activated — periodic pull paused.")

        return jsonify({
            'status': 'success',
            'maintenance_mode': new_val == 'True',
            'message': 'Maintenance mode ' + ('enabled.' if new_val == 'True' else 'disabled.')
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/admin/maintenance/status', methods=['GET'])
@require_login
def maintenance_status():
    """Return current maintenance mode status."""
    enabled = db_manager.get_setting('maintenance_mode', 'False') == 'True'
    role = session.get('user_role', '')
    return jsonify({
        'maintenance_mode': enabled,
        'is_admin': role.lower() == 'admin'
    })


@app.route('/api/admin/reconcile_employee_names', methods=['POST'])
@require_admin
def reconcile_employee_names():
    """Rewrite all routing_records.employee from 'LastName, FirstName' to 'FirstName LastName'.
    Then trigger Excel export + queue GSheets sync for the rewritten rows."""
    try:
        rewritten, total, samples = db_manager.migrate_routing_employee_names()
        msg = f'Rewritten {rewritten}/{total} employee names in routing records.'
        if samples:
            msg += ' Samples: ' + '; '.join(samples)
        # Trigger Excel export and queue all routing records for GSheets resync
        db_manager.trigger_excel_export()
        conn = db_manager.get_db_connection()
        all_ids = conn.execute("SELECT id FROM routing_records").fetchall()
        conn.close()
        db_manager.queue_sync_batch([r['id'] for r in all_ids])
        return jsonify({'status': 'success', 'message': msg, 'rewritten': rewritten, 'total': total, 'samples': samples})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/admin/migrate_master_to_gsheets', methods=['POST'])
@require_admin
def migrate_master_to_gsheets():
    """Push all 3 raw sheets from master_db.xlsx to Google Sheets as separate worksheets
    ('Master Data Sheet1', 'Master Data Sheet2', 'Master Data Sheet3')."""
    try:
        if db_manager.get_setting('maintenance_mode') == 'True':
            return jsonify({'status': 'error', 'message': 'Cannot migrate during maintenance mode.'}), 409

        enabled  = db_manager.get_setting('gsheets_enabled') == 'True'
        sheet_id = db_manager.get_setting('gsheets_id')
        creds    = db_manager.get_decrypted_setting('gsheets_credentials')
        if not enabled or not sheet_id or not creds:
            return jsonify({'status': 'error', 'message': 'Google Sheets is not configured. Enable it and save credentials first.'}), 400

        success = db_manager._gsheets_push_with_retry(
            db_manager.push_master_db_to_gsheets, sheet_id, creds
        )
        if success:
            return jsonify({
                'status': 'success',
                'message': 'Successfully pushed master_db.xlsx (Sheet1, Sheet2, Sheet3) to Google Sheets.',
            })
        else:
            return jsonify({'status': 'error', 'message': 'Migration failed after 3 retries. Check Render logs for details.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ─────────────────────────────────────────────
#  DASHBOARD STATS
# ─────────────────────────────────────────────
@app.route('/api/dashboard_stats', methods=['GET'])
@require_login
def dashboard_stats():
    """Return dashboard statistics. Query params:
       doc_type — filter by document type
       count_mode — 'schools' (default, 1 per school+doc_type) or 'rows' (per employee record)"""
    try:
        doc_type_filter = request.args.get('doc_type', '').strip()
        count_mode = request.args.get('count_mode', 'schools').strip()
        is_schools = count_mode == 'schools'

        conn = db_manager.get_db_connection()

        # Build common JOIN / WHERE fragments
        doc_join = " INNER JOIN code_lookup cl_ ON rr.code = cl_.code"
        doc_where = ""
        status_filter = " AND LOWER(rr.status) != 'deleted'"
        params = []
        if doc_type_filter:
            doc_where = " AND cl_.doc_type = ?"
            params.append(doc_type_filter)

        # Helper: build COUNT expression for status/doc_type metrics
        def count_expr(col_prefix):
            if is_schools:
                return f"COUNT(DISTINCT rr.school_office || '|' || COALESCE({col_prefix}, ''))"
            else:
                return "COUNT(*)"

        # Unique barcodes (from code_lookup)
        if doc_type_filter:
            unique_barcodes = conn.execute(
                "SELECT COUNT(*) FROM code_lookup WHERE doc_type = ?", [doc_type_filter]
            ).fetchone()[0]
        else:
            unique_barcodes = conn.execute("SELECT COUNT(*) FROM code_lookup").fetchone()[0]

        # Total employees
        total_q = f"SELECT COUNT(*) FROM routing_records rr{doc_join} WHERE 1=1{status_filter}{doc_where}"
        total_employees = conn.execute(total_q, params).fetchone()[0]

        # Active codes
        active_q = f"SELECT COUNT(DISTINCT rr.code) FROM routing_records rr{doc_join} WHERE 1=1{status_filter}{doc_where}"
        active_codes = conn.execute(active_q, params).fetchone()[0]

        # Released count
        released_q = f"SELECT {count_expr('cl_.doc_type')} FROM routing_records rr{doc_join} WHERE LOWER(rr.status) = 'released'{status_filter}{doc_where}"
        released_count = conn.execute(released_q, params).fetchone()[0]

        # For signature count
        sig_q = f"SELECT {count_expr('cl_.doc_type')} FROM routing_records rr{doc_join} WHERE LOWER(rr.status) = 'for signature'{status_filter}{doc_where}"
        for_signature_count = conn.execute(sig_q, params).fetchone()[0]

        # With corrections count
        corr_q = f"SELECT {count_expr('cl_.doc_type')} FROM routing_records rr{doc_join} WHERE LOWER(rr.status) = 'with corrections'{status_filter}{doc_where}"
        with_corrections_count = conn.execute(corr_q, params).fetchone()[0]

        # Status breakdown
        status_breakdown = []
        if is_schools:
            status_q = f"""
                SELECT LOWER(rr.status) as status,
                       COUNT(DISTINCT rr.school_office || '|' || COALESCE(cl_.doc_type, '')) as cnt
                FROM routing_records rr{doc_join}
                WHERE 1=1{status_filter}{doc_where}
                GROUP BY LOWER(rr.status)
            """
        else:
            status_q = f"""
                SELECT LOWER(rr.status) as status, COUNT(*) as cnt
                FROM routing_records rr{doc_join}
                WHERE 1=1{status_filter}{doc_where}
                GROUP BY LOWER(rr.status)
            """
        rows = conn.execute(status_q, params).fetchall()
        for row in rows:
            status_breakdown.append({'status': row['status'], 'count': row['cnt']})

        # Doc type breakdown
        if is_schools:
            dt_q = f"""
                SELECT COALESCE(cl_.doc_type, '') as doc_type,
                       COUNT(DISTINCT rr.school_office || '|' || COALESCE(cl_.doc_type, '')) as cnt
                FROM routing_records rr{doc_join}
                WHERE 1=1{status_filter}{doc_where} AND cl_.doc_type IS NOT NULL AND cl_.doc_type != ''
                GROUP BY cl_.doc_type
                ORDER BY cnt DESC LIMIT 10
            """
        else:
            dt_q = f"""
                SELECT COALESCE(cl_.doc_type, '') as doc_type, COUNT(*) as cnt
                FROM routing_records rr{doc_join}
                WHERE 1=1{status_filter}{doc_where} AND cl_.doc_type IS NOT NULL AND cl_.doc_type != ''
                GROUP BY cl_.doc_type
                ORDER BY cnt DESC LIMIT 10
            """
        doc_type_breakdown = []
        rows = conn.execute(dt_q, params).fetchall()
        for row in rows:
            label = row['doc_type'] if row['doc_type'] else '(not set)'
            doc_type_breakdown.append({'label': label, 'count': row['cnt']})

        conn.close()
        return jsonify({
            'status': 'success',
            'unique_barcodes': unique_barcodes,
            'total_employees': total_employees,
            'active_codes': active_codes,
            'released_count': released_count,
            'for_signature_count': for_signature_count,
            'with_corrections_count': with_corrections_count,
            'status_breakdown': status_breakdown,
            'doc_type_breakdown': doc_type_breakdown
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ─────────────────────────────────────────────
#  RECEIVING OFFICES (for dashboard filter)
# ─────────────────────────────────────────────
@app.route('/api/get_receiving_offices', methods=['GET'])
@require_login
def get_receiving_offices():
    """Return distinct non-empty receiving offices from all 10 columns."""
    try:
        conn = db_manager.get_db_connection()
        parts = []
        for i in range(1, 11):
            parts.append(
                f"SELECT receiving_office_{i} AS office FROM routing_records "
                f"WHERE receiving_office_{i} IS NOT NULL AND receiving_office_{i} != ''"
            )
        sql = "SELECT DISTINCT TRIM(office) AS office FROM (" + " UNION ALL ".join(parts) + ") WHERE office IS NOT NULL AND office != '' ORDER BY office"
        rows = conn.execute(sql).fetchall()
        conn.close()
        offices = [r['office'] for r in rows]
        return jsonify({'status': 'success', 'offices': offices})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ─────────────────────────────────────────────
#  RECENT ACTIVITY
# ─────────────────────────────────────────────
@app.route('/api/recent_activity', methods=['GET'])
@require_login
def recent_activity():
    """Return the 10 most recently scanned codes with employee count and latest office.
    Query params: doc_type — filter by document type, count_mode — 'rows' or 'schools',
    last_office — filter by the last scanning office."""
    try:
        doc_type_filter = request.args.get('doc_type', '').strip()
        count_mode = request.args.get('count_mode', 'rows').strip()
        last_office_filter = request.args.get('last_office', '').strip()
        is_schools = count_mode == 'schools'
        conn = db_manager.get_db_connection()

        # Build COALESCE chains to find the LAST (highest-index non-empty) scan per record
        office_coalesce = "COALESCE(" + ", ".join(f"NULLIF(receiving_office_{i}, '')" for i in range(10, 0, -1)) + ")"
        ts_coalesce = "COALESCE(" + ", ".join(f"NULLIF(timestamp_{i}, '')" for i in range(10, 0, -1)) + ")"

        # Base subquery: each record's last scan info
        last_scan_sql = f"""
            SELECT code, employee, school_office,
                   {office_coalesce} AS last_office,
                   {ts_coalesce} AS last_timestamp
            FROM routing_records
            WHERE {office_coalesce} IS NOT NULL
        """

        params = []
        if last_office_filter:
            last_scan_sql += f" AND {office_coalesce} = ?"
            params.append(last_office_filter)

        doc_where_inner = ""
        doc_where_counts = ""
        if doc_type_filter:
            doc_where_inner = " AND cl.doc_type = ?"
            doc_where_counts = " AND cl2.doc_type = ?"
            params.append(doc_type_filter)
            params.append(doc_type_filter)

        # Use CTE + ROW_NUMBER to pick the single latest scan row per group,
        # ensuring office, employee, and timestamp come from the same record.
        if is_schools:
            group_sql = f"""
                WITH last_scans AS ({last_scan_sql})
                SELECT ranked.school_office, ranked.doc_type,
                       counts.total_codes, ranked.employee AS first_employee,
                       counts.total_employees, ranked.last_office AS office,
                       ranked.last_timestamp
                FROM (
                    SELECT ls.school_office, cl.doc_type, ls.employee, ls.last_office, ls.last_timestamp,
                           ROW_NUMBER() OVER (PARTITION BY ls.school_office, cl.doc_type ORDER BY ls.last_timestamp DESC) AS rn
                    FROM last_scans ls
                    JOIN code_lookup cl ON ls.code = cl.code
                    WHERE 1=1 {doc_where_inner}
                ) ranked
                JOIN (
                    SELECT ls2.school_office, cl2.doc_type,
                           COUNT(DISTINCT ls2.code) AS total_codes,
                           COUNT(DISTINCT ls2.employee) AS total_employees
                    FROM last_scans ls2
                    JOIN code_lookup cl2 ON ls2.code = cl2.code
                    WHERE 1=1 {doc_where_counts}
                    GROUP BY ls2.school_office, cl2.doc_type
                ) counts ON counts.school_office = ranked.school_office AND counts.doc_type = ranked.doc_type
                WHERE ranked.rn = 1
                ORDER BY ranked.last_timestamp DESC LIMIT 10
            """
        else:
            group_sql = f"""
                WITH last_scans AS ({last_scan_sql})
                SELECT ranked.code,
                       ranked.employee AS first_employee,
                       counts.total_employees,
                       ranked.school_office AS origin_school,
                       ranked.last_office AS office,
                       ranked.last_timestamp
                FROM (
                    SELECT ls.code, ls.employee, ls.school_office, ls.last_office, ls.last_timestamp,
                           ROW_NUMBER() OVER (PARTITION BY ls.code ORDER BY ls.last_timestamp DESC) AS rn
                    FROM last_scans ls
                    JOIN code_lookup cl ON ls.code = cl.code
                    WHERE 1=1 {doc_where_inner}
                ) ranked
                JOIN (
                    SELECT ls2.code,
                           COUNT(DISTINCT ls2.employee) AS total_employees
                    FROM last_scans ls2
                    JOIN code_lookup cl2 ON ls2.code = cl2.code
                    WHERE 1=1 {doc_where_counts}
                    GROUP BY ls2.code
                ) counts ON counts.code = ranked.code
                WHERE ranked.rn = 1
                ORDER BY ranked.last_timestamp DESC LIMIT 10
            """
        rows = conn.execute(group_sql, params).fetchall()
        activities = []
        for row in rows:
            if is_schools:
                school = row['school_office'] or ''
                dt = row['doc_type'] or ''
                group_label = (school + ' \\u2013 ' + dt) if dt else school
                activities.append({
                    'code': '',
                    'group_label': group_label,
                    'doc_type': dt,
                    'total_codes': row['total_codes'],
                    'first_employee': row['first_employee'] or '',
                    'total_employees': row['total_employees'],
                    'origin_school': school,
                    'office': row['office'] or '',
                    'last_timestamp': row['last_timestamp'] or ''
                })
            else:
                activities.append({
                    'code': row['code'],
                    'group_label': '',
                    'doc_type': '',
                    'total_codes': 1,
                    'first_employee': row['first_employee'] or '',
                    'total_employees': row['total_employees'],
                    'origin_school': row['origin_school'] or '',
                    'office': row['office'] or '',
                    'last_timestamp': row['last_timestamp'] or ''
                })
        conn.close()
        return jsonify({'status': 'success', 'activities': activities})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ── Forgot Password ──
@app.route('/api/forgot_password', methods=['POST'])
@limiter.limit("3 per minute")
def forgot_password():
    """
    Forgot password: if user exists with an email, generate a temporary password,
    save its hash with requires_password_change=1, and attempt to email it.
    """
    try:
        data = request.json
        username = data.get('username', '').strip().lower()
        
        if not username:
            return jsonify({'status': 'error', 'message': 'Username is required.'}), 400
        
        conn = db_manager.get_db_connection()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        
        if not user:
            conn.close()
            return jsonify({'status': 'error', 'message': 'Username not found.'}), 404
        
        email = user['email'] if 'email' in user.keys() and user['email'] else ''
        if not email:
            conn.close()
            return jsonify({'status': 'error', 'message': 'No email address on file for this account. Contact your administrator.'}), 400
        
        import secrets
        temp_password = secrets.token_urlsafe(8)
        temp_hash = bcrypt.hashpw(temp_password.encode(), bcrypt.gensalt()).decode()
        
        conn.execute("UPDATE users SET password_hash = ?, requires_password_change = 1 WHERE username = ?",
                     (temp_hash, username))
        conn.commit()
        conn.close()
        
        smtp_enabled = db_manager.get_setting('smtp_enabled', 'False')
        smtp_server = db_manager.get_setting('smtp_server', '')
        smtp_port = db_manager.get_setting('smtp_port', '587')
        smtp_user = db_manager.get_setting('smtp_user', '')
        smtp_pass = db_manager.get_decrypted_setting('smtp_password', '')
        smtp_from = db_manager.get_setting('smtp_from', '')
        
        email_sent = False
        if smtp_enabled == 'True' and smtp_server and smtp_user and smtp_pass:
            try:
                import smtplib
                from email.message import EmailMessage
                
                msg = EmailMessage()
                msg.set_content(f"""Dear {username},

A password reset was requested for your Document Tracking System account.

Your temporary password is: {temp_password}

Please log in and change your password immediately.

This is an automated message. If you did not request this, please contact your administrator.

Regards,
Document Tracking System
""")
                msg['Subject'] = 'Document Tracking System - Password Reset'
                msg['From'] = smtp_from or smtp_user
                msg['To'] = email
                
                server = smtplib.SMTP(smtp_server, int(smtp_port))
                server.starttls()
                server.login(smtp_user, smtp_pass)
                server.send_message(msg)
                server.quit()
                email_sent = True
            except Exception as smtp_err:
                print(f"SMTP error sending email: {smtp_err}")
                email_sent = False
        
        if email_sent:
            return jsonify({
                'status': 'success', 
                'message': f'A temporary password has been sent to {email}. Please check your inbox and spam folder.'
            })
        else:
            return jsonify({
                'status': 'error',
                'message': 'Could not send email. Please contact your administrator for assistance.'
            })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ── SMTP Settings (Admin) ──
@app.route('/api/smtp_settings', methods=['GET'])
@require_admin
def get_smtp_settings():
    """Get SMTP configuration (password is hidden)."""
    try:
        return jsonify({
            'status': 'success',
            'smtp_enabled': db_manager.get_setting('smtp_enabled', 'False') == 'True',
            'smtp_server': db_manager.get_setting('smtp_server', ''),
            'smtp_port': db_manager.get_setting('smtp_port', '587'),
            'smtp_user': db_manager.get_setting('smtp_user', ''),
            'smtp_from': db_manager.get_setting('smtp_from', ''),
            'smtp_has_password': bool(db_manager.get_setting('smtp_password', ''))
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/smtp_settings', methods=['POST'])
@require_admin
def save_smtp_settings():
    """Save SMTP configuration."""
    try:
        data = request.json
        db_manager.save_setting('smtp_enabled', 'True' if data.get('smtp_enabled') else 'False')
        db_manager.save_setting('smtp_server', data.get('smtp_server', '').strip())
        db_manager.save_setting('smtp_port', str(data.get('smtp_port', '587')).strip())
        db_manager.save_setting('smtp_user', data.get('smtp_user', '').strip())
        db_manager.save_setting('smtp_from', data.get('smtp_from', '').strip())
        
        smtp_pass = data.get('smtp_password', '').strip()
        if smtp_pass:
            db_manager.save_encrypted_setting('smtp_password', smtp_pass)
        
        return jsonify({'status': 'success', 'message': 'SMTP settings saved!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ── Test SMTP ──
@app.route('/api/test_smtp', methods=['POST'])
@require_admin
def test_smtp():
    """Test SMTP connection by sending a test email to the current admin."""
    try:
        data = request.json
        test_email = data.get('test_email', '').strip()
        server = data.get('smtp_server', '').strip()
        port = str(data.get('smtp_port', '587')).strip()
        user = data.get('smtp_user', '').strip()
        password = data.get('smtp_password', '').strip()
        from_addr = data.get('smtp_from', '').strip()
        
        if not test_email:
            return jsonify({'status': 'error', 'message': 'Test email address is required.'}), 400
        
        import smtplib
        from email.message import EmailMessage
        
        msg = EmailMessage()
        msg.set_content('This is a test email from your Document Tracking System.\n\nSMTP configuration is working correctly!')
        msg['Subject'] = 'Document Tracking System - SMTP Test'
        msg['From'] = from_addr or user
        msg['To'] = test_email
        
        server_obj = smtplib.SMTP(server, int(port))
        server_obj.starttls()
        server_obj.login(user, password)
        server_obj.send_message(msg)
        server_obj.quit()
        
        return jsonify({'status': 'success', 'message': f'Test email sent to {test_email}!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'SMTP test failed: {str(e)}'})

# ─────────────────────────────────────────────
#  LOCAL <-> CLOUD SYNC ENDPOINTS
# ─────────────────────────────────────────────

@app.route('/api/sync/export', methods=['GET'])
@require_admin
def sync_export():
    try:
        conn = db_manager.get_db_connection()
        data = {}
        for table in ['routing_records', 'code_lookup', 'users', 'settings']:
            rows = conn.execute(f'SELECT * FROM {table}').fetchall()
            data[table] = [dict(r) for r in rows]
        conn.close()

        import time
        timestamp = time.strftime('%Y%m%d_%H%M%S')
        backup_file = f'static/sync_backup_{timestamp}.json'
        with open(backup_file, 'w', encoding='utf-8') as f:
            json_lib.dump(data, f, ensure_ascii=False, indent=2)

        return send_file(backup_file, mimetype="application/json", as_attachment=True,
                         download_name=f'db_backup_{timestamp}.json')
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/sync/import', methods=['POST'])
def sync_import():
    try:
        # Allow admin session OR valid sync token
        token = request.headers.get('X-Sync-Token', '')
        valid_token = db_manager.get_setting('sync_token', '')
        is_admin = session.get('user_role') == 'Admin'
        
        if not is_admin and token != valid_token:
            # If no token is set yet but one was provided, auto-set it (first push)
            if not valid_token and token:
                db_manager.save_setting('sync_token', token)
                valid_token = token
            else:
                return jsonify({'status': 'error', 'message': 'Unauthorized. Provide valid admin session or sync token.'}), 403
        
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': 'No file uploaded.'}), 400

        file = request.files['file']
        if not file.filename.endswith('.json'):
            return jsonify({'status': 'error', 'message': 'Only .json files accepted.'}), 400

        data = json_lib.loads(file.read().decode('utf-8'))
        conn = db_manager.get_db_connection()
        cursor = conn.cursor()
        report = {}

        for table in ['routing_records', 'code_lookup', 'settings']:
            cursor.execute(f'DELETE FROM {table}')
            count = 0
            for row in data.get(table, []):
                cols = list(row.keys())
                ph = ['?'] * len(cols)
                cursor.execute(f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({', '.join(ph)})", [row[c] for c in cols])
                count += 1
            report[table] = count

        admin_row = cursor.execute('SELECT password_hash FROM users WHERE username = ?', ('admin',)).fetchone()
        cursor.execute('DELETE FROM users')
        count = 0
        for row in data.get("users", []):
            cols = list(row.keys())
            ph = ['?'] * len(cols)
            cursor.execute(f"INSERT INTO users ({', '.join(cols)}) VALUES ({', '.join(ph)})", [row[c] for c in cols])
            count += 1
        if admin_row and not cursor.execute("SELECT username FROM users WHERE username = ?", ("admin",)).fetchone():
            cursor.execute('INSERT INTO users (username, password_hash, role, status, school_office, email, requires_password_change) VALUES (?,?,?,?,?,?,?)',
                           ('admin', admin_row[0], 'Admin', 'Approved', '', 'admin@deped.gov.ph', 0))
            count += 1
        report["users"] = count

        conn.commit()
        conn.close()
        return jsonify({"status": "success", "message": f"Imported: {report['routing_records']} records, {report['code_lookup']} codes, {report['users']} users", "report": report})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/sync/push_to_cloud', methods=['POST'])
@require_admin
def sync_push_to_cloud():
    try:
        import requests as http_requests, tempfile, time, os

        data = request.json
        cloud_url = data.get('cloud_url', '').strip().rstrip('/')
        if not cloud_url:
            return jsonify({'status': 'error', 'message': 'Cloud URL is required.'}), 400

        conn = db_manager.get_db_connection()
        db_data = {}
        for table in ['routing_records', 'code_lookup', 'users', 'settings']:
            rows = conn.execute(f'SELECT * FROM {table}').fetchall()
            db_data[table] = [dict(r) for r in rows]
        conn.close()

        timestamp = time.strftime('%Y%m%d_%H%M%S')
        tmp_path = os.path.join(tempfile.gettempdir(), f'db_push_{timestamp}.json')
        with open(tmp_path, 'w', encoding='utf-8') as f:
            json_lib.dump(db_data, f, ensure_ascii=False, indent=2)

        sync_token = db_manager.get_setting('sync_token', '')
        headers = {}
        if sync_token:
            headers['X-Sync-Token'] = sync_token
        with open(tmp_path, 'rb') as f:
            resp = http_requests.post(f'{cloud_url}/api/sync/import',
                                      files={'file': (f'db_{timestamp}.json', f, 'application/json')},
                                      headers=headers, timeout=120)
        os.remove(tmp_path)

        if resp.status_code == 200:
            return jsonify({'status': 'success', 'message': f"Cloud synced! {resp.json().get('message', '')}" })
        else:
            return jsonify({'status': 'error', 'message': f"Cloud returned {resp.status_code}: {resp.text[:300]}" })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


#  STARTUP
# ─────────────────────────────────────────────
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)