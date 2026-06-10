import os
import io
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

import db_manager

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'barcode-routing-secret-2024-change-in-prod')
CORS(app)

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
    """Return (is_enabled, sheet_id, creds) tuple. Creds are decrypted."""
    enabled = db_manager.get_setting('gsheets_enabled') == 'True'
    sheet_id = db_manager.get_setting('gsheets_id')
    creds = db_manager.get_decrypted_setting('gsheets_credentials')
    return enabled and bool(sheet_id) and bool(creds), sheet_id, creds


# Master data is now stored in the SQLite `master_data` table.
# See db_manager.py for helper functions: get_all_departments(),
# get_schools_for_department(), get_employees_for_school(), etc.


def save_code_lookup(code, department, school, employees, doc_type):
    gen_time = datetime.now().strftime("%m/%d/%Y")
    
    # Sanitize employee names: if any are still in "LastName, FirstName" format, reorder them
    employees = [db_manager.reorder_name(e) for e in employees]
    
    # Write to Google Sheets FIRST (synchronous, authoritative source)
    ok, sheet_id, creds = _get_gsheets_config()
    if ok:
        code_dict = {
            'code': code, 'department': department, 'school_office': school,
            'employees': '|||'.join(employees), 'doc_type': doc_type, 'generated_at': gen_time
        }
        # Synchronous push — fails fast if GSheets is unreachable
        db_manager._gsheets_push_with_retry(
            db_manager.push_code_to_gsheets, sheet_id, creds, code_dict
        )
    
    # Then update local SQLite cache for fast reads
    conn = db_manager.get_db_connection()
    conn.execute(
        "INSERT OR REPLACE INTO code_lookup (code, department, school_office, employees, doc_type, generated_at) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (code, department, school, '|||'.join(employees), doc_type, gen_time)
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
    
    # Handle both old (comma) and new (|||) delimiters for backward compatibility
    employees_str = row['employees']
    if '|||' in employees_str:
        employees_list = employees_str.split('|||')
    else:
        # Old format - split by comma (will be migrated)
        employees_list = employees_str.split(',')
    
    # Tolerant: reorder any names still in "LastName, FirstName" format to "FirstName LastName"
    employees_list = [db_manager.reorder_name(e) for e in employees_list]
    
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
    return render_template('scanner.html')


# ─────────────────────────────────────────────
#  AUTH ENDPOINTS
# ─────────────────────────────────────────────

@app.route('/api/login', methods=['POST'])
def login():
    try:
        data = request.json
        username = data.get('username', '').strip().lower()
        password = data.get('password', '').strip()
        
        if not username or not password:
            return jsonify({'status': 'error', 'message': 'Username and passkey are required.'}), 400
            
        conn = db_manager.get_db_connection()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        conn.close()
        
        if not user:
            return jsonify({'status': 'error', 'message': 'Invalid username or passkey.'}), 401
            
        pw_hash = hashlib.sha256(password.encode()).hexdigest()
        if user['password_hash'] != pw_hash:
            return jsonify({'status': 'error', 'message': 'Invalid username or passkey.'}), 401
            
        if user['status'] != 'Approved':
            return jsonify({'status': 'error', 'message': 'Your account is pending admin approval.'}), 403
            
        supervised = (user['supervised_schools'] or '').strip()
        session['username'] = user['username']
        session['user_role'] = user['role']
        session['school_office'] = user['school_office'] or ''
        session['supervised_schools'] = supervised
        session['email'] = user['email'] or ''
        session.permanent = True
        
        return jsonify({
            'status': 'success',
            'message': f'Logged in successfully as {user["role"]}!',
            'requires_password_change': bool(user['requires_password_change']),
            'user': {
                'username': user['username'],
                'email': user['email'] or '',
                'role': user['role'],
                'school_office': user['school_office'] or '',
                'supervised_schools': supervised
            }
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/register', methods=['POST'])
def register():
    try:
        data = request.json
        username = data.get('username', '').strip().lower()
        role_request = data.get('role', 'User').strip()
        school_office = data.get('school_office', '').strip()
        email = data.get('email', '').strip()
        
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
        
        # Default password is the username itself, and force change on first login
        password = username  # default passkey = username
        pw_hash = hashlib.sha256(password.encode()).hexdigest()
        supervised_schools = ''
        if role_request == 'Supervisor':
            supervised_schools = data.get('supervised_schools', '').strip()
        conn.execute(
            "INSERT INTO users (username, password_hash, role, status, school_office, email, requires_password_change, supervised_schools) "
            "VALUES (?, ?, ?, 'Pending', ?, ?, 1, ?)",
            (username, pw_hash, role_request, school_office, email, supervised_schools)
        )
        conn.commit()
        conn.close()
        
        # Write to Google Sheets FIRST (synchronous, authoritative source)
        ok, sheet_id, creds = _get_gsheets_config()
        if ok:
            user_dict = {'username': username, 'password_hash': pw_hash, 'role': role_request, 'status': 'Pending', 'school_office': school_office, 'email': email, 'requires_password_change': 1}
            db_manager._gsheets_push_with_retry(
                db_manager.push_user_to_gsheets, sheet_id, creds, user_dict
            )
            
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
        'supervised_schools': session.get('supervised_schools', '')
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
        supervised_schools = data.get('supervised_schools', '').strip()
        username = session['username']
        
        conn = db_manager.get_db_connection()
        user_role = conn.execute("SELECT role FROM users WHERE username = ?", (username,)).fetchone()['role']
        if user_role == 'Supervisor':
            conn.execute(
                "UPDATE users SET school_office = ?, supervised_schools = ? WHERE username = ?",
                (school_office, supervised_schools, username)
            )
        else:
            conn.execute("UPDATE users SET school_office = ? WHERE username = ?", (school_office, username))
        conn.commit()
        conn.close()
        
        session['school_office'] = school_office
        if user_role == 'Supervisor':
            session['supervised_schools'] = supervised_schools
        
        return jsonify({'status': 'success', 'message': 'Profile updated!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/get_user_profile', methods=['GET'])
@require_login
def get_user_profile():
    try:
        conn = db_manager.get_db_connection()
        user = conn.execute(
            "SELECT username, email, role, status, school_office, supervised_schools FROM users WHERE username = ?",
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
        rows = conn.execute("SELECT username, role, status, school_office, supervised_schools FROM users").fetchall()
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
            
        supervised_schools = data.get('supervised_schools', '').strip()
        conn.execute(
            "UPDATE users SET status = 'Approved', role = ?, supervised_schools = ? WHERE username = ?",
            (assigned_role, supervised_schools, username)
        )
        conn.commit()
        
        # Get updated user dict for GSheets sync
        updated_user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        conn.close()
        
        # Write to Google Sheets FIRST (synchronous, authoritative source)
        ok, sheet_id, creds = _get_gsheets_config()
        if ok:
            db_manager._gsheets_push_with_retry(
                db_manager.push_user_to_gsheets, sheet_id, creds, dict(updated_user)
            )
            
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
        
        # Push rejection status to Google Sheets synchronously
        ok, sheet_id, creds = _get_gsheets_config()
        if ok:
            user_dict = {'username': username, 'password_hash': 'DELETED', 'role': 'None', 'status': 'Rejected', 'school_office': '', 'email': ''}
            db_manager._gsheets_push_with_retry(
                db_manager.push_user_to_gsheets, sheet_id, creds, user_dict
            )
            
        return jsonify({'status': 'success', 'message': f'Account request "{username}" rejected.'})
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
            
        if len(new_password) < 6:
            return jsonify({'status': 'error', 'message': 'New password must be at least 6 characters.'}), 400
            
        conn = db_manager.get_db_connection()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        
        if not user:
            conn.close()
            return jsonify({'status': 'error', 'message': 'User not found.'}), 404
            
        cur_hash = hashlib.sha256(current_password.encode()).hexdigest()
        if user['password_hash'] != cur_hash:
            conn.close()
            return jsonify({'status': 'error', 'message': 'Current password is incorrect.'}), 400
            
        new_hash = hashlib.sha256(new_password.encode()).hexdigest()
        conn.execute("UPDATE users SET password_hash = ?, requires_password_change = 0 WHERE username = ?", (new_hash, username))
        conn.commit()
        conn.close()
        
        return jsonify({'status': 'success', 'message': 'Password changed successfully!'})
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
            
        conn = db_manager.get_db_connection()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (target_username,)).fetchone()
        if not user:
            conn.close()
            return jsonify({'status': 'error', 'message': 'User not found.'}), 404
            
        new_hash = hashlib.sha256(new_password.encode()).hexdigest()
        conn.execute("UPDATE users SET password_hash = ? WHERE username = ?", (new_hash, target_username))
        conn.commit()
        
        # Get updated user dict for GSheets sync
        updated_user = conn.execute("SELECT * FROM users WHERE username = ?", (target_username,)).fetchone()
        conn.close()
        
        # Write to Google Sheets synchronously (authoritative source)
        ok, sheet_id, creds = _get_gsheets_config()
        if ok:
            db_manager._gsheets_push_with_retry(
                db_manager.push_user_to_gsheets, sheet_id, creds, dict(updated_user)
            )
            
        return jsonify({'status': 'success', 'message': f'Passkey for "{target_username}" reset successfully!'})
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
        conn.close()

        codes = []
        for r in rows:
            d = dict(r)
            # Parse employees back to list for the frontend
            emp_str = d.get('employees', '')
            if '|||' in emp_str:
                d['employees_list'] = emp_str.split('|||')
            else:
                d['employees_list'] = emp_str.split(',')
            codes.append(d)

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
        supervised    = session.get('supervised_schools', '').strip()

        if user_role == 'User' and user_school:
            # Force to assigned school
            conn = db_manager.get_db_connection()
            row = conn.execute(
                "SELECT department FROM master_data WHERE school_office = ? LIMIT 1",
                (user_school,)
            ).fetchone()
            conn.close()
            if row:
                department = row['department']
            school = user_school

        elif user_role == 'Supervisor' and supervised:
            schools = [s.strip() for s in supervised.split(',') if s.strip()]
            if school and school not in schools:
                return jsonify({'status': 'error', 'message': 'You can only generate codes for your supervised schools.'}), 403

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
    with scan_lock:
        try:
            data             = request.json
            code             = data.get('scanned_data', '').strip().upper()
            receiving_office = data.get('receiving_office', '').strip()
            receiver_name    = data.get('receiver_name', '').strip()

            # ── ROLE-BASED SCHOOL RESTRICTIONS ──
            user_role     = session.get('user_role', '')
            user_school   = session.get('school_office', '')
            supervised    = session.get('supervised_schools', '').strip()
            if user_role == 'User' and user_school:
                receiving_office = user_school
            elif user_role == 'Supervisor' and supervised:
                allowed = [s.strip().lower() for s in supervised.split(',') if s.strip()]
                if receiving_office.strip().lower() not in allowed:
                    return jsonify({'status': 'error', 'message': 'You can only scan at your supervised schools/offices.'}), 403

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
            current_ts       = datetime.now().strftime("%m/%d/%Y")
            updated_rows_ids = []
            slot_updated     = 1

            for emp in employees:
                row = conn.execute(
                    "SELECT * FROM routing_records WHERE code = ? AND employee = ?", (code, emp)
                ).fetchone()

                if not row:
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
        supervised = session.get('supervised_schools', '').strip()

        if user_role == 'User' and user_school:
            where_sql = "WHERE school_office = ?"
            params.append(user_school)
            if search:
                where_sql += " AND (employee LIKE ? OR code LIKE ? OR department LIKE ? OR school_office LIKE ?)"
                pat = f"%{search}%"
                params += [pat, pat, pat, pat]

        elif user_role == 'Supervisor' and supervised:
            schools_list = [s.strip() for s in supervised.split(',') if s.strip()]
            placeholders = ','.join(['?'] * len(schools_list))
            where_sql = f"WHERE school_office IN ({placeholders})"
            params.extend(schools_list)
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

        # Push updated record to Google Sheets synchronously
        ok, sheet_id, creds = _get_gsheets_config()
        if ok:
            push_conn = db_manager.get_db_connection()
            updated = push_conn.execute("SELECT * FROM routing_records WHERE id = ?", (record_id,)).fetchone()
            if updated:
                db_manager._gsheets_push_with_retry(
                    db_manager.push_row_to_gsheets, sheet_id, creds, dict(updated)
                )
            push_conn.close()
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
        # Never expose the actual credentials to the frontend
        has_creds = bool(db_manager.get_setting('gsheets_credentials', ''))
        # Show whether credentials come from environment variables (persistent) or SQLite (wiped on Render)
        creds_from_env = bool(os.environ.get('GSHEETS_CREDENTIALS'))
        id_from_env = bool(os.environ.get('GSHEETS_ID'))
        return jsonify({
            'status':               'success',
            'gsheets_enabled':      db_manager.get_setting('gsheets_enabled', 'False') == 'True',
            'gsheets_id':           db_manager.get_setting('gsheets_id', ''),
            'gsheets_configured':   has_creds,
            'gsheets_from_env':     creds_from_env and id_from_env,  # True = survives Render wipes
            'scanner_pin':          db_manager.get_setting('scanner_pin', 'scanner123'),
            'admin_pin':            db_manager.get_setting('admin_pin', 'admin123'),
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
@require_admin
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

        if data.get('gsheets_pull_interval', '').strip():
            val = data['gsheets_pull_interval'].strip()
            try:
                parsed = int(val)
                if parsed >= 1:
                    db_manager.save_setting('gsheets_pull_interval', str(parsed))
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

@app.route('/api/gsheet_status', methods=['GET'])
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
def gsheets_last_pull():
    return jsonify(db_manager.get_last_pull_info())


@app.route('/api/bulk_sync', methods=['POST'])
@require_admin
def bulk_sync():
    """Push ALL existing data (routing records, codes, users, master_data) to Google Sheets at once."""
    try:
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
        for r in all_ids:
            db_manager.queue_sync(r['id'])
        return jsonify({'status': 'success', 'message': msg, 'rewritten': rewritten, 'total': total, 'samples': samples})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/admin/migrate_master_to_gsheets', methods=['POST'])
@require_admin
def migrate_master_to_gsheets():
    """Push all 3 raw sheets from master_db.xlsx to Google Sheets as separate worksheets
    ('Master Data Sheet1', 'Master Data Sheet2', 'Master Data Sheet3')."""
    try:
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
        doc_join = " LEFT JOIN code_lookup cl_ ON rr.code = cl_.code"
        doc_where = ""
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
        total_q = f"SELECT COUNT(*) FROM routing_records rr{doc_join} WHERE 1=1{doc_where}"
        total_employees = conn.execute(total_q, params).fetchone()[0]

        # Active codes
        active_q = f"SELECT COUNT(DISTINCT rr.code) FROM routing_records rr{doc_join} WHERE 1=1{doc_where}"
        active_codes = conn.execute(active_q, params).fetchone()[0]

        # Released count
        released_q = f"SELECT {count_expr('cl_.doc_type')} FROM routing_records rr{doc_join} WHERE LOWER(rr.status) = 'released'{doc_where}"
        released_count = conn.execute(released_q, params).fetchone()[0]

        # For signature count
        sig_q = f"SELECT {count_expr('cl_.doc_type')} FROM routing_records rr{doc_join} WHERE LOWER(rr.status) = 'for signature'{doc_where}"
        for_signature_count = conn.execute(sig_q, params).fetchone()[0]

        # With corrections count
        corr_q = f"SELECT {count_expr('cl_.doc_type')} FROM routing_records rr{doc_join} WHERE LOWER(rr.status) = 'with corrections'{doc_where}"
        with_corrections_count = conn.execute(corr_q, params).fetchone()[0]

        # Status breakdown
        status_breakdown = []
        if is_schools:
            status_q = f"""
                SELECT LOWER(rr.status) as status,
                       COUNT(DISTINCT rr.school_office || '|' || COALESCE(cl_.doc_type, '')) as cnt
                FROM routing_records rr{doc_join}
                WHERE 1=1{doc_where}
                GROUP BY LOWER(rr.status)
            """
        else:
            status_q = f"""
                SELECT LOWER(rr.status) as status, COUNT(*) as cnt
                FROM routing_records rr{doc_join}
                WHERE 1=1{doc_where}
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
                WHERE 1=1{doc_where}
                GROUP BY cl_.doc_type
                ORDER BY cnt DESC LIMIT 10
            """
        else:
            dt_q = f"""
                SELECT COALESCE(cl_.doc_type, '') as doc_type, COUNT(*) as cnt
                FROM routing_records rr{doc_join}
                WHERE 1=1{doc_where}
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
#  RECENT ACTIVITY
# ─────────────────────────────────────────────
@app.route('/api/recent_activity', methods=['GET'])
def recent_activity():
    """Return the 10 most recently scanned codes with employee count and latest office.
    Query params: doc_type — filter by document type, count_mode — 'rows' or 'schools'."""
    try:
        doc_type_filter = request.args.get('doc_type', '').strip()
        count_mode = request.args.get('count_mode', 'rows').strip()
        is_schools = count_mode == 'schools'
        conn = db_manager.get_db_connection()
        subqueries = []
        for i in range(1, 11):
            subqueries.append(
                f"SELECT rr.code, rr.employee, rr.school_office AS origin_school, "
                f"rr.receiving_office_{i} AS office, "
                f"rr.timestamp_{i} AS ts "
                f"FROM routing_records rr "
                f"WHERE rr.receiving_office_{i} IS NOT NULL AND rr.receiving_office_{i} != ''"
            )
        union_sql = " UNION ALL ".join(subqueries)
        params = []
        doc_where = ""
        if doc_type_filter:
            doc_where = " AND cl.doc_type = ?"
            params.append(doc_type_filter)
        if is_schools:
            group_sql = (
                f"SELECT scans.origin_school, "
                f"cl.doc_type, "
                f"COUNT(DISTINCT scans.code) AS total_codes, "
                f"MIN(scans.employee) AS first_employee, "
                f"COUNT(DISTINCT scans.employee) AS total_employees, "
                f"MAX(scans.office) AS office, "
                f"MAX(scans.ts) AS last_timestamp "
                f"FROM ({union_sql}) AS scans "
                f"JOIN code_lookup cl ON scans.code = cl.code "
                f"WHERE 1=1{doc_where} "
                f"GROUP BY scans.origin_school, cl.doc_type "
                f"ORDER BY last_timestamp DESC LIMIT 10"
            )
        else:
            group_sql = (
                f"SELECT scans.code, "
                f"MIN(scans.employee) AS first_employee, "
                f"COUNT(DISTINCT scans.employee) AS total_employees, "
                f"MIN(scans.origin_school) AS origin_school, "
                f"MAX(scans.office) AS office, "
                f"MAX(scans.ts) AS last_timestamp "
                f"FROM ({union_sql}) AS scans "
                f"JOIN code_lookup cl ON scans.code = cl.code "
                f"WHERE 1=1{doc_where} "
                f"GROUP BY scans.code "
                f"ORDER BY last_timestamp DESC LIMIT 10"
            )
        rows = conn.execute(group_sql, params).fetchall()
        activities = []
        for row in rows:
            if is_schools:
                school = row['origin_school'] or ''
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
def forgot_password():
    """
    Forgot password: if user exists with an email, generate a temporary password,
    save its hash with requires_password_change=1, and attempt to email it.
    Since passwords are stored as SHA256 hashes, we cannot retrieve the original.
    Instead we generate a new temporary password.
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
        
        # Generate a random temporary password
        import secrets
        temp_password = secrets.token_urlsafe(8)  # e.g. "abc123_xY"
        temp_hash = hashlib.sha256(temp_password.encode()).hexdigest()
        
        conn.execute("UPDATE users SET password_hash = ?, requires_password_change = 1 WHERE username = ?",
                     (temp_hash, username))
        conn.commit()
        conn.close()
        
        # Try to send email using SMTP settings
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
            # Email not sent but password was reset - show temp password directly (fallback)
            return jsonify({
                'status': 'warning',
                'message': f'Could not send email. Your temporary password is: {temp_password}. Please change it after logging in.',
                'temp_password': temp_password  # only shown in fallback mode
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