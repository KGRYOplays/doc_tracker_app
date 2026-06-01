import os
import io
import pandas as pd
import qrcode
import random
import string
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, jsonify, session, send_file, redirect, url_for
from PIL import Image, ImageDraw, ImageFont
import hashlib

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
    gen_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = db_manager.get_db_connection()
    conn.execute(
        "INSERT OR REPLACE INTO code_lookup (code, department, school_office, employees, doc_type, generated_at) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (code, department, school, '|||'.join(employees), doc_type, gen_time)
    )
    conn.commit()
    conn.close()
    
    # Push to GSheets
    enabled = db_manager.get_setting('gsheets_enabled') == 'True'
    sheet_id = db_manager.get_setting('gsheets_id')
    creds = db_manager.get_decrypted_setting('gsheets_credentials')
    if enabled and sheet_id and creds:
        import threading
        code_dict = {'code': code, 'department': department, 'school_office': school, 'employees': '|||'.join(employees), 'doc_type': doc_type, 'generated_at': gen_time}
        threading.Thread(target=db_manager.push_code_to_gsheets, args=(sheet_id, creds, code_dict), daemon=True).start()
        
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
    """Generate a barcode PIL image with 1.5-inch width in a bordered box."""
    from barcode import Code128
    from barcode.writer import ImageWriter

    # 1.5 inches = 144 pixels at 96 DPI (but we'll use 300 DPI for quality)
    # Target: 1.5 inches ≈ 450 pixels at 300 DPI
    DPI = 300
    TARGET_W_INCHES = 1.5

    # First pass: generate base barcode with custom writer options
    writer_options = {
        'module_width': 0.35,      # wider modules for readability
        'module_height': 15.0,     # taller barcode
        'font_size': 10,
        'text_distance': 3,
        'quiet_zone': 1.0,
        'write_text': True,
        'background': 'white',
        'foreground': 'black',
    }
    
    writer = ImageWriter()
    barcode_obj = Code128(code, writer=writer)
    
    # Render to bytes first
    buf = io.BytesIO()
    barcode_obj.write(buf, writer_options)
    buf.seek(0)
    img = Image.open(buf).convert('RGB')
    
    # Resize to exactly 1.5 inches width at 300 DPI
    target_px = int(TARGET_W_INCHES * DPI)
    aspect = img.height / img.width
    target_h = int(target_px * aspect)
    img = img.resize((target_px, target_h), Image.LANCZOS)
    
    # Add border (box) around the barcode
    border_px = 8
    boxed_w = target_px + border_px * 2
    boxed_h = target_h + border_px * 2 + 30  # extra space for code text below
    boxed = Image.new('RGB', (boxed_w, boxed_h), 'white')
    draw = ImageDraw.Draw(boxed)
    
    # Draw outer border box
    draw.rectangle([0, 0, boxed_w - 1, boxed_h - 1], outline='black', width=2)
    
    # Paste barcode image centered in box
    paste_x = (boxed_w - target_px) // 2
    paste_y = border_px
    boxed.paste(img, (paste_x, paste_y))
    
    # Add code text below the barcode
    try:
        font = ImageFont.truetype("arial.ttf", 18)
    except:
        font = ImageFont.load_default()
    
    text = f"CODE: {code}"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    text_x = (boxed_w - text_w) // 2
    text_y = boxed_h - 25
    draw.text((text_x, text_y), text, fill='black', font=font)
    
    return boxed


def generate_qr_image(code):
    """Generate a QR code PIL image with 1.5-inch width in a bordered box."""
    qr = qrcode.QRCode(version=1, box_size=10, border=2)
    qr.add_data(code)
    qr.make(fit=True)
    img = qr.make_image(fill="black", back_color="white").convert('RGB')
    
    # Resize to 1.5 inches at 300 DPI
    DPI = 300
    TARGET_W_INCHES = 1.5
    target_px = int(TARGET_W_INCHES * DPI)
    aspect = img.height / img.width
    target_h = int(target_px * aspect)
    img = img.resize((target_px, target_h), Image.LANCZOS)
    
    # Add border box
    border_px = 8
    boxed_w = target_px + border_px * 2
    boxed_h = target_h + border_px * 2 + 30
    boxed = Image.new('RGB', (boxed_w, boxed_h), 'white')
    draw = ImageDraw.Draw(boxed)
    
    draw.rectangle([0, 0, boxed_w - 1, boxed_h - 1], outline='black', width=2)
    paste_x = (boxed_w - target_px) // 2
    paste_y = border_px
    boxed.paste(img, (paste_x, paste_y))
    
    try:
        font = ImageFont.truetype("arial.ttf", 18)
    except:
        font = ImageFont.load_default()
    
    text = f"CODE: {code}"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    text_x = (boxed_w - text_w) // 2
    text_y = boxed_h - 25
    draw.text((text_x, text_y), text, fill='black', font=font)
    
    return boxed


def create_label_pdf(codes, doc_type='label', dpi=300, label_width_in=1.5):
    """
    Create a PDF with properly sized barcode labels (1.5 inches wide)
    suitable for batch printing on sticker/label sheets.
    """
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib.units import inch, mm
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
    
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    
    # Label dimensions
    label_w = label_width_in * inch
    label_h = label_w * 0.75  # 1.5 x 1.125 inch labels
    margin_x = 0.5 * inch
    margin_y = 0.5 * inch
    gap = 0.15 * inch
    
    # Calculate labels per page
    usable_w = letter[0] - 2 * margin_x
    usable_h = letter[1] - 2 * margin_y
    cols = int((usable_w + gap) / (label_w + gap))
    rows = int((usable_h + gap) / (label_h + gap))
    padding_x = (usable_w - (cols * label_w + (cols - 1) * gap)) / 2
    padding_y = (usable_h - (rows * label_h + (rows - 1) * gap)) / 2
    
    for idx, item in enumerate(codes):
        code = item.get('code', '')
        img_type = item.get('type', 'barcode')
        
        if idx > 0 and idx % (cols * rows) == 0:
            c.showPage()
        
        pos = idx % (cols * rows)
        col = pos % cols
        row = pos // cols
        
        x = margin_x + padding_x + col * (label_w + gap)
        y = margin_y + padding_y + row * (label_h + gap)
        
        # Generate image
        if img_type == 'qr':
            pil_img = generate_qr_image(code)
        else:
            pil_img = generate_barcode_image(code)
        
        # Scale to fit within label area with small padding
        img_w, img_h = pil_img.size
        scale = min(label_w * 0.85 / img_w, label_h * 0.85 / img_h)
        disp_w = img_w * scale
        disp_h = img_h * scale
        img_x = x + (label_w - disp_w) / 2
        img_y = y + (label_h - disp_h) / 2
        
        temp_buf = io.BytesIO()
        pil_img.save(temp_buf, format='PNG')
        temp_buf.seek(0)
        c.drawImage(ImageReader(temp_buf), img_x, img_y, width=disp_w, height=disp_h)
        
        # Draw label border outline (dashed)
        c.setStrokeColorRGB(0.7, 0.7, 0.7)
        c.setLineWidth(0.5)
        c.rect(x, y, label_w, label_h)
        c.setStrokeColorRGB(0, 0, 0)
        c.setLineWidth(1)
    
    c.save()
    buf.seek(0)
    return buf


def create_document_page_pdf(code, doc_type='barcode'):
    """
    Create a PDF with the barcode placed at the TOP RIGHT corner of a document.
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
    
    # ── Generate the barcode image ──
    if doc_type == 'qr':
        pil_img = generate_qr_image(code)
    else:
        pil_img = generate_barcode_image(code)
    
    # Convert PIL to ImageReader
    temp_buf = io.BytesIO()
    pil_img.save(temp_buf, format='PNG')
    temp_buf.seek(0)
    
    # ── Place barcode at TOP RIGHT corner ──
    # Barcode target size: 1.5 inches wide on the page
    target_w_inches = 1.5
    target_w = target_w_inches * inch  # 108 points
    aspect = pil_img.height / pil_img.width
    target_h = target_w * aspect
    
    # Position: top right corner with 0.5 inch margins
    margin = 0.5 * inch
    barcode_x = width - margin - target_w
    barcode_y = height - margin - target_h
    
    c.drawImage(ImageReader(temp_buf), barcode_x, barcode_y, width=target_w, height=target_h)
    
    # ── Document header / title ──
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin, height - 0.8 * inch, "DOCUMENT ROUTING SLIP")
    
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
                    office = row.get(f'receiving_office_{i}', '') or ''
                    receiver = row.get(f'receiver_name_{i}', '') or ''
                    ts = row.get(f'timestamp_{i}', '') or ''
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


def create_barcode_overlay_pdf(code, doc_type='barcode', position='top-right', page_size='legal'):
    """
    Create a minimal PDF with ONLY the barcode image at the specified position.
    No headers, no routing table - just the barcode on a blank page.
    Useful for printing the barcode onto an already-printed document.
    
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
    
    if doc_type == 'qr':
        pil_img = generate_qr_image(code)
    else:
        pil_img = generate_barcode_image(code)
    
    temp_buf = io.BytesIO()
    pil_img.save(temp_buf, format='PNG')
    temp_buf.seek(0)
    
    target_w_inches = 1.5
    target_w = target_w_inches * inch
    aspect = pil_img.height / pil_img.width
    target_h = target_w * aspect
    
    margin = 0.5 * inch
    
    if position == 'top-right':
        barcode_x = width - margin - target_w
        barcode_y = height - margin - target_h
    elif position == 'top-left':
        barcode_x = margin
        barcode_y = height - margin - target_h
    elif position == 'bottom-right':
        barcode_x = width - margin - target_w
        barcode_y = margin
    elif position == 'bottom-left':
        barcode_x = margin
        barcode_y = margin
    else:
        barcode_x = width - margin - target_w
        barcode_y = height - margin - target_h
    
    c.drawImage(ImageReader(temp_buf), barcode_x, barcode_y, width=target_w, height=target_h)
    
    c.setFont("Helvetica", 8)
    code_label = f"CODE: {code}"
    code_label_w = c.stringWidth(code_label, "Helvetica", 8)
    label_x = barcode_x + (target_w - code_label_w) / 2
    label_y = barcode_y - 12
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
            
        session['username'] = user['username']
        session['user_role'] = user['role']
        session['school_office'] = user['school_office'] or ''
        session.permanent = True
        
        return jsonify({
            'status': 'success',
            'message': f'Logged in successfully as {user["role"]}!',
            'user': {
                'username': user['username'],
                'role': user['role'],
                'school_office': user['school_office'] or ''
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
        conn.execute("INSERT INTO users (username, password_hash, role, status, school_office, email, requires_password_change) VALUES (?, ?, ?, 'Pending', ?, ?, 1)",
                     (username, pw_hash, role_request, school_office, email))
        conn.commit()
        conn.close()
        
        # Sync users to sheets if enabled
        enabled = db_manager.get_setting('gsheets_enabled') == 'True'
        sheet_id = db_manager.get_setting('gsheets_id')
        creds = db_manager.get_decrypted_setting('gsheets_credentials')
        if enabled and sheet_id and creds:
            user_dict = {'username': username, 'password_hash': pw_hash, 'role': role_request, 'status': 'Pending', 'school_office': school_office, 'email': email}
            import threading
            threading.Thread(target=db_manager.push_user_to_gsheets, args=(sheet_id, creds, user_dict), daemon=True).start()
            
        return jsonify({'status': 'success', 'message': 'Registration request submitted! Please wait for Admin approval.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ── Guest login (no credentials needed, read-only records view) ──
@app.route('/api/guest_login', methods=['POST'])
def guest_login():
    try:
        session['username'] = 'guest'
        session['user_role'] = 'Guest'
        session['school_office'] = ''
        session.permanent = True
        return jsonify({
            'status': 'success',
            'message': 'Logged in as Guest!',
            'user': {
                'username': 'guest',
                'role': 'Guest',
                'school_office': ''
            }
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/check_auth', methods=['GET'])
def check_auth():
    return jsonify({
        'logged_in': bool(session.get('username')),
        'username': session.get('username', ''),
        'role': session.get('user_role', 'Viewer'),
        'school_office': session.get('school_office', '')
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
        user = conn.execute("SELECT username, role, status, school_office FROM users WHERE username = ?", (session['username'],)).fetchone()
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
        rows = conn.execute("SELECT username, role, status, school_office FROM users").fetchall()
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
            
        conn.execute("UPDATE users SET status = 'Approved', role = ? WHERE username = ?", (assigned_role, username))
        conn.commit()
        
        # Get updated user dict for GSheets sync
        updated_user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        conn.close()
        
        enabled = db_manager.get_setting('gsheets_enabled') == 'True'
        sheet_id = db_manager.get_setting('gsheets_id')
        creds = db_manager.get_decrypted_setting('gsheets_credentials')
        if enabled and sheet_id and creds:
            import threading
            threading.Thread(target=db_manager.push_user_to_gsheets, args=(sheet_id, creds, dict(updated_user)), daemon=True).start()
            
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
        
        # Note: In sheets we delete by updating status to 'Rejected' or simply updating sheet (will re-sync on bulk)
        # We can push to sheets with 'Rejected' status to notify cloud
        enabled = db_manager.get_setting('gsheets_enabled') == 'True'
        sheet_id = db_manager.get_setting('gsheets_id')
        creds = db_manager.get_decrypted_setting('gsheets_credentials')
        if enabled and sheet_id and creds:
            import threading
            user_dict = {'username': username, 'password_hash': 'DELETED', 'role': 'None', 'status': 'Rejected', 'school_office': ''}
            threading.Thread(target=db_manager.push_user_to_gsheets, args=(sheet_id, creds, user_dict), daemon=True).start()
            
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
        
        enabled = db_manager.get_setting('gsheets_enabled') == 'True'
        sheet_id = db_manager.get_setting('gsheets_id')
        creds = db_manager.get_decrypted_setting('gsheets_credentials')
        if enabled and sheet_id and creds:
            import threading
            threading.Thread(target=db_manager.push_user_to_gsheets, args=(sheet_id, creds, dict(updated_user)), daemon=True).start()
            
        return jsonify({'status': 'success', 'message': f'Passkey for "{target_username}" reset successfully!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


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
#  BARCODE LABEL / DOCUMENT PDF ENDPOINTS
# ─────────────────────────────────────────────

@app.route('/api/label_image/<code>')
def label_image(code):
    """
    Serve a PNG image of the barcode at 1.5-inch width with border box.
    Query param: type=barcode|qr (default: barcode)
    """
    try:
        img_type = request.args.get('type', 'barcode')
        if img_type == 'qr':
            img = generate_qr_image(code.upper())
        else:
            img = generate_barcode_image(code.upper())
        
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        buf.seek(0)
        return send_file(buf, mimetype='image/png')
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/document_page', methods=['POST'])
def document_page():
    """
    Generate a PDF with barcode at TOP RIGHT corner.
    Body: { "code": "...", "type": "barcode|qr" }
    """
    try:
        data = request.json
        code = data.get('code', '').strip().upper()
        img_type = data.get('type', 'barcode')
        
        if not code:
            return jsonify({'status': 'error', 'message': 'Code is required.'}), 400
        
        pdf_buf = create_document_page_pdf(code, img_type)
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
    Generate a PDF for batch printing multiple barcodes.
    Body: { "codes": [{"code": "...", "type": "barcode|qr"}, ...] }
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
    Generate a minimal overlay PDF with ONLY the barcode at the specified position.
    Body: { "code": "...", "type": "barcode|qr", "position": "top-right|top-left|bottom-right|bottom-left", "page_size": "legal|letter" }
    """
    try:
        data = request.json
        code = data.get('code', '').strip().upper()
        img_type = data.get('type', 'barcode')
        position = data.get('position', 'top-right')
        page_size = data.get('page_size', 'legal')
        
        if not code:
            return jsonify({'status': 'error', 'message': 'Code is required.'}), 400
        
        pdf_buf = create_barcode_overlay_pdf(code, img_type, position, page_size)
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
@require_role('Admin', 'Supervisor')
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
                latest_receiver = None
                for i in sorted(stages, reverse=True):
                    key = f'receiving_office_{i}'
                    if key in row.keys() and row[key]:
                        latest_office = row[key]
                        latest_receiver = row.get(f'receiver_name_{i}', '')
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

        db_manager.queue_sync(record_id)
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
        return jsonify({
            'status':               'success',
            'gsheets_enabled':      db_manager.get_setting('gsheets_enabled', 'False') == 'True',
            'gsheets_id':           db_manager.get_setting('gsheets_id', ''),
            'gsheets_configured':   has_creds,
            'scanner_pin':          db_manager.get_setting('scanner_pin', 'scanner123'),
            'admin_pin':            db_manager.get_setting('admin_pin', 'admin123'),
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/save_settings', methods=['POST'])
@require_admin
def save_settings():
    try:
        data = request.json
        db_manager.save_setting('gsheets_enabled', 'True' if data.get('gsheets_enabled') else 'False')
        db_manager.save_setting('gsheets_id', data.get('gsheets_id', '').strip())

        if data.get('gsheets_credentials', '').strip():
            # Use encrypted save - will validate JSON and encrypt
            db_manager.save_encrypted_setting('gsheets_credentials', data['gsheets_credentials'].strip())

        if data.get('scanner_pin', '').strip():
            db_manager.save_setting('scanner_pin', data['scanner_pin'].strip())
        if data.get('admin_pin', '').strip():
            db_manager.save_setting('admin_pin', data['admin_pin'].strip())

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


@app.route('/api/bulk_sync', methods=['POST'])
@require_admin
def bulk_sync():
    """Push ALL existing routing records to Google Sheets at once."""
    try:
        count, errors = db_manager.bulk_sync_to_gsheets()
        if errors and count == 0:
            return jsonify({
                'status': 'error',
                'message': f'Bulk sync failed: {errors[0][:200]}',
                'errors': errors[:10]
            })
        msg = f'Synced {count} records to Google Sheets.'
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


# ─────────────────────────────────────────────
#  STARTUP
# ─────────────────────────────────────────────
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)