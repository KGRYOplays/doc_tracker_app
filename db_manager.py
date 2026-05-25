import os
import json
import base64
import sqlite3
import threading
import queue
import time
from datetime import datetime
import pandas as pd
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

# Files
DB_FILE = 'app.db'
MASTER_DB = 'master_db.xlsx'
CODE_LOOKUP_EXCEL = 'code_lookup.xlsx'
ROUTING_RECORD_EXCEL = 'routing_record.xlsx'
OLD_LOG_EXCEL = 'scan_log.xlsx'

# Background Queue for Google Sheets Sync
gs_sync_queue = queue.Queue()

# ─────────────────────────────────────────────
#  CREDENTIAL ENCRYPTION HELPERS
# ─────────────────────────────────────────────

def _get_encryption_key():
    """Derive a Fernet encryption key from the app's secret key."""
    secret_key = os.environ.get('SECRET_KEY', 'barcode-routing-secret-2024-change-in-prod')
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=b'barcode_app_gsheets_salt_2024',
        iterations=480000,
    )
    key = base64.urlsafe_b64encode(kdf.derive(secret_key.encode()))
    return key


def encrypt_credentials(plaintext_json):
    """Encrypt Google service account credentials JSON for storage."""
    if not plaintext_json:
        return ''
    f = Fernet(_get_encryption_key())
    return f.encrypt(plaintext_json.encode()).decode()


def decrypt_credentials(encrypted_value):
    """Decrypt previously encrypted credentials. Returns empty string on failure."""
    if not encrypted_value:
        return ''
    try:
        f = Fernet(_get_encryption_key())
        return f.decrypt(encrypted_value.encode()).decode()
    except Exception:
        # If decryption fails, assume it's already plaintext (legacy data)
        return encrypted_value


def get_decrypted_setting(key, default=None):
    """Get a setting value, automatically decrypting 'gsheets_credentials'."""
    value = get_setting(key, default)
    if key == 'gsheets_credentials' and value:
        decrypted = decrypt_credentials(value)
        # Validate that it's actually JSON — if not, keep encrypted form
        try:
            json.loads(decrypted)
            return decrypted
        except (json.JSONDecodeError, TypeError):
            pass
    return value


def save_encrypted_setting(key, value):
    """Save a setting, automatically encrypting 'gsheets_credentials'."""
    if key == 'gsheets_credentials' and value:
        # Validate it's valid JSON before encrypting
        try:
            parsed = json.loads(value)
            required_fields = ['type', 'project_id', 'private_key', 'client_email']
            missing = [f for f in required_fields if f not in parsed]
            if missing:
                raise ValueError(f"Missing required fields in credentials JSON: {', '.join(missing)}")
        except json.JSONDecodeError:
            raise ValueError("Credentials must be valid JSON")
        encrypted = encrypt_credentials(value)
        save_setting(key, encrypted)
    else:
        save_setting(key, value)


def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Code Lookup Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS code_lookup (
            code TEXT PRIMARY KEY,
            department TEXT NOT NULL,
            school_office TEXT NOT NULL,
            employees TEXT NOT NULL,
            doc_type TEXT,
            generated_at TEXT NOT NULL
        )
    ''')
    
    # 2. Routing Records Table — 3 columns per stage: office, receiver_name, timestamp
    columns = [
        "id INTEGER PRIMARY KEY AUTOINCREMENT",
        "department TEXT NOT NULL",
        "school_office TEXT NOT NULL",
        "employee TEXT NOT NULL",
        "code TEXT NOT NULL",
    ]
    for i in range(1, 11):
        columns.append(f"receiving_office_{i} TEXT")
        columns.append(f"receiver_name_{i} TEXT")
        columns.append(f"timestamp_{i} TEXT")
    
    cursor.execute(f"CREATE TABLE IF NOT EXISTS routing_records ({', '.join(columns)})")
    cursor.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_code_employee ON routing_records(code, employee)")
    
    # 3. Settings Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    ''')
    
    conn.commit()
    
    # 4. Migrate: add receiver_name columns to existing databases
    _migrate_add_receiver_names(conn)
    
    conn.close()
    
    # Run data migrations (if old excel files exist and db is empty)
    migrate_from_excel()
    
    # Start Google Sheets background worker thread
    start_gsheets_worker()


def _migrate_add_receiver_names(conn):
    """Add receiver_name_N columns if they don't exist (for upgrading old databases)."""
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(routing_records)")
    existing_cols = [row[1] for row in cursor.fetchall()]
    
    # Find all stage numbers from receiving_office_N columns
    stages = set()
    for col in existing_cols:
        if col.startswith('receiving_office_'):
            try:
                stages.add(int(col.split('_')[-1]))
            except:
                pass
    
    changed = False
    for stage in sorted(stages):
        col = f"receiver_name_{stage}"
        if col not in existing_cols:
            try:
                cursor.execute(f"ALTER TABLE routing_records ADD COLUMN {col} TEXT")
                changed = True
                print(f"Migration: added {col} to routing_records")
            except Exception as e:
                print(f"Migration warning (non-fatal): {e}")
    
    if changed:
        conn.commit()


def ensure_routing_columns(conn, stage):
    """Ensure routing_records has all 3 columns for the given stage."""
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(routing_records)")
    existing_cols = [row[1] for row in cursor.fetchall()]
    
    col_office   = f"receiving_office_{stage}"
    col_receiver = f"receiver_name_{stage}"
    col_ts       = f"timestamp_{stage}"
    
    changed = False
    for col in [col_office, col_receiver, col_ts]:
        if col not in existing_cols:
            try:
                cursor.execute(f"ALTER TABLE routing_records ADD COLUMN {col} TEXT")
                changed = True
            except Exception as e:
                print(f"Error adding column {col}: {e}")
    
    if changed:
        conn.commit()
        print(f"Dynamically expanded SQLite routing schema → Stage {stage} columns added.")
    return changed


def get_setting(key, default=None):
    conn = get_db_connection()
    row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    conn.close()
    return row['value'] if row else default


def save_setting(key, value):
    conn = get_db_connection()
    conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, str(value)))
    conn.commit()
    conn.close()


# --- REAL-TIME EXCEL EXPORT ---
def trigger_excel_export():
    """Trigger a thread-safe export of SQLite data to local Excel files."""
    def _export():
        try:
            conn = sqlite3.connect(DB_FILE)
            
            # Export code lookup
            df_lookup = pd.read_sql_query(
                "SELECT code AS Code, department AS Department, school_office AS [School/Office], "
                "employees AS Employees, doc_type AS [Doc Type], generated_at AS Generated FROM code_lookup",
                conn
            )
            df_lookup.to_excel(CODE_LOOKUP_EXCEL, index=False)
            
            # Export routing records — dynamically determine stages
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(routing_records)")
            cols = [row[1] for row in cursor.fetchall()]
            
            stages = []
            for c in cols:
                if c.startswith("receiving_office_"):
                    try:
                        stages.append(int(c.split("_")[-1]))
                    except:
                        pass
            stages = sorted(list(set(stages)))
            if not stages:
                stages = list(range(1, 11))
            
            sql = ("SELECT department AS Department, school_office AS [School/Office], "
                   "employee AS Employee, code AS Code")
            for i in stages:
                sql += f", receiving_office_{i} AS [Receiving Office {i}]"
                # Only add receiver_name if column exists
                if f"receiver_name_{i}" in cols:
                    sql += f", receiver_name_{i} AS [Receiver Name {i}]"
                sql += f", timestamp_{i} AS [Timestamp {i}]"
            sql += " FROM routing_records"
            
            df_routing = pd.read_sql_query(sql, conn)
            df_routing.to_excel(ROUTING_RECORD_EXCEL, index=False)
            
            conn.close()
            print("Excel databases successfully updated.")
        except Exception as e:
            print(f"Error exporting to Excel: {e}")
    
    threading.Thread(target=_export, daemon=True).start()


# --- DATA MIGRATION FROM OLD EXCEL ---
def migrate_from_excel():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Migrate code lookup
    cursor.execute("SELECT COUNT(*) FROM code_lookup")
    if cursor.fetchone()[0] == 0 and os.path.exists(CODE_LOOKUP_EXCEL):
        try:
            df = pd.read_excel(CODE_LOOKUP_EXCEL)
            for _, row in df.iterrows():
                code     = str(row.get('Code', '')).strip()
                dept     = str(row.get('Department', 'Unknown')).strip()
                school   = str(row.get('School/Office', 'Unknown')).strip()
                employees = str(row.get('Employees', '')).strip()
                doc_type  = str(row.get('Doc Type', '')).strip()
                gen       = str(row.get('Generated', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))).strip()
                if code:
                    cursor.execute(
                        "INSERT OR IGNORE INTO code_lookup (code, department, school_office, employees, doc_type, generated_at) "
                        "VALUES (?, ?, ?, ?, ?, ?)",
                        (code, dept, school, employees, doc_type, gen)
                    )
            conn.commit()
            print("Migrated code lookup database successfully.")
        except Exception as e:
            print(f"Error migrating code lookup: {e}")
    
    # 2. Migrate routing records
    cursor.execute("SELECT COUNT(*) FROM routing_records")
    if cursor.fetchone()[0] == 0:
        if os.path.exists(ROUTING_RECORD_EXCEL):
            try:
                df = pd.read_excel(ROUTING_RECORD_EXCEL)
                for _, row in df.iterrows():
                    dept   = str(row.get('Department', '')).strip()
                    school = str(row.get('School/Office', '')).strip()
                    emp    = str(row.get('Employee', '')).strip()
                    code   = str(row.get('Code', '')).strip()
                    if not (emp and code):
                        continue
                    vals = [dept, school, emp, code]
                    placeholders = ["?", "?", "?", "?"]
                    for i in range(1, 11):
                        off  = row.get(f'Receiving Office {i}')
                        rn   = row.get(f'Receiver Name {i}')
                        ts   = row.get(f'Timestamp {i}')
                        vals.append(None if pd.isna(off) else str(off))
                        vals.append(None if pd.isna(rn)  else str(rn))
                        vals.append(None if pd.isna(ts)  else str(ts))
                        placeholders.extend(["?", "?", "?"])
                    sql = ("INSERT OR IGNORE INTO routing_records "
                           "(department, school_office, employee, code, " +
                           ", ".join([f"receiving_office_{i}, receiver_name_{i}, timestamp_{i}" for i in range(1, 11)]) +
                           f") VALUES ({', '.join(placeholders)})")
                    cursor.execute(sql, vals)
                conn.commit()
                print("Imported existing routing_record.xlsx successfully.")
            except Exception as e:
                print(f"Error importing existing routing record excel: {e}")

        elif os.path.exists(OLD_LOG_EXCEL):
            try:
                df = pd.read_excel(OLD_LOG_EXCEL)
                df = df.sort_values(by='Timestamp')
                for _, row in df.iterrows():
                    code  = str(row.get('Code', '')).strip()
                    emp   = str(row.get('Employee Name', '')).strip()
                    dept  = str(row.get('Department', 'Unknown')).strip()
                    school = str(row.get('School/Office', 'Unknown')).strip()
                    ts    = str(row.get('Timestamp', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))).strip()
                    if not (code and emp):
                        continue
                    cursor.execute("SELECT * FROM routing_records WHERE code = ? AND employee = ?", (code, emp))
                    exist = cursor.fetchone()
                    if not exist:
                        cursor.execute(
                            "INSERT INTO routing_records (department, school_office, employee, code, receiving_office_1, receiver_name_1, timestamp_1) "
                            "VALUES (?, ?, ?, ?, ?, ?, ?)",
                            (dept, school, emp, code, school, '', ts)
                        )
                    else:
                        record_id = exist['id']
                        for idx in range(1, 11):
                            if not exist[f'receiving_office_{idx}']:
                                cursor.execute(
                                    f"UPDATE routing_records SET receiving_office_{idx} = ?, timestamp_{idx} = ? WHERE id = ?",
                                    (school, ts, record_id)
                                )
                                break
                conn.commit()
                print("Migrated scan logs from scan_log.xlsx successfully.")
            except Exception as e:
                print(f"Error migrating scan logs: {e}")
    
    conn.close()
    trigger_excel_export()


# --- GOOGLE SHEETS INTEGRATION ---
def test_google_sheets_connection(sheet_id, service_account_json):
    """Test connecting to a Google Sheet and initialize headers if empty."""
    import gspread
    from google.oauth2.service_account import Credentials
    import json
    
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    try:
        creds_dict = json.loads(service_account_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        sh = client.open_by_key(sheet_id)
        
        try:
            ws = sh.worksheet("Routing Records")
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title="Routing Records", rows=1000, cols=50)
        
        # Headers include receiver_name
        headers = ["Department", "School/Office", "Employee", "Code"]
        for i in range(1, 11):
            headers += [f"Receiving Office {i}", f"Receiver Name {i}", f"Timestamp {i}"]
        
        existing_headers = ws.row_values(1)
        if not existing_headers:
            ws.insert_row(headers, 1)
            end_col = _col_to_letter(len(headers))
            ws.format(f"A1:{end_col}1", {
                "backgroundColor": {"red": 0.05, "green": 0.25, "blue": 0.45},
                "horizontalAlignment": "CENTER",
                "textFormat": {
                    "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "bold": True, "fontSize": 11
                }
            })
            ws.freeze(rows=1)
        
        return {"status": "success", "title": sh.title}
    except Exception as e:
        return {"status": "error", "message": str(e)}


def _col_to_letter(col):
    """Convert 1-based column number to spreadsheet letter (e.g. 1→A, 27→AA)."""
    let = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        let = chr(65 + remainder) + let
    return let


def push_row_to_gsheets(sheet_id, service_account_json, row_dict):
    """Safely updates or appends a row in Google Sheets. Handles receiver_name columns."""
    import gspread
    from google.oauth2.service_account import Credentials
    import json
    
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = json.loads(service_account_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    
    sh = client.open_by_key(sheet_id)
    ws = sh.worksheet("Routing Records")
    
    # Determine max stage from row_dict
    stages = []
    for k in row_dict.keys():
        if k.startswith("receiving_office_"):
            try:
                stages.append(int(k.split("_")[-1]))
            except:
                pass
    max_stage = max(stages) if stages else 10
    if max_stage < 10:
        max_stage = 10
    
    target_headers = ["Department", "School/Office", "Employee", "Code"]
    for i in range(1, max_stage + 1):
        target_headers += [f"Receiving Office {i}", f"Receiver Name {i}", f"Timestamp {i}"]
    
    row_data = [
        row_dict.get('department', ''),
        row_dict.get('school_office', ''),
        row_dict.get('employee', ''),
        row_dict.get('code', '')
    ]
    for i in range(1, max_stage + 1):
        row_data.append(row_dict.get(f'receiving_office_{i}') or "")
        row_data.append(row_dict.get(f'receiver_name_{i}') or "")
        row_data.append(row_dict.get(f'timestamp_{i}') or "")
    
    # Expand headers if needed
    all_rows = ws.get_all_values()
    existing_headers = all_rows[0] if all_rows else []
    if len(existing_headers) < len(target_headers):
        end_col = _col_to_letter(len(target_headers))
        ws.update(range_name=f"A1:{end_col}1", values=[target_headers])
        ws.format(f"A1:{end_col}1", {
            "backgroundColor": {"red": 0.05, "green": 0.25, "blue": 0.45},
            "horizontalAlignment": "CENTER",
            "textFormat": {
                "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                "bold": True, "fontSize": 11
            }
        })
    
    # Find matching row by Code + Employee
    match_row_num = None
    code_to_match = str(row_dict.get('code', '')).strip().upper()
    emp_to_match  = str(row_dict.get('employee', '')).strip().lower()
    for idx, row in enumerate(all_rows[1:], start=2):
        if len(row) >= 4:
            if str(row[2]).strip().lower() == emp_to_match and str(row[3]).strip().upper() == code_to_match:
                match_row_num = idx
                break
    
    end_col = _col_to_letter(len(target_headers))
    if match_row_num:
        ws.update(range_name=f"A{match_row_num}:{end_col}{match_row_num}", values=[row_data])
        print(f"Updated Google Sheets row {match_row_num} for {row_dict.get('employee')}")
    else:
        ws.append_row(row_data)
        print(f"Appended new Google Sheets row for {row_dict.get('employee')}")

def start_gsheets_worker():
    """Start background queue consumer for Google Sheets sync."""
    def _worker():
        print("Google Sheets Background Worker Started.")
        while True:
            try:
                task = gs_sync_queue.get()
                if task is None:
                    break
                enabled  = get_setting('gsheets_enabled') == 'True'
                sheet_id = get_setting('gsheets_id')
                # Use decrypted credentials for the worker
                creds    = get_decrypted_setting('gsheets_credentials')
                if enabled and sheet_id and creds:
                    try:
                        row_id = task.get('row_id')
                        conn = get_db_connection()
                        row = conn.execute("SELECT * FROM routing_records WHERE id = ?", (row_id,)).fetchone()
                        conn.close()
                        if row:
                            push_row_to_gsheets(sheet_id, creds, dict(row))
                    except Exception as e:
                        print(f"Google Sheet sync worker error: {e}")
                else:
                    print("Google Sheet Sync disabled. Skipping.")
                gs_sync_queue.task_done()
                time.sleep(1)  # Rate-limit protection
            except Exception as ex:
                print(f"Worker critical error: {ex}")
                time.sleep(5)
    
    worker_thread = threading.Thread(target=_worker, daemon=True)
    worker_thread.start()


def queue_sync(row_id):
    """Queue a row update to the background Google Sheets worker."""
    gs_sync_queue.put({"row_id": row_id})


def bulk_sync_to_gsheets():
    """Push ALL existing routing records to Google Sheets. Returns (count, error_messages)."""
    enabled = get_setting('gsheets_enabled') == 'True'
    sheet_id = get_setting('gsheets_id')
    creds = get_decrypted_setting('gsheets_credentials')
    
    if not enabled or not sheet_id or not creds:
        return 0, ["Google Sheets sync is not configured. Enable it and save credentials first."]
    
    errors = []
    count = 0
    
    try:
        conn = get_db_connection()
        rows = conn.execute("SELECT * FROM routing_records").fetchall()
        conn.close()
        
        for row in rows:
            try:
                row_dict = dict(row)
                push_row_to_gsheets(sheet_id, creds, row_dict)
                count += 1
            except Exception as e:
                emp = row['employee'] if row['employee'] else '?'
                code = row['code'] if row['code'] else '?'
                errors.append(f"Row {emp}/{code}: {str(e)[:100]}")
        
        return count, errors
    except Exception as e:
        return 0, [str(e)]

def force_import_from_excel():
    """Force reload data from Excel without exporting back - preserves manual Excel edits."""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    imported_count = 0
    
    # 1. Import code_lookup
    if os.path.exists(CODE_LOOKUP_EXCEL):
        try:
            df = pd.read_excel(CODE_LOOKUP_EXCEL)
            for _, row in df.iterrows():
                code = str(row.get('Code', '')).strip()
                if code:
                    # Check if code already exists in SQLite
                    cursor.execute("SELECT code FROM code_lookup WHERE code = ?", (code,))
                    if not cursor.fetchone():
                        dept = str(row.get('Department', 'Unknown')).strip()
                        school = str(row.get('School/Office', 'Unknown')).strip()
                        employees = str(row.get('Employees', '')).strip()
                        doc_type = str(row.get('Doc Type', '')).strip()
                        gen = str(row.get('Generated', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))).strip()
                        cursor.execute(
                            "INSERT INTO code_lookup (code, department, school_office, employees, doc_type, generated_at) "
                            "VALUES (?, ?, ?, ?, ?, ?)",
                            (code, dept, school, employees, doc_type, gen)
                        )
                        imported_count += 1
            conn.commit()
            print(f"Force imported {imported_count} code lookup records.")
        except Exception as e:
            print(f"Error force importing code lookup: {e}")
    
    # 2. Import routing_records
    if os.path.exists(ROUTING_RECORD_EXCEL):
        try:
            df = pd.read_excel(ROUTING_RECORD_EXCEL)
            for _, row in df.iterrows():
                code = str(row.get('Code', '')).strip()
                emp = str(row.get('Employee', '')).strip()
                if code and emp:
                    # Check if record already exists in SQLite
                    cursor.execute("SELECT id FROM routing_records WHERE code = ? AND employee = ?", (code, emp))
                    if not cursor.fetchone():
                        dept = str(row.get('Department', '')).strip()
                        school = str(row.get('School/Office', '')).strip()
                        # Get all receiving_office columns dynamically
                        vals = [dept, school, emp, code]
                        for col in df.columns:
                            if 'Receiving Office' in col or 'Timestamp' in col or 'Receiver Name' in col:
                                vals.append(str(row.get(col, '')))
                        placeholders = ['?'] * len(vals)
                        col_names = ['department', 'school_office', 'employee', 'code']
                        for col in df.columns:
                            if 'Receiving Office' in col:
                                col_names.append(col)
                            elif 'Timestamp' in col:
                                col_names.append(col)
                            elif 'Receiver Name' in col:
                                col_names.append(col)
                        sql = f"INSERT INTO routing_records ({', '.join(col_names)}) VALUES ({', '.join(placeholders)})"
                        cursor.execute(sql, vals)
                        imported_count += 1
            conn.commit()
            print(f"Force imported routing records.")
        except Exception as e:
            print(f"Error force importing routing records: {e}")
    
    conn.close()
    return imported_count