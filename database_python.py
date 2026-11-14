import sqlite3
import json
import os
import re
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

DB_PATH = os.path.join(BASE_DIR, "database_sqlite.db")

def get_db_connection():
    conn = sqlite3.connect(DB_PATH, timeout=15.0)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = sqlite3.connect(DB_PATH, timeout=15.0)
    try:
        cursor = conn.cursor()
        # Enable WAL mode for high concurrency
        cursor.execute("PRAGMA journal_mode=WAL;")
        
        # 1. Test records table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_name TEXT NOT NULL,
                job_no TEXT NOT NULL,
                customer TEXT,
                conveyor_no TEXT,
                capacity TEXT,
                created_at TEXT NOT NULL,
                data_json TEXT NOT NULL,
                docx_filename TEXT NOT NULL,
                pdf_filename TEXT,
                status TEXT DEFAULT 'Pending Verification',
                tested_by TEXT,
                verified_by TEXT,
                verified_at TEXT,
                hod_by TEXT,
                hod_at TEXT
            )
        """)

        # Non-destructive additive schema migrations for existing SQLite databases
        cursor.execute("PRAGMA table_info(records)")
        existing_cols = {row[1] for row in cursor.fetchall()}
        required_migrations = {
            "pdf_filename": "TEXT",
            "status": "TEXT DEFAULT 'Pending Verification'",
            "tested_by": "TEXT",
            "verified_by": "TEXT",
            "verified_at": "TEXT",
            "hod_by": "TEXT",
            "hod_at": "TEXT",
        }
        for col_name, col_def in required_migrations.items():
            if col_name not in existing_cols:
                cursor.execute(f"ALTER TABLE records ADD COLUMN {col_name} {col_def}")

        # 2. User accounts table
                # 3. ST-ERP Offline Query Cache
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sterp_job_cache (
                job_no TEXT PRIMARY KEY,
                response_json TEXT NOT NULL,
                customer TEXT,
                has_multiple_systems INTEGER DEFAULT 0,
                cached_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                full_name TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'Production Engineer',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Read admin and tech credentials securely from .env
        admin_user = os.environ.get("ADMIN_USERNAME", "admin").strip()
        admin_pass = os.environ.get("ADMIN_PASSWORD", "Admin@2026!").strip()
        tech_user = os.environ.get("TECH_USERNAME", "tech").strip()
        tech_pass = os.environ.get("TECH_PASSWORD", "Tech@2026!").strip()

        # Seed admin and technician users with hashed passwords if empty
        cursor.execute("SELECT COUNT(*) FROM users")
        if cursor.fetchone()[0] == 0:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            admin_hash = generate_password_hash(admin_pass)
            tech_hash = generate_password_hash(tech_pass)
            cursor.execute("INSERT INTO users (username, password_hash, full_name, role, created_at) VALUES (?, ?, ?, ?, ?)",
                           (admin_user, admin_hash, "System Administrator", "Admin", now))
            cursor.execute("INSERT INTO users (username, password_hash, full_name, role, created_at) VALUES (?, ?, ?, ?, ?)",
                           (tech_user, tech_hash, "Field Technician", "Production Engineer", now))
        else:
            # Sync admin password from .env if admin account exists
            cursor.execute("SELECT id, username, password_hash FROM users WHERE username = ?", (admin_user,))
            admin_row = cursor.fetchone()
            if admin_row:
                admin_hash = generate_password_hash(admin_pass)
                cursor.execute("UPDATE users SET password_hash = ? WHERE id = ?", (admin_hash, admin_row[0]))

            # Re-hash plain text passwords if legacy accounts exist
            cursor.execute("SELECT id, username, password_hash FROM users")
            rows = cursor.fetchall()
            valid_prefixes = ("scrypt:", "pbkdf2:", "argon2:", "bcrypt:", "$2b$", "$2a$")
            for user_id, uname, pwd_hash in rows:
                if pwd_hash and not any(pwd_hash.startswith(prefix) for prefix in valid_prefixes):
                    new_hash = generate_password_hash(pwd_hash)
                    cursor.execute("UPDATE users SET password_hash = ? WHERE id = ?", (new_hash, user_id))

        conn.commit()
    finally:
        conn.close()

def save_record(user_name, job_no, customer, conveyor_no, capacity, data_dict, docx_filename, pdf_filename=None, tested_by=None):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        now_dt = datetime.now()
        created_at = now_dt.strftime("%Y-%m-%d %H:%M:%S")
        data_json = json.dumps(data_dict)
        tested_by_val = (tested_by or (data_dict.get("tested_by") if isinstance(data_dict, dict) else "") or user_name or "").strip()
        if not pdf_filename and docx_filename and docx_filename.lower().endswith(".docx"):
            pdf_filename = docx_filename[:-5] + ".pdf"

        # Check for rapid identical duplicate submission (within 10 seconds)
        cursor.execute("""
            SELECT id, created_at, data_json FROM records 
            WHERE job_no = ? AND user_name = ? AND docx_filename = ? 
            ORDER BY id DESC LIMIT 1
        """, (job_no, user_name, docx_filename))
        recent_row = cursor.fetchone()
        if recent_row:
            try:
                prev_time = datetime.strptime(recent_row["created_at"], "%Y-%m-%d %H:%M:%S")
                if (now_dt - prev_time).total_seconds() < 10 and recent_row["data_json"] == data_json:
                    # Return existing record ID only for identical double-click duplicates
                    return recent_row["id"]
            except Exception:
                pass

        cursor.execute("""
            INSERT INTO records (user_name, job_no, customer, conveyor_no, capacity, created_at, data_json, docx_filename, pdf_filename, tested_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (user_name, job_no, customer, conveyor_no, capacity, created_at, data_json, docx_filename, pdf_filename, tested_by_val))

        record_id = cursor.lastrowid
        conn.commit()
        return record_id
    finally:
        conn.close()

def get_all_records():
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, user_name, job_no, customer, conveyor_no, capacity, created_at,
                   docx_filename, pdf_filename, status, tested_by, verified_by, verified_at, hod_by, hod_at
            FROM records
            ORDER BY id DESC
        """)
        rows = cursor.fetchall()
        results = []
        for row in rows:
            item = dict(row)
            if not item.get("pdf_filename") and item.get("docx_filename", "").lower().endswith(".docx"):
                item["pdf_filename"] = item["docx_filename"][:-5] + ".pdf"
            if not item.get("status"):
                item["status"] = "Pending Verification"
            results.append(item)
        return results
    finally:
        conn.close()

def get_record_by_id(record_id):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM records WHERE id = ?", (record_id,))
        row = cursor.fetchone()
        if row:
            res = dict(row)
            res["data"] = json.loads(res["data_json"])
            return res
        return None
    finally:
        conn.close()

def get_all_users():
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, full_name, role, created_at FROM users ORDER BY id ASC")
        rows = cursor.fetchall()
        return [dict(row) for row in rows]
    finally:
        conn.close()

def get_user_by_id(user_id):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, full_name, role, created_at FROM users WHERE id = ?", (user_id,))
        row = cursor.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()

def get_user_by_username(username):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, full_name, role, created_at FROM users WHERE username = ?", (username,))
        row = cursor.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()

ALLOWED_ROLES = {"Admin", "Production Engineer", "Verifying Engineer", "HOD"}

def create_user(username, password, full_name, role):
    username = (username or "").strip()
    if not re.match(r'^[a-zA-Z0-9_\.\-]{3,50}$', username):
        raise ValueError("Username must be between 3 and 50 characters and contain only letters, numbers, dots, underscores, or hyphens.")
    
    if not password or len(password) < 6:
        raise ValueError("Password must be at least 6 characters long.")

    full_name = (full_name or "").strip() or username
    role = (role or "Production Engineer").strip()
    if role not in ALLOWED_ROLES:
        raise ValueError(f"Invalid role '{role}'. Allowed roles: {', '.join(sorted(ALLOWED_ROLES))}")

    if get_user_by_username(username):
        raise ValueError(f"Username '{username}' already exists.")
        
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        hashed_pwd = generate_password_hash(password)
        cursor.execute("INSERT INTO users (username, password_hash, full_name, role, created_at) VALUES (?, ?, ?, ?, ?)",
                       (username, hashed_pwd, full_name, role, created_at))
        user_id = cursor.lastrowid
        conn.commit()
        return user_id
    finally:
        conn.close()

def update_user_role(user_id, role):
    role = (role or "").strip()
    if role not in ALLOWED_ROLES:
        raise ValueError(f"Invalid role '{role}'. Allowed roles: {', '.join(sorted(ALLOWED_ROLES))}")

    user = get_user_by_id(user_id)
    if not user:
        raise ValueError("User not found.")
    if user["username"] == "admin" and role != "Admin":
        raise ValueError("Cannot revoke Admin role from primary admin account.")
        
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET role = ? WHERE id = ?", (role, user_id))
        conn.commit()
        return True
    finally:
        conn.close()

def delete_user(user_id):
    user = get_user_by_id(user_id)
    if not user:
        raise ValueError("User not found.")
    if user["username"] == "admin":
        raise ValueError("Primary superadmin account 'admin' cannot be deleted.")

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
        conn.commit()
        return True
    finally:
        conn.close()

def verify_user(username, password):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, password_hash, full_name, role FROM users WHERE username = ?", (username,))
        row = cursor.fetchone()
        if row:
            user_dict = dict(row)
            stored_hash = user_dict.pop("password_hash")
            if check_password_hash(stored_hash, password):
                return user_dict
        return None
    finally:
        conn.close()

def get_pending_records(role=None, only_pending=False):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        if only_pending and role == 'Verifying Engineer':
            cursor.execute("SELECT * FROM records WHERE COALESCE(status, 'Pending Verification') = 'Pending Verification' ORDER BY id DESC")
        elif only_pending and role == 'HOD':
            cursor.execute("SELECT * FROM records WHERE COALESCE(status, 'Pending Verification') IN ('Pending HOD', 'Pending Verification') ORDER BY id DESC")
        else:
            cursor.execute("SELECT * FROM records ORDER BY id DESC")
        rows = cursor.fetchall()
        results = []
        for row in rows:
            item = dict(row)
            if not item.get("pdf_filename") and item.get("docx_filename", "").lower().endswith(".docx"):
                item["pdf_filename"] = item["docx_filename"][:-5] + ".pdf"
            if not item.get("status"):
                item["status"] = "Pending Verification"
            results.append(item)
        return results
    finally:
        conn.close()

def update_record_approval(record_id, user_full_name, role):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM records WHERE id = ?", (record_id,))
        row = cursor.fetchone()
        if not row:
            return None
            
        data = dict(row)
        curr_status = (data.get('status') or 'Pending Verification').strip()
        if curr_status == 'Approved':
            raise ValueError("Record is already approved.")

        data_dict = json.loads(data['data_json']) if data.get('data_json') else {}
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        approver_name = (user_full_name or "").strip() or role
        
        if role in ['Verifying Engineer', 'Admin'] and curr_status == 'Pending Verification':
            status = 'Pending HOD'
            data_dict['verified_by'] = approver_name
            data_dict['verified_at'] = now_str
            data_dict['approved_by'] = approver_name + " (Verified)"
            new_json = json.dumps(data_dict)
            cursor.execute(
                "UPDATE records SET status = ?, verified_by = ?, verified_at = ?, data_json = ? WHERE id = ?",
                (status, approver_name, now_str, new_json, record_id)
            )
                
        elif role in ['HOD', 'Admin'] and curr_status in ['Pending HOD', 'Pending Verification']:
            status = 'Approved'
            data_dict['hod_by'] = approver_name
            data_dict['hod_at'] = now_str
            data_dict['approved_by'] = approver_name + " (HOD)"
            new_json = json.dumps(data_dict)
            cursor.execute(
                "UPDATE records SET status = ?, hod_by = ?, hod_at = ?, data_json = ? WHERE id = ?",
                (status, approver_name, now_str, new_json, record_id)
            )
        else:
            raise ValueError(f"Cannot approve record with status '{curr_status}' as role '{role}'.")
            
        docx_filename = data['docx_filename']
        
        conn.commit()
        return data_dict, docx_filename
    finally:
        conn.close()

if __name__ == "__main__":
    init_db()
    print("Database initialized successfully at", DB_PATH)

def save_sterp_cache(job_no, data_dict):
    """Cache successful ST-ERP query result for offline resilience."""
    try:
        conn = get_db_connection()
        try:
            cursor = conn.cursor()
            clean_job = str(job_no).upper().strip()
            cust = data_dict.get("customer") or ""
            has_multi = 1 if data_dict.get("has_multiple_systems") else 0
            cursor.execute('''
                INSERT OR REPLACE INTO sterp_job_cache (job_no, response_json, customer, has_multiple_systems, cached_at)
                VALUES (?, ?, ?, ?, datetime('now', 'localtime'))
            ''', (clean_job, json.dumps(data_dict), cust, has_multi))
            conn.commit()
        finally:
            conn.close()
    except Exception as e:
        print(f"ST-ERP Cache save notice: {e}")

def get_sterp_cache(job_no):
    """Retrieve cached ST-ERP job data when SQL Server is offline or unreachable."""
    try:
        conn = get_db_connection()
        try:
            cursor = conn.cursor()
            clean_job = str(job_no).upper().strip()
            clean_raw = re.sub(r'^(JOB[-_]?)', '', clean_job)
            cursor.execute('''
                SELECT response_json FROM sterp_job_cache
                WHERE job_no = ? OR job_no = ? OR job_no = ?
            ''', (clean_job, clean_raw, f"JOB-{clean_raw}"))
            row = cursor.fetchone()
            if row and row["response_json"]:
                data = json.loads(row["response_json"])
                data["from_cache"] = True
                data["source"] = "sterp_cache"
                return data
        finally:
            conn.close()
    except Exception as e:
        print(f"ST-ERP Cache get notice: {e}")
    return None
