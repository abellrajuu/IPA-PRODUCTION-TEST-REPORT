-- APEX ProScale™ Industrial Suite Database Schema
-- File: database_sql_scripts/schema.sql
-- Target Database: database_sqlite.db (SQLite 3)

-- 1. Users Table (Authentication & Access Control)
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    full_name TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'Production Engineer',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 2. Records Table (Production & Calibration Test Reports)
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
);

-- 3. ST-ERP Offline Query Cache
CREATE TABLE IF NOT EXISTS sterp_job_cache (
    job_no TEXT PRIMARY KEY,
    response_json TEXT NOT NULL,
    customer TEXT,
    has_multiple_systems INTEGER DEFAULT 0,
    cached_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);