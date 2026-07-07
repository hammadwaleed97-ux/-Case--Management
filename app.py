# ==================================================
# ⚖️ GOVERNMENT CASE MANAGEMENT SYSTEM
# 🏛️ PART 1 - CORE FOUNDATION
# ==================================================

import streamlit as st
import sqlite3
from datetime import datetime
import os

# ==================================================
# ⚖️ APP CONFIG
# ==================================================
st.set_page_config(
    page_title="نظام إدارة القضايا الحكومي",
    page_icon="⚖️",
    layout="wide"
)

# ==================================================
# ⚖️ DATABASE CONNECTION
# ==================================================
conn = sqlite3.connect("gov_cases.db", check_same_thread=False)
c = conn.cursor()

# ==================================================
# ⚖️ INIT DATABASE (GOV STRUCTURE)
# ==================================================
def init_db():

    # ---------------- CASES ----------------
    c.execute("""
    CREATE TABLE IF NOT EXISTS cases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        case_number TEXT,
        year TEXT,
        court TEXT,
        plaintiff TEXT,
        defendant TEXT,
        subject TEXT,
        status TEXT,
        last_session TEXT,
        created_at TEXT
    )
    """)

    # ---------------- SESSIONS ----------------
    c.execute("""
    CREATE TABLE IF NOT EXISTS sessions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        case_id INTEGER,
        session_date TEXT,
        roll TEXT,
        decision TEXT,
        created_at TEXT
    )
    """)

    # ---------------- DOCUMENTS ----------------
    c.execute("""
    CREATE TABLE IF NOT EXISTS documents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        case_id INTEGER,
        doc_type TEXT,
        file_path TEXT,
        created_at TEXT
    )
    """)

    # ---------------- ARCHIVE ----------------
    c.execute("""
    CREATE TABLE IF NOT EXISTS archive (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        case_id INTEGER,
        judgment TEXT,
        result TEXT,
        judgment_date TEXT,
        created_at TEXT
    )
    """)

    # ---------------- AUDIT LOG ----------------
    c.execute("""
    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        action TEXT,
        details TEXT,
        created_at TEXT
    )
    """)

    conn.commit()

init_db()

# ==================================================
# ⚖️ AUDIT SYSTEM
# ==================================================
def log_action(action, details):

    c.execute("""
        INSERT INTO audit_log VALUES (NULL,?,?,?)
    """, (
        action,
        details,
        str(datetime.now())
    ))

    conn.commit()

# ==================================================
# ⚖️ FILE STORAGE
# ==================================================
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ==================================================
# ⚖️ END PART 1
# ==================================================
