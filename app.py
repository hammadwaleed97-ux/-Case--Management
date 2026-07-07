# ============================================================
# ⚖️ نظام إدارة القضايا
# إعداد وتطوير : وليد شعبان حماد
# الهيئة القومية للتأمين الاجتماعى
# الإدارة المركزية للشئون القانونية
# الإدارة العامة للقضايا
# ============================================================

# ============================================================
# IMPORTS
# ============================================================

import streamlit as st
import sqlite3
import pandas as pd
import os
import shutil
from datetime import datetime, timedelta
from pathlib import Path

# =====================================
# تصدير Word
# =====================================

from docx import Document
from docx.shared import Pt,Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn

# =====================================
# تصدير PDF
# =====================================

from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer
)

from reportlab.lib.styles import getSampleStyleSheet

from reportlab.lib.enums import TA_CENTER

from reportlab.lib import colors

# =====================================
# APP CONFIG
# =====================================

st.set_page_config(

    page_title="إدارة القضايا",

    page_icon="⚖️",

    layout="wide",

    initial_sidebar_state="collapsed"

)

# ============================================================
# إنشاء المجلدات
# ============================================================

BASE_DIR = Path(".")

DATABASE = BASE_DIR / "government_cases.db"

UPLOAD_DIR = BASE_DIR / "uploads"

LIBRARY_DIR = BASE_DIR / "legal_library"

REPORT_DIR = BASE_DIR / "reports"

BACKUP_DIR = BASE_DIR / "backup"

UPLOAD_DIR.mkdir(exist_ok=True)

LIBRARY_DIR.mkdir(exist_ok=True)

REPORT_DIR.mkdir(exist_ok=True)

BACKUP_DIR.mkdir(exist_ok=True)

# ============================================================
# DATABASE CONNECTION
# ============================================================

conn = sqlite3.connect(

    DATABASE,

    check_same_thread=False

)

conn.row_factory = sqlite3.Row

cur = conn.cursor()

# ============================================================
# إنشاء الجداول
# ============================================================

def create_tables():

    # ====================================================
    # جدول القضايا
    # ====================================================

    cur.execute("""

    CREATE TABLE IF NOT EXISTS cases(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        case_type TEXT,

        court_level TEXT,

        court_name TEXT,

        mission TEXT,

        case_number TEXT,

        judicial_year TEXT,

        circuit TEXT,

        case_category TEXT,

        plaintiff TEXT,

        defendant TEXT,

        subject TEXT,

        first_session_date TEXT,

        roll TEXT,

        required_action TEXT,

        notes TEXT,

        whatsapp_enabled INTEGER DEFAULT 0,

        whatsapp_number TEXT,

        last_session_date TEXT,

        last_session_reason TEXT,

        judgment_date TEXT,

        judgment_text TEXT,

        judgment_result TEXT,

        appeal_number TEXT,

        archive_note TEXT,

        status TEXT,

        created_at TEXT,

        updated_at TEXT

    )

    """)

    # ====================================================
    # جدول الجلسات
    # ====================================================

    cur.execute("""

    CREATE TABLE IF NOT EXISTS sessions(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        case_id INTEGER,

        session_date TEXT,

        roll TEXT,

        procedure_text TEXT,

        created_at TEXT,

        FOREIGN KEY(case_id)

        REFERENCES cases(id)

    )

    """)

    # ====================================================
    # جدول المستندات
    # ====================================================

    cur.execute("""

    CREATE TABLE IF NOT EXISTS case_documents(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        case_id INTEGER,

        document_type TEXT,

        document_title TEXT,

        file_name TEXT,

        file_path TEXT,

        upload_date TEXT,

        FOREIGN KEY(case_id)

        REFERENCES cases(id)

    )

    """)
        # ====================================================
    # جدول الأحكام
    # ====================================================

    cur.execute("""

    CREATE TABLE IF NOT EXISTS judgments(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        case_id INTEGER,

        judgment_date TEXT,

        judgment_text TEXT,

        judgment_result TEXT,

        judgment_type TEXT,

        created_at TEXT,

        FOREIGN KEY(case_id)

        REFERENCES cases(id)

    )

    """)

    # ====================================================
    # جدول الأرشيف
    # ====================================================

    cur.execute("""

    CREATE TABLE IF NOT EXISTS archive(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        case_id INTEGER,

        archive_date TEXT,

        appeal_number TEXT,

        archive_note TEXT,

        created_at TEXT,

        FOREIGN KEY(case_id)

        REFERENCES cases(id)

    )

    """)

    # ====================================================
    # جدول التنبيهات
    # ====================================================

    cur.execute("""

    CREATE TABLE IF NOT EXISTS notifications(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        case_id INTEGER,

        notification_type TEXT,

        notification_date TEXT,

        whatsapp_enabled INTEGER DEFAULT 0,

        whatsapp_number TEXT,

        sent INTEGER DEFAULT 0,

        created_at TEXT,

        FOREIGN KEY(case_id)

        REFERENCES cases(id)

    )

    """)

    # ====================================================
    # جدول المكتبة القانونية
    # ====================================================

    cur.execute("""

    CREATE TABLE IF NOT EXISTS legal_library(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        category TEXT,

        document_name TEXT,

        document_number TEXT,

        document_year TEXT,

        description TEXT,

        file_name TEXT,

        file_path TEXT,

        created_at TEXT

    )

    """)

    # ====================================================
    # جدول المستخدمين
    # ====================================================

    cur.execute("""

    CREATE TABLE IF NOT EXISTS users(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        full_name TEXT,

        username TEXT UNIQUE,

        password TEXT,

        role TEXT,

        active INTEGER DEFAULT 1,

        created_at TEXT

    )

    """)

    # ====================================================
    # جدول سجل الحركات
    # ====================================================

    cur.execute("""

    CREATE TABLE IF NOT EXISTS audit_log(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        username TEXT,

        action TEXT,

        details TEXT,

        action_date TEXT

    )

    """)

    conn.commit()

create_tables()

# ====================================================
# دوال قاعدة البيانات العامة
# ====================================================

def execute(query, params=()):

    cur.execute(query, params)

    conn.commit()

    return cur

def fetch_all(query, params=()):

    cur.execute(query, params)

    return cur.fetchall()

def fetch_one(query, params=()):

    cur.execute(query, params)

    return cur.fetchone()

# ====================================================
# تسجيل العمليات
# ====================================================

def log_action(username, action, details):

    execute("""

    INSERT INTO audit_log(

    username,

    action,

    details,

    action_date

    )

    VALUES(

    ?,?,?,?

    )

    """,(

    username,

    action,

    details,

    datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ))

# ====================================================
# إنشاء نسخة احتياطية
# ====================================================

def backup_database():

    backup_name = BACKUP_DIR / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"

    shutil.copy(DATABASE, backup_name)

# ====================================================
# نهاية الجزء 1-ب
# ====================================================
