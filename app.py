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
