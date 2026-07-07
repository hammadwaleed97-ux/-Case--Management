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
# ============================================================
# CSS الحكومى الفخم
# ============================================================

st.markdown("""
<style>

html,body,.stApp{

    background:#082B52;

}

.block-container{

    padding-top:10px;

    padding-bottom:20px;

}

.main-title{

    text-align:center;

    color:#FFD700;

    font-size:42px;

    font-weight:bold;

}

.second-title{

    text-align:center;

    color:white;

    font-size:18px;

}

.case-card{

    background:#103C73;

    border:2px solid #D4AF37;

    border-radius:12px;

    padding:15px;

    margin-bottom:12px;

}

table{

    width:100%;

}

thead tr{

    background:#0B4A86;

    color:white;

}

tbody tr{

    background:white;

    color:black;

}

.stButton>button{

    width:100%;

    border-radius:8px;

    background:#0F5EA8;

    color:white;

    font-weight:bold;

    border:none;

    height:45px;

}

.stButton>button:hover{

    background:#D4AF37;

    color:black;

}

div[data-testid="stSidebar"]{

    display:none;

}

footer{

    visibility:hidden;

}

header{

    visibility:hidden;

}

</style>

""",unsafe_allow_html=True)

# ============================================================
# إنشاء الصفحات
# ============================================================

if "page" not in st.session_state:

    st.session_state.page="home"

if "selected_case" not in st.session_state:

    st.session_state.selected_case=None

# ============================================================
# الصفحة الرئيسية
# ============================================================

def home_page():

    st.markdown("""

    <div class='main-title'>

    ⚖️ إدارة القضايا

    </div>

    """,unsafe_allow_html=True)

    st.markdown("""

    <div class='second-title'>

    الهيئة القومية للتأمين الاجتماعى

    <br>

    الإدارة المركزية للشئون القانونية

    <br>

    الإدارة العامة للقضايا

    </div>

    """,unsafe_allow_html=True)

    st.write("")

    a,b,c=st.columns(3)

    with a:

        if st.button("📑 تسجيل القضايا"):

            st.session_state.page="register"

            st.rerun()

    with b:

        if st.button("📂 الحصر العام"):

            st.session_state.page="general"

            st.rerun()

    with c:

        if st.button("🔔 التنبيهات"):

            st.session_state.page="notifications"

            st.rerun()

    d,e,f=st.columns(3)

    with d:

        if st.button("📊 التقارير"):

            st.session_state.page="reports"

            st.rerun()

    with e:

        if st.button("🗄️ الأرشيف"):

            st.session_state.page="archive"

            st.rerun()

    with f:

        if st.button("📚 المكتبة القانونية"):

            st.session_state.page="library"

            st.rerun()

    g,h=st.columns(2)

    with g:

        if st.button("🔍 البحث عن دعوى"):

            st.session_state.page="search"

            st.rerun()

    with h:

        if st.button("📈 الإحصائيات"):

            st.session_state.page="statistics"

            st.rerun()

    st.write("")

    st.markdown("""

    <h5 style='text-align:center;color:gold'>

    مع تحيات / وليد شعبان حماد

    <br>

    الإدارة العامة للشئون القانونية

    <br>

    ديوان عام منطقة البحيرة

    <br>

    الهيئة القومية للتأمين الاجتماعى

    </h5>

    """,unsafe_allow_html=True)

# ============================================================
# تشغيل الصفحة الرئيسية
# ============================================================

if st.session_state.page=="home":

    home_page()

# ============================================================
# نهاية الجزء 1-ج
# ============================================================
