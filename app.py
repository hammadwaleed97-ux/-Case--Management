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
# ==================================================
# ⚖️ PART 2-A-1
# GOVERNMENT CASE MANAGEMENT SYSTEM
# تسجيل القضايا - الواجهة الرئيسية
# ==================================================

# ==================================================
# SESSION STATE
# ==================================================
if "page" not in st.session_state:
    st.session_state.page = "home"

if "selected_case" not in st.session_state:
    st.session_state.selected_case = None

# ==================================================
# COURT THEME
# ==================================================
st.markdown("""
<style>

.stApp{
    background:#08284d;
}

h1,h2,h3,label{
    color:white!important;
}

.block-container{
    padding-top:1rem;
}

.main-title{
    text-align:center;
    color:gold;
    font-size:38px;
    font-weight:bold;
}

.sub-title{
    text-align:center;
    color:white;
    font-size:18px;
}

div[data-testid="stForm"]{
    background:#103b6d;
    padding:20px;
    border-radius:15px;
    border:2px solid gold;
}

.stButton>button{
    background:#1d5da8;
    color:white;
    border-radius:8px;
    border:none;
    font-weight:bold;
}

.stButton>button:hover{
    background:#2f7bd6;
}

.footer{
    text-align:center;
    color:gold;
    font-size:14px;
    margin-top:30px;
}

</style>
""", unsafe_allow_html=True)

# ==================================================
# HEADER
# ==================================================

st.markdown("""
<div class='main-title'>
⚖️ إدارة القضايا
</div>

<div class='sub-title'>
الهيئة القومية للتأمين الاجتماعى
</div>

<div class='sub-title'>
الإدارة المركزية للشئون القانونية
</div>

<div class='sub-title'>
الإدارة العامة للقضايا
</div>
""", unsafe_allow_html=True)

# ==================================================
# HOME PAGE
# ==================================================

def home_page():

    st.markdown("## الصفحة الرئيسية")

    c1,c2,c3=st.columns(3)

    with c1:
        if st.button("📌 تسجيل القضايا",use_container_width=True):
            st.session_state.page="register"

    with c2:
        if st.button("📂 الحصر العام",use_container_width=True):
            st.session_state.page="registry"

    with c3:
        if st.button("🔔 التنبيهات",use_container_width=True):
            st.session_state.page="alerts"

    c4,c5,c6=st.columns(3)

    with c4:
        if st.button("📊 التقارير",use_container_width=True):
            st.session_state.page="reports"

    with c5:
        if st.button("🗃 الأرشيف",use_container_width=True):
            st.session_state.page="archive"

    with c6:
        if st.button("📚 المكتبة القانونية",use_container_width=True):
            st.session_state.page="library"

    c7,c8=st.columns(2)

    with c7:
        if st.button("🔎 البحث عن دعوى",use_container_width=True):
            st.session_state.page="search"

    with c8:
        if st.button("📈 الإحصائيات",use_container_width=True):
            st.session_state.page="statistics"

# ==================================================
# REGISTER PAGE
# ==================================================

def register_case():

    st.markdown("## ⚖️ تسجيل القضايا")

    with st.form("register_case"):

        case_kind=st.selectbox(
            "نوع الدعوى",
            ["دعوى","استئناف","طعن"]
        )

        court_type=st.selectbox(
            "المحكمة",
            [
                "الابتدائية",
                "الاستئناف",
                "النقض",
                "الإدارية",
                "القضاء الإدارى",
                "الإدارية العليا"
            ]
        )

        court_name=st.text_input("اسم المحكمة")

        mission=""

        if court_type=="الاستئناف":
            mission=st.text_input("المأمورية")

        case_number=st.text_input("رقم الدعوى / الاستئناف / الطعن")

        judicial_year=st.text_input("السنة القضائية")

        circuit=st.text_input("الدائرة")

        case_type=st.text_input("النوع")

        plaintiff=st.text_input(
            "اسم المدعى / المستأنف / الطاعن"
        )

        defendant=st.text_input(
            "اسم المدعى عليه / المستأنف ضده / المطعون ضده"
        )

        subject=st.text_area("موضوع الدعوى")

        first_session=st.date_input("تاريخ أول جلسة")

        roll=st.text_input("الرول")

        required_action=st.text_area("الإجراء المطلوب")

        notes=st.text_area("ملاحظات")
