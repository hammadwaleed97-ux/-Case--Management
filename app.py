# ==========================================================
# الجزء 1 - الأساس والتصميم العام (Core + UI Shell)
# ==========================================================

import streamlit as st
import sqlite3
from datetime import datetime

# ===================== إعداد الصفحة =====================
st.set_page_config(
    page_title="إدارة القضايا",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ===================== قاعدة البيانات =====================
conn = sqlite3.connect("cases.db", check_same_thread=False)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS cases (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    case_type TEXT,
    court TEXT,
    court_name TEXT,
    mission TEXT,
    case_number TEXT,
    year TEXT,
    circle TEXT,
    case_kind TEXT,
    plaintiff TEXT,
    defendant TEXT,
    subject TEXT,
    first_session_date TEXT,
    roll TEXT,
    action TEXT,
    notes TEXT,
    whatsapp_enabled INTEGER,
    whatsapp_number TEXT,
    doc_type TEXT,
    created_at TEXT
)
""")

conn.commit()

# ===================== CSS تصميم فخم =====================
st.markdown("""
<style>

@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;700&display=swap');

html, body, [class*="css"]  {
    font-family: 'Cairo', sans-serif;
    background-color: #0b1f3a;
    color: #ffffff;
}

.main {
    background: linear-gradient(180deg, #071427, #0b1f3a);
}

.block-container {
    padding: 2rem;
}

h1, h2, h3 {
    color: #00d4ff;
}

.stButton>button {
    background: linear-gradient(90deg, #00d4ff, #005eff);
    color: white;
    border-radius: 10px;
    border: none;
    padding: 10px 20px;
    font-weight: bold;
}

.stTextInput>div>div>input {
    background-color: #112b4a;
    color: white;
    border-radius: 8px;
}

.sidebar .sidebar-content {
    background-color: #071427;
}

.card {
    background-color: #112b4a;
    padding: 15px;
    border-radius: 15px;
    margin-bottom: 10px;
    border: 1px solid #1f3b5c;
}

.logo {
    text-align: center;
    font-size: 28px;
    font-weight: bold;
    color: #00d4ff;
    margin-bottom: 5px;
}

.scale {
    text-align: center;
    font-size: 35px;
    color: gold;
    animation: glow 2s infinite;
}

@keyframes glow {
    0% {opacity: 0.4;}
    50% {opacity: 1;}
    100% {opacity: 0.4;}
}

.footer {
    position: fixed;
    bottom: 0;
    width: 100%;
    text-align: center;
    color: #00d4ff;
    background-color: rgba(0,0,0,0.3);
    padding: 5px;
    animation: fade 3s infinite;
}

@keyframes fade {
    0% {opacity: 0.2;}
    50% {opacity: 1;}
    100% {opacity: 0.2;}
}

</style>
""", unsafe_allow_html=True)

# ===================== الهيدر =====================
st.markdown("<div class='scale'>⚖️</div>", unsafe_allow_html=True)
st.markdown("<div class='logo'>إدارة القضايا</div>", unsafe_allow_html=True)

st.markdown("""
<div style='text-align:center;color:#ffffff;'>
مع تحيات / وليد شعبان حماد<br>
الإدارة العامة للشئون القانونية ـ ديوان عام منطقة البحيرة ـ الهيئة القومية للتأمين الاجتماعى
</div>
""", unsafe_allow_html=True)

# ===================== حالة التطبيق =====================
if "page" not in st.session_state:
    st.session_state.page = "home"

# ===================== القائمة الجانبية =====================
st.sidebar.title("القائمة")

menu = st.sidebar.radio("اختر القسم", [
    "تسجيل القضايا",
    "الحصر العام",
    "التنبيهات",
    "التقارير",
    "الأرشيف",
    "المكتبة القانونية",
    "البحث عن دعوى"
])

st.session_state.page = menu

# ===================== الصفحة الرئيسية =====================
if st.session_state.page == "home":
    st.markdown("## مرحباً بك في نظام إدارة القضايا")
    st.markdown("<div class='card'>نظام احترافي لإدارة القضايا والتقارير والتنبيهات</div>", unsafe_allow_html=True)

# ===================== فوتر =====================
st.markdown("""
<div class='footer'>
مع تحيات / وليد شعبان حماد - الإدارة العامة للقضايا
</div>
""", unsafe_allow_html=True)

# ================= نهاية الجزء 1 =================
