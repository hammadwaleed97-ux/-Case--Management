# ================================================
# ============ الجزء الاول: الاساسيات ============
# ================================================
import streamlit as st
import pandas as pd
import json
import os
import smtplib
import secrets
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

# ============= التصميم النهائي المصلح =============
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');
    * { font-family: 'Cairo', sans-serif !important; }
    html, body { direction: rtl; color: #FFFFFF !important; }
    .stApp { background: linear-gradient(180deg, #0A1428 0%, #1E2A47 100%); }
    
    .marquee {
        background: linear-gradient(90deg, #D4AF37 0%, #FFD700 50%, #D4AF37 100%);
        color: #0A1428; padding: 12px; font-weight: 900; font-size: 16px;
        white-space: nowrap; overflow: hidden; border-radius: 0 0 15px 15px;
    }
    .marquee span { display: inline-block; animation: marquee 15s linear infinite; }
    @keyframes marquee { 0% { transform: translateX(-100%); } 100% { transform: translateX(100%); } }
    
    .main-title { color: #D4AF37; text-align: center; font-size: 36px; font-weight: 900; padding: 15px 0; }
    h1, h2, h3 { color: #D4AF37 !important; text-align: center !important; font-weight: 900; }
    
    div[data-testid="column"] { display: flex; justify-content: center; }
    [data-testid="stForm"] label, .stMarkdown { color: #FFFFFF !important; font-weight: 700; }
    
    .stButton > button {
        color: #000 !important; font-weight: 900 !important; font-size: 18px !important;
        border: none !important; border-radius: 15px !important; padding: 16px !important;
        width: 100% !important; max-width: 400px !important; margin: 10px auto !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.4) !important; display: block;
    }
    
    .btn-add button { background: linear-gradient(180deg, #4DA8DA 0%, #2C5282 100%) !important; }
    .btn-list button { background: linear-gradient(180deg, #4CAF50 0%, #2E7D32 100%) !important; }
    .btn-alert button { background: linear-gradient(180deg, #FF5252 0%, #D32F2F 100%) !important; animation: pulse 1.5s infinite; }
    .btn-report button { background: linear-gradient(180deg, #FF9800 0%, #F57C00 100%) !important; }
    .btn-lib button { background: linear-gradient(180deg, #3F51B5 0%, #303F9F 100%) !important; }
    .btn-arch button { background: linear-gradient(180deg, #9E9E9E 0%, #616161 100%) !important; }
    .btn-search button { background: linear-gradient(180deg, #9C27B0 0%, #6A1B9A 100%) !important; }
    
    .stTextInput > div > div > input, .stTextArea > div > div > textarea, .stSelectbox > div > div > select {
        background-color: #FFFFFF !important; color: #000 !important;
        border: 2px solid #D4AF37 !important; border-radius: 12px !important;
        padding: 12px !important; text-align: right !important; font-weight: 700 !important;
    }
    
    @keyframes pulse { 0% { box-shadow: 0 0 0 0 rgba(255, 82, 82, 0.7); } 70% { box-shadow: 0 0 0 10px rgba(255, 82, 82, 0); } 100% { box-shadow: 0 0 0 0 rgba(255, 82, 82, 0); } }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="marquee">
<span>مع تحيات وليد حماد - الإدارة العامة للشئون القانونية بديوان عام منطقة البحيرة بالهيئة القومية للتأمين الاجتماعي</span>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">⚖️ إدارة القضايا ⚖️</div>', unsafe_allow_html=True)

# ====== المتغيرات العامة ======
DATA_FILE = "cases_data.json"
UPLOAD_FOLDER = "uploads"
TOKENS_FILE = "tokens.json"
ANWA3_MOSTANDAT = ["صحيفة دعوى", "صحيفة استئناف", "صحيفة طعن", "مذكرة دفاع", "حافظة مستندات", "تقرير خبير", "تقرير طب شرعى", "تقرير لجنة طبية", "صحيفة تجديد من الشطب", "صحيفة تعجيل من الوقف", "صورة حكم تمهيدى", "أخرى"]

SENDER_EMAIL = "" # <--- حط ايميل الجيميل بتاعك هنا
SENDER_PASSWORD = "" # <--- حط باسورد التطبيق هنا
APP_URL = "https://qpyqpsmkqcvdou4imbfunp.streamlit.app/"

if not os.path.exists(UPLOAD_FOLDER): os.makedirs(UPLOAD_FOLDER)

if 'page' not in st.session_state: st.session_state.page = "الرئيسية"
if 'selected_case_id' not in st.session_state: st.session_state.selected_case_id = None

# ====== دوال التحميل والحفظ ======
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return {"cases": []}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=4)

def load_tokens():
    if os.path.exists(TOKENS_FILE):
        with open(TOKENS_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return {"tokens": []}

def save_tokens(tokens_data):
    with open(TOKENS_FILE, "w", encoding="utf-8") as f: json.dump(tokens_data, f, ensure_ascii=False, indent=4)

data = load_data()

# ========== الصفحة الرئيسية ==========
if st.session_state.page == "الرئيسية":
    st.markdown('<h2>الأقسام</h2>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="btn-add">', unsafe_allow_html=True)
        if st.button("➕ تسجيل القضايا", use_container_width=True): st.session_state.page = "تسجيل"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="btn-list">', unsafe_allow_html=True)
        if st.button("📋 الحصر العام", use_container_width=True): st.session_state.page = "حصر"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown('<div class="btn-alert">', unsafe_allow_html=True)
        if st.button("🔴 مركز التنبيهات", use_container_width=True): st.session_state.page = "تنبيهات"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="btn-report">', unsafe_allow_html=True)
        if st.button("📊 التقارير", use_container_width=True): st.session_state.page = "تقارير"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="btn-lib">', unsafe_allow_html=True)
        if st.button("📚 المكتبة القانونية", use_container_width=True): st.session_state.page = "مكتبة"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown('<div class="btn-arch">', unsafe_allow_html=True)
        if st.button("🗄️ الأرشيف", use_container_width=True): st.session_state.page = "ارشيف"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown('<div class="btn-search">', unsafe_allow_html=True)
        if st.button("🔍 البحث عن دعوى", use_container_width=True): st.session_state.page = "بحث"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# ================================================
# ========== نهاية الجزء الاول ==========
# ==============================================
