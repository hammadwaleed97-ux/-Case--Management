# ============================================================
# ================== إدارة القضايا v5.7 =====================
# ========== الإدارة العامة للشئون القانونية البحيرة ==========
# ============================================================

import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

DATA_FILE = "cases_data.json"
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return {"cases": []}
def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=4)

data = load_data()
today = datetime.now().strftime("%A, %d %B %Y")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;800&display=swap');
    html, body {font-family: 'Cairo', sans-serif; direction: rtl;}
.stApp {background: linear-gradient(180deg, #0F1C2E 0%, #1A2F4F 100%);}
.stApp::before {
        content: "⚖️"; position: fixed; top: 50%; left: 50%;
        transform: translate(-50%, -50%); font-size: 400px;
        opacity: 0.04; z-index: 0; color: #C9A961;
    }
.marquee-container {
        background: linear-gradient(90deg, #C9A961, #D4B96A);
        padding: 10px; overflow: hidden; border-radius: 8px;
        margin-bottom: 15px; box-shadow: 0 0 10px rgba(201,169,97,0.4);
    }
.marquee-text {
        color: #0F1C2E; font-weight: 800; font-size: 14px;
        white-space: nowrap; display: inline-block;
        animation: scroll-rtl 18s linear infinite;
    }
    @keyframes scroll-rtl {
        0% {transform: translateX(-100%);}
        100% {transform: translateX(100%);}
    }
.header-calm {
        background: linear-gradient(135deg, #1A2F4F, #2C4A73);
        padding: 18px; border-radius: 12px; text-align: center;
        border: 2px solid #C9A961; margin-bottom: 20px;
    }
.header-calm h1 {color: #D4B96A; font-size: 30px; font-weight: 800; margin: 0;}
.header-calm p {color: #E8E8E8; font-size: 13px; font-weight: 600; margin: 6px 0 0 0;}
.section-title {color: #C9A961; text-align: center; font-size: 22px; font-weight: 800; margin: 15px 0;}
.section-divider {height: 4px; background: linear-gradient(90deg, transparent, #C9A961, transparent); margin: 15px 0;}

    label,.stTextInput label,.stSelectbox label,.stTextArea label,.stDateInput label {
        color: #D4B96A!important;
        font-weight: 700!important;
        font-size: 15px!important;
    }
    input, textarea, select {color: #0F1C2E!important; font-weight: 600; background-color: #FFFFFF!important;}

    div[data-testid="stButton"] > button {
        background: linear-gradient(135deg, #2C4A73, #3A5F8A)!important;
        color: #E8E8E8!important; width: 100%; padding: 20px 10px; border-radius: 12px;
        border: 2px solid #C9A961; font-weight: 700; font-size: 15px;
        height: 90px; margin: 5px 0;
    }
.btn-back {background: linear-gradient(135deg, #2C4A73, #3A5F8A)!important; border: 2px solid #C9A961!important; height: 55px!important;}

.card {
        background: rgba(26,47,79,0.85); padding: 18px; border-radius: 12px;
        border: 2px solid #C9A961; margin-bottom: 18px;
    }
.card-title {color: #D4B96A; font-weight: 800; font-size: 18px; margin-bottom: 15px; border-bottom: 2px solid #C9A961; padding-bottom: 10px;}

.table-container {overflow-x: auto; border: 3px solid #C9A961; border-radius: 10px; background: white;}
.case-table {
        width: 100%; border-collapse: collapse; background: #FFFFFF; color: #0F1C2E; font-size: 14px;
    }
.case-table th {
        background: linear-gradient(90deg, #D4B96A, #C9A961); color: #0F1C2E; padding: 12px 8px;
        text-align: center; font-weight: 800; border: 1px solid #C9A961;
    }
.case-table td {
        padding: 10px 8px; text-align: center; border: 1px solid #DDD; font-weight: 600; white-space: pre-line;
    }
.case-table tr.row1 {background: #FFF8E1;}
.case-table tr.row2 {background: #F0F4F8;}
.case-table tr.row-hey2a {background: #FFCDD2;} 
    h2, h3, h4, p {color: #FFFFFF!important;}
</style>
""", unsafe_allow_html=True)

if 'page' not in st.session_state: st.session_state.page = "الرئيسية"
if 'selected_case_id' not in st.session_state: st.session_state.selected_case_id = None

st.markdown("""
<div class='marquee-container'>
    <div class='marquee-text'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>
</div>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class='header-calm'>
    <div style='font-size:40px'>⚖️</div>
    <h1>إدارة القضايا</h1>
    <p>📅 {today}</p>
</div>
""", unsafe_allow_html=True)

st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)

# ==================== بداية قسم 1: الرئيسية ====================
if st.session_state.page == "الرئيسية":
    st.markdown("<div class='section-title'>الأقسام</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("تسجيل القضايا", use_container_width=True): st.session_state.page = "تسجيل"; st.rerun()
    with col2:
        if st.button("الحصر العام", use_container_width=True): st.session_state.page = "حصر"; st.rerun()
# ==================== نهاية قسم 1: الرئيسية ====================

# ==================== بداية قسم 2: التسجيل ====================
elif st.session_state.page == "تسجيل":
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#C9A961; text-align:center'>تسجيل القضايا</h2>", unsafe_allow_html=True)
    if st.button("العودة للرئيسية"): st.session_state.page = "الرئيسية"; st.rerun()

    نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])

    with st.form("form_case"):
        st.markdown("<div class='card'><div class='card-title'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            محكمة_نوع = st.selectbox("نوع المحكمة", ["الابتدائية", "الاستئناف", "النقض", "الإدارية", "القضاء الإدارى", "الإدارية العليا"])
        with col2:
            محكمة_اسم = st.text_input("اسم المحكمة")
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("<div class='card'><div class='card-title'>2- بيانات الدعوى</div>", unsafe_allow_html=True)
        رقم = st.text_input("رقم الدعوى")
        سنة = st.text_input("السنة")
        st.markdown("</div>", unsafe_allow_html=True)

        if st.form_submit_button("حفظ القضية", use_container_width=True):
            new_case = {"id": len(data["cases"])+1, "نوع": نوع, "محكمة_نوع": محكمة_نوع, "محكمة_اسم": محكمة_اسم, "رقم": رقم, "سنة": سنة, "جلسات": []}
            data["cases"].append(new_case)
            save_data(data)
            st.success("تم الحفظ")
            st.session_state.page = "الرئيسية"
            st.rerun()
# ==================== نهاية قسم 2: التسجيل ====================

# ==================== بداية قسم 3: الحصر العام ====================
elif st.session_state.page == "حصر":
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📊 الحصر العام</h2>", unsafe_allow_html=True)
    if st.button("العودة للرئيسية"): st.session_state.page = "الرئيسية"; st.rerun()

    if data["cases"]:
        st.markdown("<div class='table-container'>", unsafe_allow_html=True)
        table_html = "<table class='case-table'><tr><th>م</th><th>الرقم</th><th>المحكمة</th><th>فتح</th></tr>"
        for idx, case in enumerate(data["cases"], 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            محكمة = f"{case.get('محكمة_نوع','')} {case.get('محكمة_اسم','')}"
            table_html += f"<tr><td>{idx}</td><td>{رقم_كامل}</td><td>{محكمة}</td><td><button>فتح</button></td></tr>"
        table_html += "</table></div>"
        st.markdown(table_html, unsafe_allow_html=True)
    else:
        st.info("لا توجد قضايا")
# ==================== نهاية قسم 3: الحصر العام ====================
