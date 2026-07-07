import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime

# إعداد الصفحة وتنسيقها (التصميم الفخم)
st.set_page_config(page_title="نظام إدارة القضايا القانونية", layout="wide")

def inject_custom_css():
    st.markdown("""
        <style>
        .main { background-color: #001f3f; color: white; }
        h1, h2, h3 { color: #FFD700 !important; text-align: center; font-family: 'Arial'; }
        .footer { 
            position: fixed; bottom: 0; width: 100%; text-align: center; 
            color: #FFD700; font-size: 14px; background: #001f3f; 
            padding: 10px; border-top: 2px solid #FFD700; z-index: 999;
        }
        .stDataFrame { border: 2px solid #FFD700; }
        </style>
    """, unsafe_allow_html=True)

inject_custom_css()

# تهيئة قاعدة البيانات (التنفيذ النهائي)
def init_database():
    conn = sqlite3.connect('legal_data.db')
    c = conn.cursor()
    # جدول القضايا الرئيسي
    c.execute('''CREATE TABLE IF NOT EXISTS cases (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    type TEXT, court TEXT, mamoriya TEXT, case_num TEXT,
                    year TEXT, circuit TEXT, kind TEXT, plaintiff TEXT,
                    defendant TEXT, subject TEXT, first_session DATE,
                    roll TEXT, procedure TEXT, notes TEXT, 
                    whatsapp_active BOOLEAN, whatsapp_num TEXT)''')
    conn.commit()
    conn.close()

init_database()

# الهيكل التنظيمي للبرنامج
st.markdown("<h1>⚖️ إدارة القضايا - منطقة البحيرة ⚖️</h1>", unsafe_allow_html=True)

menu = ["تسجيل القضايا", "الحصر العام", "التقارير", "الارشيف", "المكتبة القانونية"]
choice = st.sidebar.selectbox("القائمة الرئيسية", menu)

# التنفيذ حسب القسم
if choice == "تسجيل القضايا":
    with st.form("reg_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            c_type = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
            court = st.selectbox("المحكمة", ["الابتدائية", "الاستئناف", "النقض", "القضاء الإدارى", "الإدارية العليا"])
            c_name = st.text_input("اسم المحكمة")
        with col2:
            num = st.text_input("رقم القضية")
            year = st.text_input("السنة القضائية")
            circuit = st.text_input("الدائرة")
        
        plaintiff = st.text_input("اسم المدعى / الطاعن")
        defendant = st.text_input("اسم المدعى عليه / المطعون ضده")
        subject = st.text_area("موضوع الدعوى")
        
        submitted = st.form_submit_button("حفظ القضية")
        if submitted:
            conn = sqlite3.connect('legal_data.db')
            c = conn.cursor()
            c.execute("INSERT INTO cases (type, court, case_num, year, plaintiff, defendant, subject) VALUES (?,?,?,?,?,?,?)",
                      (c_type, court, num, year, plaintiff, defendant, subject))
            conn.commit()
            conn.close()
            st.success("تم الحفظ بنجاح")

elif choice == "الحصر العام":
    st.subheader("جدول الحصر العام")
    conn = sqlite3.connect('legal_data.db')
    df = pd.read_sql_query("SELECT * FROM cases", conn)
    conn.close()
    st.dataframe(df, use_container_width=True)

# الفوتر الثابت (يظهر دائماً)
st.markdown('<div class="footer">مع تحيات وليد حماد الادارة العامة للشءون القانونية ديوان عام منطقة البحيرة</div>', unsafe_allow_html=True)
