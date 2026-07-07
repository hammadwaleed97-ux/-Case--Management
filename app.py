import streamlit as st
import sqlite3
from datetime import datetime

# 1. إعداد الصفحة وتصميمها (فخم، أزرق داكن، ذهبي)
st.set_page_config(page_title="نظام إدارة القضايا", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #001f3f; color: white; }
    .stApp { background-color: #001f3f; }
    h1, h2, h3 { color: #FFD700 !important; text-align: center; }
    .stButton>button { width: 100%; background-color: #FFD700; color: #001f3f; font-weight: bold; }
    .footer { 
        position: fixed; bottom: 0; width: 100%; text-align: center; 
        color: #FFD700; font-size: 14px; background: #001f3f; padding: 10px; border-top: 2px solid #FFD700;
    }
    </style>
    """, unsafe_allow_html=True)

# 2. إنشاء قاعدة البيانات (تنفيذ نهائي - حفظ دائم)
def init_db():
    conn = sqlite3.connect('legal_database.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS cases 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, type TEXT, court TEXT, 
                  mamoriya TEXT, case_num TEXT, year TEXT, circuit TEXT, 
                  case_kind TEXT, plaintiff TEXT, defendant TEXT, subject TEXT, 
                  first_session DATE, roll TEXT, procedure TEXT, notes TEXT, 
                  whatsapp_active BOOLEAN, whatsapp_num TEXT)''')
    conn.commit()
    conn.close()

init_db()

# 3. الهيكل الرئيسي للبرنامج
st.markdown("<h1>⚖️ إدارة القضايا ⚖️</h1>", unsafe_allow_html=True)

menu = ["تسجيل القضايا", "الحصر العام", "التقارير", "الاحصائيات", "التنبيهات", "الارشيف", "المكتبة القانونية", "البحث عن دعوى"]
choice = st.sidebar.selectbox("القائمة الرئيسية", menu)

# 4. منطق قسم تسجيل القضايا
if choice == "تسجيل القضايا":
    st.subheader("تسجيل القضايا")
    with st.form("reg_form"):
        col1, col2 = st.columns(2)
        with col1:
            c_type = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
            court = st.selectbox("المحكمة", ["الابتدائية", "الاستئناف", "النقض", "الإدارية", "القضاء الإدارى", "الإدارية العليا"])
            c_name = st.text_input("اسم المحكمة")
            if c_type == "استئناف": mamoriya = st.text_input("المأمورية")
            else: mamoriya = ""
        with col2:
            num = st.text_input("رقم الدعوى / الاستئناف / الطعن")
            year = st.text_input("السنة القضائية")
            circuit = st.text_input("الدائرة")
            case_kind = st.text_input("النوع")
        
        plaintiff = st.text_input("اسم المدعى / المستأنف / الطاعن")
        defendant = st.text_input("اسم المدعى عليه / المستأنف ضده / المطعون ضده")
        subject = st.text_area("موضوع الدعوى")
        
        col3, col4 = st.columns(2)
        with col3:
            first_session = st.date_input("تاريخ أول جلسة")
            roll = st.text_input("الرول")
        with col4:
            procedure = st.text_input("الإجراء المطلوب")
            notes = st.text_area("ملاحظات")
            
        wa_active = st.checkbox("تفعيل تنبيهات واتس اب")
        wa_num = st.text_input("رقم واتس اب")
        
        if st.form_submit_button("حفظ القضية"):
            conn = sqlite3.connect('legal_database.db')
            c = conn.cursor()
            c.execute("INSERT INTO cases (type, court, mamoriya, case_num, year, circuit, case_kind, plaintiff, defendant, subject, first_session, roll, procedure, notes, whatsapp_active, whatsapp_num) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                      (c_type, court, c_name, mamoriya, num, year, circuit, case_kind, plaintiff, defendant, subject, first_session, roll, procedure, notes, wa_active, wa_num))
            conn.commit()
            conn.close()
            st.success("تم الحفظ في الحصر العام بنجاح")

# 5. الفوتر الثابت
st.markdown('<div class="footer">مع تحيات / وليد شعبان حماد - الادارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>', unsafe_allow_html=True)
