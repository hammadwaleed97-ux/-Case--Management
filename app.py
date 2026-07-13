import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

st.set_page_config(page_title="المستشار القانوني", layout="wide", page_icon="⚖️")

if 'cases' not in st.session_state:
    st.session_state.cases = []

if 'notifications_enabled' not in st.session_state:
    st.session_state.notifications_enabled = False

SENDER_EMAIL = "your_email@gmail.com"
SENDER_PASSWORD = "your_app_password"

def send_email(to_email, subject, body):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except:
        return False

st.title("⚖️ نظام إدارة القضايا القانونية")

menu = st.sidebar.selectbox(
    "القائمة الرئيسية",
    ["الرئيسية", "تسجيل قضية جديدة", "الحصر العام", "مركز التنبيهات", "البحث عن دعوى"]
)

if menu == "الرئيسية":
    st.header("الرئيسية")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("إجمالي القضايا", len(st.session_state.cases))
    with col2:
        st.metric("القضايا النشطة", len([c for c in st.session_state.cases if c['الحالة'] == 'نشطة']))
    with col3:
        st.metric("جلسات هذا الأسبوع", 0)

elif menu == "تسجيل قضية جديدة":
    st.header("تسجيل قضية جديدة")
    with st.form("case_form"):
        col1, col2 = st.columns(2)
        with col1:
            case_number = st.text_input("رقم القضية")
            case_type = st.selectbox("نوع القضية", ["مدني", "جنائي", "تجاري", "عمالي", "أسرة"])
            client_name = st.text_input("اسم الموكل")
        with col2:
            court = st.text_input("المحكمة")
            next_session = st.date_input("تاريخ الجلسة القادمة")
            status = st.selectbox("الحالة", ["نشطة", "مؤجلة", "منتهية"])
        
        notes = st.text_area("ملاحظات")
        lawyer_email = st.text_input("ايميل المحامي للتنبيهات")
        
        submitted = st.form_submit_button("حفظ القضية")
        if submitted:
            new_case = {
                'رقم القضية': case_number,
                'النوع': case_type,
                'اسم الموكل': client_name,
                'المحكمة': court,
                'تاريخ الجلسة': next_session,
                'الحالة': status,
                'ملاحظات': notes,
                'الايميل': lawyer_email
            }
            st.session_state.cases.append(new_case)
            st.success("تم حفظ القضية بنجاح!")

elif menu == "الحصر العام":
    st.header("الحصر العام للقضايا")
    if st.session_state.cases:
        df = pd.DataFrame(st.session_state.cases)
        st.dataframe(df, use_container_width=True)
    else:
        st.info("لا توجد قضايا مسجلة بعد")

elif menu == "مركز التنبيهات":
    st.header("مركز التنبيهات")
    st.session_state.notifications_enabled = st.toggle("تفعيل التنبيهات بالايميل", st.session_state.notifications_enabled)
    
    if st.button("فحص الجلسات القريبة"):
        today = datetime.now().date()
        upcoming = [c for c in st.session_state.cases if c['تاريخ الجلسة'] <= today + timedelta(days=3)]
        if upcoming:
            for case in upcoming:
                st.warning(f"جلسة قريبة: قضية {case['رقم القضية']} بتاريخ {case['تاريخ الجلسة']}")
        else:
            st.success("لا توجد جلسات قريبة")

elif menu == "البحث عن دعوى":
    st.header("البحث عن دعوى")
    search_term = st.text_input("ابحث برقم القضية أو اسم الموكل")
    if search_term:
        results = [c for c in st.session_state.cases if search_term in c['رقم القضية'] or search_term in c['اسم الموكل']]
        if results:
            for case in results:
                with st.expander(f"قضية {case['رقم القضية']} - {case['اسم الموكل']}"):
                    st.write(case)
        else:
            st.error("لا توجد نتائج")
