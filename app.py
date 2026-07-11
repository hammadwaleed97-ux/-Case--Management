# ================== إدارة القضايا v5.37 ====================
# ========== الإدارة العامة للشئون القانونية البحيرة ==========
# ============================================================
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

DATA_FILE = "cases_data.json"
UPLOAD_FOLDER = "uploads"
TOKENS_FILE = "tokens.json"
if not os.path.exists(UPLOAD_FOLDER): os.makedirs(UPLOAD_FOLDER)

# ============= حط بياناتك هنا بالاحمر فقط =============
SENDER_EMAIL = "hammadwaleed97@gmail.com" # <--- احمر
SENDER_PASSWORD = "r v y q q a y j o n w h u o x r" # <--- احمر
APP_URL = "https://qpyqpsmkqcvdou4imbfunp.streamlit.app/" # ده بتاعك
# ==================================================
# ==================================================================
# ================== بداية الجزء 1: الدوال والتسجيل ==================
# ==================================================================

ANWA3_MOSTANDAT = ["صحيفة دعوى", "صحيفة استئناف", "صحيفة طعن", "مذكرة دفاع", "حافظة مستندات", "تقرير خبير", "تقرير طب شرعى", "تقرير لجنة طبية", "صحيفة تجديد من الشطب", "صحيفة تعجيل من الوقف", "صورة حكم تمهيدى", "أخرى"]

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

def send_verification_email(recipient_email, token):
    verify_link = f"{APP_URL}?verify_token={token}"
    subject = "تفعيل تنبيهات القضايا - الشئون القانونية البحيرة"
    body = f"مرحبا,\n\nلقد قمت بالتسجيل لتلقي تنبيهات الجلسات.\nمن فضلك فعل الاشتراك بالضغط على الرابط التالي:\n{verify_link}\n\nالرابط صالح لمدة 24 ساعة."
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = recipient_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, recipient_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"خطأ في ارسال الايميل: {e}")
        return False

def verify_token(token):
    tokens_data = load_tokens()
    now = datetime.now()
    for t in tokens_data["tokens"]:
        if t["token"] == token and datetime.strptime(t["expires"], "%Y-%m-%d %H:%M:%S") > now:
            if not t["verified"]:
                t["verified"] = True
                save_tokens(tokens_data)
                return t["email"]
    return None

data = load_data()
today = datetime.now().strftime("%A, %d %B %Y")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;800&display=swap');
html, body {font-family: 'Cairo', sans-serif; direction: rtl;}
.stApp {background: linear-gradient(180deg, #0F1C2E 0%, #1A2F4F 100%);}
.marquee-container {background: linear-gradient(90deg, #C9A961, #D4B96A); padding: 10px; border-radius: 8px; margin-bottom: 15px;}
.marquee-text {color: #0F1C2E; font-weight: 800; font-size: 14px; white-space: nowrap; display: inline-block; animation: scroll-rtl 18s linear infinite;}
@keyframes scroll-rtl {0% {transform: translateX(-100%);} 100% {transform: translateX(100%);}}
.header-calm {background: linear-gradient(135deg, #1A2F4F, #2C4A73); padding: 18px; border-radius: 12px; text-align: center; border: 2px solid #C9A961; margin-bottom: 20px;}
.header-calm h1 {color: #D4B96A; font-size: 30px; font-weight: 800; margin: 0;}
.header-calm p {color: #E8E8E8; font-size: 13px; font-weight: 600; margin: 6px 0 0 0;}
.section-divider {height: 4px; background: linear-gradient(90deg, transparent, #C9A961, transparent); margin: 15px 0;}
label,.stTextInput label,.stSelectbox label,.stTextArea label,.stDateInput label {color: #D4B96A!important; font-weight: 700!important; font-size: 15px!important; text-align: right;}
input, textarea, select {color: #0F1C2E!important; font-weight: 600; background-color: #FFFFFF!important;}
div[data-testid="stButton"] > button {background: linear-gradient(135deg, #2C4A73, #3A5F8A)!important; color: #E8E8E8!important; width: 100%; border-radius: 12px; border: 2px solid #C9A961; font-weight: 700; font-size: 15px; height: 90px;}
.btn-back button {height: 55px!important;}
.btn-save button {background: linear-gradient(135deg, #C9A961, #D4B96A)!important; color: #0F1C2E!important; height: 50px!important; font-weight:800!important}
.card {background: rgba(26,47,79,0.85); padding: 18px; border-radius: 12px; border: 2px solid #C9A961; margin-bottom: 18px;}
.card-title {color: #D4B96A; font-weight: 800; font-size: 18px; margin-bottom: 15px; border-bottom: 2px solid #C9A961; padding-bottom: 10px; text-align: right;}
.table-container {overflow-x: auto; border: 3px solid #C9A961; border-radius: 10px; background: white; margin-bottom:15px}
.case-table {width: 100%; border-collapse: collapse; background: #FFFFFF; color: #0F1C2E; font-size: 14px;}
.case-table th {background: linear-gradient(90deg, #D4B96A, #C9A961); color: #0F1C2E; padding: 12px 8px; text-align: center; font-weight: 800;}
.case-table td {padding: 10px 8px; text-align: center; border: 1px solid #DDD; font-weight: 600;}
.row1 {background:#F9F9F9}.row2 {background:#FFFFFF}.row-hey2a {background:#FFE5E5; font-weight:800}.row-judgment {background:#FFDCDC; font-weight:800}
h2, h3, h4, p {color: #FFFFFF!important;}
div[data-testid="stButton"] > button[kind="secondary"] {background: linear-gradient(135deg, #8B0000, #A52A2A)!important; color: #FFD700!important; border: 3px solid #FFD700!important; font-size: 18px!important; font-weight: 800!important; box-shadow: 0 0 15px #8B0000;}
</style>
""", unsafe_allow_html=True)

if 'page' not in st.session_state: st.session_state.page = "الرئيسية"
if 'selected_case_id' not in st.session_state: st.session_state.selected_case_id = None

st.markdown("<div class='marquee-container'><div class='marquee-text'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div></div>", unsafe_allow_html=True)
st.markdown(f"<div class='header-calm'><div style='font-size:40px'>⚖️</div><h1>إدارة القضايا</h1><p>📅 {today}</p></div>", unsafe_allow_html=True)
st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)

if st.session_state.page == "الرئيسية":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>الأقسام</h2>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("تسجيل القضايا", use_container_width=True): st.session_state.page = "تسجيل"; st.rerun()
    with col2:
        if st.button("الحصر العام", use_container_width=True): st.session_state.page = "حصر"; st.rerun()
    if st.button("📧 مركز التنبيهات", type="secondary", use_container_width=True): st.session_state.page = "التنبيهات"; st.rerun()

elif st.session_state.page == "تسجيل":
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#C9A961; text-align:center'>تسجيل القضايا</h2>", unsafe_allow_html=True)
    if st.button("العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()
    نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
    with st.form("form_case"):
        st.markdown("<div class='card'><div class='card-title'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        محكمة_اسم = st.text_input("اسم المحكمة")
        مأمورية = st.text_input("المأمورية") if نوع == "استئناف" else ""
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<div class='card'><div class='card-title'>2- بيانات الدعوى</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: رقم = st.text_input("رقم الدعوى / الاستئناف / الطعن")
        with col2: سنة = st.text_input("السنة القضائية")
        دائرة = st.text_input("الدائرة")
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<div class='card'><div class='card-title'>3- بيانات الخصوم</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: مدعي = st.text_input("اسم المدعى / المستأنف / الطاعن")
        with col2: مدعي_عليه = st.text_input("اسم المدعى عليه / المستأنف ضده / المطعون ضده")
        موضوع = st.text_area("موضوع الدعوى")
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<div class='card'><div class='card-title'>4- بيانات الجلسة</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: تاريخ_جلسة = st.date_input("تاريخ أول جلسة", value=datetime.now())
        with col2: الرول = st.text_input("الرول")
        سبب = st.text_input("سبب الجلسة")
        ملاحظات = st.text_area("ملاحظات")
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<div class='card'><div class='card-title'>5- التنبيهات والمستندات</div>", unsafe_allow_html=True)
        تنبيه = st.checkbox("تفعيل التنبيهات عبر الواتس اب")
        واتس = st.text_input("رقم هاتف واتس اب") if تنبيه else ""
        col1, col2 = st.columns(2)
        with col1: مستند_نوع = st.selectbox("نوع المستند", ANWA3_MOSTANDAT)
        with col2: مستند_ملف = st.file_uploader("اختر الملف")
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<div class='btn-save'>", unsafe_allow_html=True)
        if st.form_submit_button("حفظ القضية", use_container_width=True):
            if not رقم or not سنة: st.error("❌ من فضلك ادخل رقم الدعوى والسنة")
            else:
                new_case = {"id": len(data["cases"])+1, "نوع": نوع, "محكمة_اسم": محكمة_اسم, "مأمورية": مأمورية, "رقم": رقم, "سنة": سنة, "دائرة": دائرة, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة), "الرول": الرول, "سبب": سبب, "ملاحظات": ملاحظات, "تنبيه": تنبيه, "واتس": واتس, "جلسات": [], "مستندات": [], "حالة": "متداولة"}
                if الرول or سبب: new_case["جلسات"].append({"تاريخ":str(تاريخ_جلسة),"الرول":الرول,"سبب":سبب,"ملاحظات":ملاحظات})
                if مستند_ملف:
                    file_path = os.path.join(UPLOAD_FOLDER, f"{new_case['id']}_{مستند_ملف.name}")
                    with open(file_path, "wb") as f: f.write(مستند_ملف.getbuffer())
                    new_case["مستندات"].append({'نوع': مستند_نوع, 'اسم': مستند_ملف.name, 'مسار': file_path})
                data["cases"].append(new_case); save_data(data)
                st.success(f"✅ تم حفظ القضية رقم {رقم} لسنة {سنة}")
                st.session_state.page = "حصر"; st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
# ==================================================================
# ================== نهاية الجزء 1: الدوال والتسجيل ==================
# ==================================================================
