# ===== إدارة القضايا v5.39 =====================
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

# ============= دوال مساعدة =============

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"cases": []}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# ============ دوال التنبيهات =============
def render_notification_center():
    st.markdown("---")
    st.markdown("<h1 style='text-align: center; color: #D4AF37;'>📧 مركز التنبيهات</h1>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True):
        st.session_state.page = "الرئيسية"
        st.rerun()

    query_params = st.query_params
    if "verify_token" in query_params:
        email = verify_token(query_params["verify_token"])
        if email:
            st.success(f"✅ تم تفعيل الايميل {email} بنجاح. ستصلك التنبيهات الان")
            st.session_state['saved_email'] = email
        else:
            st.error("❌ الرابط غير صالح او منتهي")
        st.query_params.clear()

    st.markdown("<h3 style='color:#FFFFFF; text-align:center'>📊 ادارة التنبيهات</h3>", unsafe_allow_html=True)
    tokens_data = load_tokens()
    verified_emails = [t['email'] for t in tokens_data['tokens'] if t['verified']]

    with st.container(border=True):
        st.markdown("<div class='card-title'>تسجيل ايميل جديد للتنبيهات</div>", unsafe_allow_html=True)
        user_email = st.text_input("البريد الالكتروني", placeholder="example@domain.com", value=st.session_state.get('saved_email',''))
        if st.button("ارسال رابط التفعيل", type="primary", use_container_width=True):
            if user_email:
                token = secrets.token_urlsafe(32)
                expires = datetime.now() + timedelta(days=1)
                tokens_data["tokens"].append({"email": user_email, "token": token, "expires": expires.strftime("%Y-%m-%d %H:%M:%S"), "verified": False})
                save_tokens(tokens_data)
                if send_verification_email(user_email, token):
                    st.success("تم ارسال رابط التفعيل للايميل. من فضلك افتح الايميل وفعل الاشتراك")
                else:
                    st.error("فشل ارسال الايميل. راجع الايميل والباسورد الاحمر")
            else:
                st.warning("من فضلك ادخل الايميل")

    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("### 📅 الجلسات خلال 7 ايام القادمة")
    today = datetime.now().date()
    week_later = today + timedelta(days=7)
    data = load_data()

    # ========== ارسال تنبيهات تلقائية ==========
    if verified_emails and data["cases"]:
        for case in data["cases"]:

            # 1. تنبيهات الجلسات - زي ما هي
            if case.get('تاريخ_جلسة'):
                try:
                    session_date = datetime.strptime(case['تاريخ_جلسة'], '%Y-%m-%d').date()
                    days_to_session = (session_date - today).days
                    if days_to_session == 3: # قبل الجلسة ب 3 ايام
                        for email in verified_emails:
                            send_case_alert_email(email, case, "جلسة")
                except: pass

            # 2. تنبيهات الاحكام - الاضافة الجديدة
            if case.get('تاريخ_انتهاء_الطعن') and case.get('تاريخ_انتهاء_الطعن')!= '':
                try:
                    appeal_end_date = datetime.strptime(case['تاريخ_انتهاء_الطعن'], '%Y-%m-%d').date()
                    days_to_appeal = (appeal_end_date - today).days
                    # نبعت قبل انتهاء الطعن ب 15 يوم و ب 7 ايام و ب 3 ايام و يوم الانتهاء
                    if days_to_appeal in [15, 7, 3, 0]:
                        for email in verified_emails:
                            send_case_alert_email(email, case, "حكم")
                except: pass

    if not data["cases"]:
        st.warning("لا توجد قضايا مسجلة")
        return

    df = pd.DataFrame(data["cases"])
    df['تاريخ_جلسة'] = pd.to_datetime(df['تاريخ_جلسة'], errors='coerce').dt.date
    upcoming = df[(df['تاريخ_جلسة'] >= today) & (df['تاريخ_جلسة'] <= week_later)]
    st.info(f"عدد المشتركين المفعلين: {len(verified_emails)}")

    # عرض الجلسات - زي ما هي
    if not upcoming.empty:
        for idx, row in enumerate(upcoming.iterrows(), 1):
            case = row[1].to_dict()
            رقم_كامل = f"{case['رقم']} لسنة {case['سنة']}"
            محكمة_كاملة = f"{case['نوع']} {case['محكمة_اسم']}"
            if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = f"{case.get('دائرة','')} عمال" if case.get('دائرة','') else ""
            محكمة_كاملة += f"<br>{دائرة_كاملة}"
            خصوم = f"{case.get('مدعي','')}<br>ضد<br>{case.get('مدعي_عليه','')}"

            row_class = "row-judgment" if case.get('حالة') == 'منتهية' else "row1"

            st.markdown("<div class='table-container'>", unsafe_allow_html=True)
            table_html = f"<table class='case-table'><tr><th>م</th><th>الرقم والسنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>اخر جلسة</th><th>السبب</th><th>الحالة</th></tr>"
            table_html += f"<tr class='{row_class}'><td>{idx}</td><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td>{case['تاريخ_جلسة']}</td><td>{case.get('سبب','')}</td><td>{case.get('حالة','متداولة')}</td></tr></table></div>"
            st.markdown(table_html, unsafe_allow_html=True)

            # ========== زر الفتح ==========
            c1, c2, c3 = st.columns([4,1,4])
            with c2:
                if st.button("فتح", key=f"open_notif_{case['id']}"):
                    st.session_state.selected_case_id = case['id']
                    st.session_state.page = "تفاصيل"
                    st.rerun()
            st.markdown("<br>", unsafe_allow_html=True)
    else:
        st.info("مفيش جلسات خلال 7 ايام")

    # ========== عرض الاحكام الجديدة ==========
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("### ⚖️ الاحكام اللي الطعن بتاعها هيقفل خلال 15 يوم")

    if 'تاريخ_انتهاء_الطعن' in df.columns:
        df['تاريخ_انتهاء_الطعن'] = pd.to_datetime(df['تاريخ_انتهاء_الطعن'], errors='coerce').dt.date
        appeals_soon = df[(df['تاريخ_انتهاء_الطعن'] >= today) & (df['تاريخ_انتهاء_الطعن'] <= today + timedelta(days=15))]

        if not appeals_soon.empty:
            for idx, row in enumerate(appeals_soon.iterrows(), 1):
                case = row[1].to_dict()
                رقم_كامل = f"{case['رقم']} لسنة {case['سنة']}"
                متبقي = (case['تاريخ_انتهاء_الطعن'] - today).days
                لون = "red" if متبقي <= 3 else "orange" if متبقي <= 7 else "#D4AF37"
                الحكم_لصالح = case.get('الحكم_لصالح','')

                st.markdown(f"<div style='border:2px solid {لون}; padding:10px; border-radius:10px; margin-bottom:10px;'>", unsafe_allow_html=True)
                st.markdown(f"<h4 style='color:{لون};'> {idx}. {رقم_كامل} - الحكم {الحكم_لصالح}</h4>", unsafe_allow_html=True)
                st.markdown(f"**متبقي:** {متبقي} يوم | **تاريخ انتهاء الطعن:** {case['تاريخ_انتهاء_الطعن']}")
                st.markdown(f"**الخصوم:** {case.get('مدعي','')} ضد {case.get('مدعي_عليه','')}")

                c1, c2, c3 = st.columns([4,1,4])
                with c2:
                    if st.button("فتح", key=f"open_appeal_{case['id']}"):
                        st.session_state.selected_case_id = case['id']
                        st.session_state.page = "تفاصيل"
                        st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.info("مفيش احكام الطعن بتاعها هيقفل خلال 15 يوم")
# ============= حط بياناتك هنا بالاحمر فقط =============
SENDER_EMAIL = "hammadwaleed97@gmail.com" # <--- حط ايميل الجيميل بتاعك هنا
SENDER_PASSWORD = "r v y q q a y j o n w h u o x r" # <--- حط باسورد التطبيق هنا
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

# ==================================================================
# ===================================
# ==================================================================
# ================== بداية الجزء 1: الرئيسية والتسجيل ==================
# ==================================================================
if st.session_state.page == "الرئيسية":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>الأقسام</h2>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("تسجيل القضايا", use_container_width=True): st.session_state.page = "تسجيل"; st.rerun()
    with col2:
        if st.button("الحصر العام", use_container_width=True): st.session_state.page = "حصر"; st.rerun()
    if st.button("📧 مركز التنبيهات", type="secondary", use_container_width=True): st.session_state.page = "تنبيهات"; st.rerun()

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
# ================== نهاية الجزء 1: الرئيسية والتسجيل ==================
# =================================================================
# ==================================================================
# ================== بداية الجزء 2: الحصر والتفاصيل ==================
# ==================================================================

elif st.session_state.page == "حصر":
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📊 الحصر العام الخارجي</h2>", unsafe_allow_html=True)
    if st.button("العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()

    if not data["cases"]:
        st.info("لا توجد قضايا مسجلة")
    else:
        for i, case in enumerate(data["cases"]):
            if "id" not in case: case["id"] = i + 1
            if "مستندات" not in case: case["مستندات"] = []

        save_data(data)
        sorted_cases = sorted(data["cases"], key=lambda x: x.get("تاريخ_جلسة","9999"))
        total = len(sorted_cases)
        this_week = len([c for c in sorted_cases if c.get('تاريخ_جلسة') and datetime.strptime(c['تاريخ_جلسة'],'%Y-%m-%d') <= datetime.now() + timedelta(days=7)])
        st.info(f"📊 اجمالي القضايا: {total} | جلسات هذا الاسبوع: {this_week}")

        for idx, case in enumerate(sorted_cases, 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
            if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = f"{case.get('دائرة','')} عمال" if case.get('دائرة','') else ""
            محكمة_كاملة += f"<br>{دائرة_كاملة}"
            خصوم = f"{case.get('مدعي','')}<br>ضد<br>{case.get('مدعي_عليه','')}"

            if case.get('حالة') == 'منتهية': row_class = "row-judgment"
            elif "الهيئة" in str(case.get('مدعي','')): row_class = "row-hey2a"
            else: row_class = "row1" if idx % 2 == 1 else "row2"

            st.markdown("<div class='table-container'>", unsafe_allow_html=True)
            table_html = f"<table class='case-table'><tr><th>م</th><th>الرقم والسنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>اخر جلسة</th><th>السبب</th><th>الحالة</th></tr><tr class='{row_class}'><td>{idx}</td><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td>{case.get('تاريخ_جلسة','')}</td><td>{case.get('سبب','')}</td><td>{case.get('حالة','متداولة')}</td></tr></table></div>"
            st.markdown(table_html, unsafe_allow_html=True)

            c1, c2, c3 = st.columns([4,1,4])
            with c2:
                if st.button("فتح", key=f"open_{case['id']}"): st.session_state.selected_case_id = case['id']; st.session_state.page = "تفاصيل"; st.rerun()

elif st.session_state.page == "تنبيهات":
    render_notification_center()

elif st.session_state.page == "تفاصيل":
    case = next((c for c in data["cases"] if c['id'] == st.session_state.selected_case_id), None)
    if not case: st.error("القضية غير موجودة"); st.session_state.page = "حصر"; st.rerun()
    if 'جلسات' not in case: case['جلسات'] = []
    if 'مستندات' not in case: case['مستندات'] = []

    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown(f"<h2 style='color:#FFFFFF; text-align:center'>📄 تفاصيل القضية رقم {case.get('رقم')} لسنة {case.get('سنة')}</h2>", unsafe_allow_html=True)
    if st.button("العودة للحصر", use_container_width=True): st.session_state.page = "حصر"; st.rerun()

    st.markdown("<h3 style='color:#C9A961'>📌 بيانات القضية</h3>", unsafe_allow_html=True)
    table_html = f"<div style='background:linear-gradient(180deg, #0A1428 0%, #1E2A47 100%); padding:15px; border-radius:18px; border:2px solid #C9A961'><table style='width:100%; border-spacing:8px 8px'><tr><td style='background:linear-gradient(145deg, #1E3A6B 0%, #2C5282 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:center'><div style='font-size:12px; color:#FFD700'>رقم القضية</div><div style='font-size:20px; color:#FFFFFF'>{case.get('رقم')}</div></td><td style='background:linear-gradient(145deg, #2C5282 0%, #1E3A6B 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:center'><div style='font-size:12px; color:#FFD700'>السنة</div><div style='font-size:20px; color:#FFFFFF'>{case.get('سنة')}</div></td><td style='background:linear-gradient(145deg, #1E3A6B 0%, #2C5282 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:center'><div style='font-size:12px; color:#FFD700'>الدائرة</div><div style='font-size:18px; color:#FFFFFF'>{case.get('دائرة')} عمال</div></td><td style='background:linear-gradient(145deg, #2C5282 0%, #1E3A6B 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:center'><div style='font-size:12px; color:#FFD700'>النوع</div><div style='font-size:16px; color:#FFFFFF'>{case.get('نوع')}</div></td></tr><tr><td colspan='2' style='background:linear-gradient(145deg, #1E3A6B 0%, #2C5282 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:center'><div style='font-size:12px; color:#FFD700'>المحكمة</div><div style='font-size:15px; color:#FFFFFF'>{case.get('محكمة_اسم')}</div></td><td colspan='2' style='background:linear-gradient(145deg, #2C5282 0%, #1E3A6B 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:center'><div style='font-size:12px; color:#FFD700'>المأمورية</div><div style='font-size:15px; color:#FFFFFF'>{case.get('مأمورية') or '-'}</div></td></tr><tr><td colspan='2' style='background:linear-gradient(145deg, #FFF3CD 0%, #FFE69C 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:center'><div style='font-size:12px; color:#8B6914'>المدعي</div><div style='font-size:15px; color:#1E3A6B'>{case.get('مدعي')}</div></td><td colspan='2' style='background:linear-gradient(145deg, #CFF4FC 0%, #9EEAF9 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:center'><div style='font-size:12px; color:#055160'>المدعى عليه</div><div style='font-size:15px; color:#1E3A6B'>{case.get('مدعي_عليه')}</div></td></tr><tr><td colspan='4' style='background:linear-gradient(145deg, #1E3A6B 0%, #2C5282 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:center'><div style='font-size:12px; color:#FFD700'>الموضوع</div><div style='font-size:15px; color:#FFFFFF'>{case.get('موضوع')}</div></td></tr></table></div>"
    st.markdown(table_html, unsafe_allow_html=True)

    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h3 style='color:#C9A961'>📅 متابعة الجلسات</h3>", unsafe_allow_html=True)
    if len(case['جلسات']) > 0:
        جلسات_مرتبة = sorted(enumerate(case["جلسات"]), key=lambda x: x[1]['تاريخ'])
        html = "<table style='width:100%; border:3px solid #C9A961; background:#0A1428; border-radius:12px'><tr style='background:#C9A961; color:#000'><th>م</th><th>الرول</th><th>الجلسات</th><th>الإجراءات</th><th>ملاحظات</th><th>تحكم</th></tr>"
        for i, (idx, ج) in enumerate(جلسات_مرتبة, 1):
            لون = "#1E2A47" if i % 2 == 0 else "#142038"
            html += f"<tr style='background:{لون}; color:#FFF'><td>{i}</td><td style='color:#FFD700'>{ج.get('الرول','-')}</td><td>{ج['تاريخ']}</td><td>{ج.get('سبب','-')}</td><td>{ج.get('ملاحظات','-')}</td><td><button onclick=\"window.location.href='?edit={idx}'\">✏️ تعديل</button></td></tr>"
        html += "</table>"
        st.markdown(html, unsafe_allow_html=True)

    edit_idx = st.query_params.get("edit")
    with st.expander("➕ اضافة جلسة" if edit_idx is None else f"✏️ تعديل جلسة سابقة رقم {int(edit_idx)+1}"):
        with st.form("session_form"):
            t_val = datetime.strptime(case['جلسات'][int(edit_idx)]['تاريخ'], '%Y-%m-%d').date() if edit_idx else datetime.now().date()
            r_val = case['جلسات'][int(edit_idx)]['الرول'] if edit_idx else ""
            s_val = case['جلسات'][int(edit_idx)]['سبب'] if edit_idx else ""
            m_val = case['جلسات'][int(edit_idx)]['ملاحظات'] if edit_idx else ""
            c1,c2 = st.columns(2)
            t = c1.date_input("تاريخ الجلسة", value=t_val)
            r = c2.text_input("الرول", value=r_val)
            s = st.text_input("سبب التأجيل", value=s_val)
            m = st.text_area("ملاحظات", value=m_val)
            if st.form_submit_button("حفظ"):
                session_data = {'تاريخ':str(t),'الرول':r,'سبب':s,'ملاحظات':m}
                if edit_idx: case['جلسات'][int(edit_idx)] = session_data; st.query_params.clear()
                else: case['جلسات'].append(session_data)
                case['تاريخ_جلسة']=str(t); case['سبب']=s
                save_data(data); st.success("✅ تم الحفظ"); st.rerun()

    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h3 style='color:#C9A961'>📎 مستندات القضية</h3>", unsafe_allow_html=True)
    with st.form("upload_form"):
        نوع_المستند = st.selectbox("نوع المستند", ANWA3_MOSTANDAT)
        بيان_المستند = st.text_input("بيان المستند")
        uploaded_file = st.file_uploader("اختر الملف")
        if st.form_submit_button("حفظ المستند"):
            if uploaded_file:
                file_path = os.path.join(UPLOAD_FOLDER, f"{case['id']}_{uploaded_file.name}")
                with open(file_path, "wb") as f: f.write(uploaded_file.getbuffer())
                case['مستندات'].append({'نوع': نوع_المستند, 'بيان': بيان_المستند, 'اسم': uploaded_file.name, 'مسار': file_path})
                save_data(data); st.success("تم رفع المستند"); st.rerun()

    if case['مستندات']:
        for i, مستند in enumerate(case['مستندات']):
            if 'مسار' in مستند and os.path.exists(مستند['مسار']):
                c1, c2, c3, c4 = st.columns([2,2,2,1])
                with c1: st.write(f"{i+1}. {مستند.get('نوع','')}")
                with c2: st.write(مستند.get('بيان','-'))
                with c3: st.write(مستند.get('اسم',''))
                with c4:
                    with open(مستند['مسار'], "rb") as f: st.download_button("تحميل", f, file_name=مستند['اسم'], key=f"dl{i}")

    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h3 style='color:#FF6B6B'>⚖️ جلسة الحكم</h3>", unsafe_allow_html=True)
    if 'حكم' not in case:
        with st.form("judgment_form"):
            c1, c2 = st.columns(2)
            تاريخ_حكم = c1.date_input("تاريخ الحكم")
            مسندة_ل = c2.selectbox("مسندة لـ", ["الصالح", "الضد"])
            منطوق_الحكم = st.text_area("منطوق الحكم")
            if st.form_submit_button("حفظ الحكم واغلاق القضية"):
                case['حكم'] = {'تاريخ': str(تاريخ_حكم), 'المنطوق': منطوق_الحكم, 'مسندة': مسندة_ل}
                case['حالة'] = 'منتهية'
                case['جلسات'].append({'تاريخ':str(تاريخ_حكم),'الرول':'-','سبب':f'الحكم - مسندة لـ {مسندة_ل}','ملاحظات':منطوق_الحكم})
                save_data(data); st.success("✅ تم حفظ الحكم"); st.rerun()
    else:
        st.success(f"✅ تم الحكم بتاريخ: {case['حكم']['تاريخ']} - مسندة لـ: {case['حكم']['مسندة']}")
        st.info(f"المنطوق: {case['حكم']['المنطوق']}")

    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    if st.button("🗑️ حذف القضية نهائيا", type="primary"):
        data["cases"] = [c for c in data["cases"] if c['id']!= case['id']]
        save_data(data); st.success("تم الحذف")

# ==================================================================
# ================== نهاية الجزء 2: الحصر والتفاصيل ==================
# ==================================================================
# ===============================================================
