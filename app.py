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

# ============= CSS =============
st.markdown("""<style>@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');* { font-family: 'Cairo', sans-serif !important; }html, body { direction: rtl; color: #FFFFFF !important; }.stApp { background: linear-gradient(180deg, #0A1428 0%, #1E2A47 100%); }.marquee { background: linear-gradient(90deg, #D4AF37 0%, #FFD700 50%, #D4AF37 100%); color: #0A1428; padding: 12px; font-weight: 900; font-size: 16px; white-space: nowrap; overflow: hidden; border-radius: 0 0 15px 15px; }.marquee span { display: inline-block; animation: marquee 15s linear infinite; }@keyframes marquee { 0% { transform: translateX(-100%); } 100% { transform: translateX(100%); } }.main-title { color: #D4AF37; text-align: center; font-size: 36px; font-weight: 900; padding: 15px 0; }h1, h2, h3 { color: #D4AF37 !important; text-align: center !important; font-weight: 900; }div[data-testid="column"] { display: flex; justify-content: center; }[data-testid="stForm"] label, .stMarkdown { color: #FFFFFF !important; font-weight: 700; }.stButton > button { color: #000 !important; font-weight: 900 !important; font-size: 18px !important; border: none !important; border-radius: 15px !important; padding: 16px !important; width: 100% !important; max-width: 400px !important; margin: 10px auto !important; box-shadow: 0 4px 12px rgba(0,0,0,0.4) !important; display: block; }.btn-add button { background: linear-gradient(180deg, #4DA8DA 0%, #2C5282 100%) !important; }.btn-list button { background: linear-gradient(180deg, #4CAF50 0%, #2E7D32 100%) !important; }.btn-alert button { background: linear-gradient(180deg, #FF5252 0%, #D32F2F 100%) !important; animation: pulse 1.5s infinite; }.btn-report button { background: linear-gradient(180deg, #FF9800 0%, #F57C00 100%) !important; }.btn-lib button { background: linear-gradient(180deg, #3F51B5 0%, #303F9F 100%) !important; }.btn-arch button { background: linear-gradient(180deg, #9E9E9E 0%, #616161 100%) !important; }.btn-search button { background: linear-gradient(180deg, #9C27B0 0%, #6A1B9A 100%) !important; }.stTextInput > div > div > input, .stTextArea > div > div > textarea, .stSelectbox > div > div > select { background-color: #FFFFFF !important; color: #000 !important; border: 2px solid #D4AF37 !important; border-radius: 12px !important; padding: 12px !important; text-align: right !important; font-weight: 700 !important; }.stTextInput > div > label, .stSelectbox > div > label, .stTextArea > div > label { color: #FFFFFF !important; font-weight: 700 !important; font-size: 16px !important; }@keyframes pulse { 0% { box-shadow: 0 0 0 0 rgba(255, 82, 82, 0.7); } 70% { box-shadow: 0 0 0 10px rgba(255, 82, 82, 0); } 100% { box-shadow: 0 0 0 0 rgba(255, 82, 82, 0); } }</style>""", unsafe_allow_html=True)

st.markdown('<div class="marquee"><span>مع تحيات وليد حماد - الإدارة العامة للشئون القانونية بديوان عام منطقة البحيرة</span></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">⚖️ إدارة القضايا ⚖️</div>', unsafe_allow_html=True)

# ============= الداتا =============
DATA_FILE = "cases_data.json"; UPLOAD_FOLDER = "uploads"; TOKENS_FILE = "tokens.json"
SENDER_EMAIL = "" # <--- حط ايميلك
SENDER_PASSWORD = "" # <--- حط باسورد التطبيق
APP_URL = "https://qpyqpsmkqcvdou4imbfunp.streamlit.app/"
if not os.path.exists(UPLOAD_FOLDER): os.makedirs(UPLOAD_FOLDER)
if 'page' not in st.session_state: st.session_state.page = "الرئيسية"
if 'selected_case_id' not in st.session_state: st.session_state.selected_case_id = None

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
    verify_link = f"{APP_URL}?verify_token={token}"; subject = "تفعيل تنبيهات القضايا"
    body = f"مرحبا,\n\nفعل الاشتراك:\n{verify_link}\n\nالرابط صالح 24 ساعة."
    msg = MIMEMultipart(); msg["From"] = SENDER_EMAIL; msg["To"] = recipient_email; msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    try: server = smtplib.SMTP("smtp.gmail.com", 587); server.starttls(); server.login(SENDER_EMAIL, SENDER_PASSWORD); server.sendmail(SENDER_EMAIL, recipient_email, msg.as_string()); server.quit(); return True
    except Exception as e: st.error(f"خطأ: {e}"); return False
def verify_token(token):
    tokens_data = load_tokens(); now = datetime.now()
    for t in tokens_data["tokens"]:
        if t["token"] == token and datetime.strptime(t["expires"], "%Y-%m-%d %H:%M:%S") > now:
            if not t["verified"]: t["verified"] = True; save_tokens(tokens_data); return t["email"]
    return None
def send_case_alert_email(recipient_email, case, نوع_التنبيه):
    subject = f"تنبيه: {نوع_التنبيه} - قضية رقم {case['رقم']}"; body = f"الرقم: {case['رقم']} لسنة {case['سنة']}\nالمحكمة: {case['محكمة_اسم']}\nتاريخ الجلسة: {case['تاريخ_جلسة']}"
    msg = MIMEMultipart(); msg["From"] = SENDER_EMAIL; msg["To"] = recipient_email; msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    try: server = smtplib.SMTP("smtp.gmail.com", 587); server.starttls(); server.login(SENDER_EMAIL, SENDER_PASSWORD); server.sendmail(SENDER_EMAIL, recipient_email, msg.as_string()); server.quit(); return True
    except: return False

data = load_data()

# ============= 1. الرئيسية =============
if st.session_state.page == "الرئيسية":
    st.markdown('<h2>الأقسام</h2>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1: st.markdown('<div class="btn-add">', unsafe_allow_html=True); st.button("➕ تسجيل القضايا", on_click=lambda: st.session_state.update(page="تسجيل")); st.markdown('</div>', unsafe_allow_html=True)
    with col2: st.markdown('<div class="btn-list">', unsafe_allow_html=True); st.button("📋 الحصر العام", on_click=lambda: st.session_state.update(page="الحصر")); st.markdown('</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1,2,1])
    with col2: st.markdown('<div class="btn-alert">', unsafe_allow_html=True); st.button("🔴 مركز التنبيهات", on_click=lambda: st.session_state.update(page="التنبيهات")); st.markdown('</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1: st.markdown('<div class="btn-report">', unsafe_allow_html=True); st.button("📊 التقارير", on_click=lambda: st.session_state.update(page="التقارير")); st.markdown('</div>', unsafe_allow_html=True)
    with col2: st.markdown('<div class="btn-lib">', unsafe_allow_html=True); st.button("📚 المكتبة القانونية", on_click=lambda: st.session_state.update(page="المكتبة")); st.markdown('</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1,2,1])
    with col2: st.markdown('<div class="btn-arch">', unsafe_allow_html=True); st.button("🗄️ الأرشيف", on_click=lambda: st.session_state.update(page="الأرشيف")); st.markdown('</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1,2,1])
    with col2: st.markdown('<div class="btn-search">', unsafe_allow_html=True); st.button("🔍 البحث عن دعوى", on_click=lambda: st.session_state.update(page="البحث")); st.markdown('</div>', unsafe_allow_html=True)
    st.rerun()

# ============= 2. التسجيل =============
elif st.session_state.page == "تسجيل":
    st.markdown("<div style='height:2px; background:#D4AF37; margin:20px 0;'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#D4AF37; text-align:center'>➕ تسجيل قضية جديدة</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()
    ANWA3_MOSTANDAT = ["صحيفة دعوى", "صحيفة استئناف", "صحيفة طعن", "مذكرة دفاع", "حافظة مستندات", "تقرير خبير", "أخرى"]
    نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
    with st.form("form_case"):
        محكمة_اسم = st.text_input("اسم المحكمة"); مأمورية = st.text_input("المأمورية") if نوع == "استئناف" else ""
        col1, col2 = st.columns(2); رقم = col1.text_input("رقم الدعوى"); سنة = col2.text_input("السنة")
        دائرة = st.text_input("الدائرة")
        col1, col2 = st.columns(2); مدعي = col1.text_input("اسم المدعى"); مدعي_عليه = col2.text_input("اسم المدعى عليه")
        موضوع = st.text_area("موضوع الدعوى", height=100)
        col1, col2 = st.columns(2); تاريخ_جلسة = col1.date_input("تاريخ أول جلسة", value=datetime.now()); الرول = col2.text_input("الرول")
        سبب = st.text_input("سبب الجلسة"); ملاحظات = st.text_area("ملاحظات", height=80)
        تنبيه = st.checkbox("تفعيل التنبيهات"); واتس = st.text_input("رقم واتس اب") if تنبيه else ""
        col1, col2 = st.columns(2); مستند_نوع = col1.selectbox("نوع المستند", ANWA3_MOSTANDAT); مستند_ملف = col2.file_uploader("اختر الملف")
        submitted = st.form_submit_button("💾 حفظ القضية", use_container_width=True)
        if submitted:
            if not رقم or not سنة: st.error("❌ من فضلك ادخل رقم الدعوى والسنة")
            else:
                new_case = {"id": len(data["cases"])+1, "نوع": نوع, "محكمة_اسم": محكمة_اسم, "مأمورية": مأمورية, "رقم": رقم, "سنة": سنة, "دائرة": دائرة, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة), "الرول": الرول, "سبب": سبب, "ملاحظات": ملاحظات, "تنبيه": تنبيه, "واتس": واتس, "جلسات": [], "مستندات": [], "حالة": "متداولة"}
                if الرول or سبب: new_case["جلسات"].append({"تاريخ":str(تاريخ_جلسة),"الرول":الرول,"سبب":سبب,"ملاحظات":ملاحظات})
                if مستند_ملف: file_path = os.path.join(UPLOAD_FOLDER, f"{new_case['id']}_{مستند_ملف.name}"); open(file_path, "wb").write(مستند_ملف.getbuffer()); new_case["مستندات"].append({'نوع': مستند_نوع, 'اسم': مستند_ملف.name, 'مسار': file_path})
                data["cases"].append(new_case); save_data(data); st.success(f"✅ تم الحفظ"); st.session_state.page = "الحصر"; st.rerun()
    st.markdown("<div style='height:2px; background:#D4AF37; margin:20px 0;'></div>", unsafe_allow_html=True)

# ============= 3. الحصر =============
elif st.session_state.page == "الحصر":
    st.markdown("<div style='height:2px; background:#D4AF37; margin:20px 0;'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#D4AF37; text-align:center'>📋 الحصر العام الخارجي</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()
    if not data["cases"]: st.info("لا توجد قضايا مسجلة")
    else:
        for i, case in enumerate(data["cases"]): 
            if "id" not in case: case["id"] = i + 1
            if "مستندات" not in case: case["مستندات"] = []
        save_data(data); sorted_cases = sorted(data["cases"], key=lambda x: x.get("تاريخ_جلسة","9999"))
        st.success(f"📊 اجمالي القضايا: {len(sorted_cases)}")
        for idx, case in enumerate(sorted_cases, 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            st.markdown(f"<div style='border:2px solid #D4AF37; border-radius:12px; padding:10px; margin-bottom:10px; background:#1E2A47;'><b>{idx}. {رقم_كامل}</b><br>{case.get('مدعي','')} ضد {case.get('مدعي_عليه','')}<br>الجلسة: {case.get('تاريخ_جلسة','')}</div>", unsafe_allow_html=True)
            if st.button("فتح", key=f"open_{case['id']}"): st.session_state.selected_case_id = case['id']; st.session_state.page = "تفاصيل"; st.rerun()
    st.markdown("<div style='height:2px; background:#D4AF37; margin:20px 0;'></div>", unsafe_allow_html=True)

# ============= 4. التنبيهات =============
elif st.session_state.page == "التنبيهات":
    st.markdown("<div style='height:2px; background:#D4AF37; margin:20px 0;'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#D4AF37; text-align:center'>🔴 مركز التنبيهات</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()
    query_params = st.query_params
    if "verify_token" in query_params:
        email = verify_token(query_params["verify_token"])
        if email: st.success(f"✅ تم تفعيل {email}"); st.session_state['saved_email'] = email
        else: st.error("❌ الرابط منتهي")
        st.query_params.clear()
    tokens_data = load_tokens(); verified_emails = [t['email'] for t in tokens_data['tokens'] if t['verified']]
    user_email = st.text_input("البريد الالكتروني", value=st.session_state.get('saved_email',''))
    if st.button("ارسال رابط التفعيل", type="primary", use_container_width=True):
        if user_email: token = secrets.token_urlsafe(32); expires = datetime.now() + timedelta(days=1); tokens_data["tokens"].append({"email": user_email, "token": token, "expires": expires.strftime("%Y-%m-%d %H:%M:%S"), "verified": False}); save_tokens(tokens_data); send_verification_email(user_email, token); st.success("تم الارسال")
    st.info(f"المشتركين المفعلين: {len(verified_emails)}")
    st.markdown("<div style='height:2px; background:#D4AF37; margin:20px 0;'></div>", unsafe_allow_html=True)

# ============= 5. البحث =============
elif st.session_state.page == "البحث":
    st.markdown("<div style='height:2px; background:#D4AF37; margin:20px 0;'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#D4AF37; text-align:center'>🔍 البحث عن دعوى</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()
    col1, col2, col3 = st.columns(3)
    search_name = col1.text_input("اسم المدعى"); search_number = col2.text_input("رقم القضية"); search_year = col3.text_input("سنة القضية")
    if st.button("🔴 اضغط للبحث", use_container_width=True):
        results = [case for case in data["cases"] if (search_name and search_name.lower() in case.get("مدعي", "").lower()) or (search_number and search_year and str(search_number) == str(case.get("رقم", "")) and str(search_year) == str(case.get("سنة", "")))]
        if results:
            st.success(f"تم العثور على {len(results)} نتيجة")
            for r in results:
                with st.expander(f"⚖️ قضية رقم {r.get('رقم')} سنة {r.get('سنة')}"):
                    st.write(f"المدعى: {r.get('مدعي')}"); st.write(f"المدعى عليه: {r.get('مدعي_عليه')}")
                    if st.button("فتح", key=f"search_open_{r['id']}"): st.session_state.selected_case_id = r['id']; st.session_state.page = "تفاصيل"; st.rerun()
        else: st.warning("لم يتم العثور على نتائج")
    st.markdown("<div style='height:2px; background:#D4AF37; margin:20px 0;'></div>", unsafe_allow_html=True)

# ============= 6. التفاصيل =============
elif st.session_state.page == "تفاصيل":
    case = next((c for c in data["cases"] if c['id'] == st.session_state.selected_case_id), None)
    if not case: st.error("القضية غير موجودة"); st.session_state.page = "الحصر"; st.rerun()
    st.markdown(f"<h2 style='color:#D4AF37; text-align:center'>📄 تفاصيل القضية {case.get('رقم')} لسنة {case.get('سنة')}</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للحصر", use_container_width=True): st.session_state.page = "الحصر"; st.rerun()
    st.write(f"**المحكمة:** {case.get('محكمة_اسم')}"); st.write(f"**المدعي:** {case.get('مدعي')}"); st.write(f"**المدعى عليه:** {case.get('مدعي_عليه')}")
