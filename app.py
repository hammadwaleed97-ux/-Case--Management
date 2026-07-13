# =============================================
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
    
    .case-table { width:100%; color:#FFFFFF; text-align:center; border-collapse: collapse; }
    .case-table th { background:#D4AF37; color:#0A1428; padding:8px; font-weight:900; }
    .case-table td { padding:8px; border-bottom: 1px solid #D4AF37; }
    .table-container { background:#1E2A47; padding:10px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px; }
    .row1 { background: #142038; }
    .row2 { background: #1E2A47; }
    .row-hey2a { background: #1E3A6B; }
    .row-judgment { background: #2C2F33; }
    
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

SENDER_EMAIL = "" 
SENDER_PASSWORD = "" 
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
    with open(TOKENS_FILE, "w", encoding="utf-8") as f:
        json.dump(tokens_data, f, ensure_ascii=False, indent=4)
        
# ========= دوال التنبيهات =========
def get_alert_cases():
    data = load_data()
    today = datetime.now().date()
    all_cases = data["cases"]
    alerts = {"sessions": [], "appeals": []}
    for case in all_cases:
        if case.get('حالة') == 'متداولة' and case.get('تاريخ_جلسة'):
            try:
                session_date = datetime.strptime(case['تاريخ_جلسة'], '%Y-%m-%d').date()
                days_left = (session_date - today).days
                if 0 <= days_left <= 7:
                    case['days_left'] = days_left
                    alerts["sessions"].append(case)
            except: pass
        if case.get('حالة') == 'منتهية' and case.get('مسندة_ل_الحكم') == 'الضد' and case.get('تاريخ_الحكم'):
            try:
                judgment_date = datetime.strptime(case['تاريخ_الحكم'], '%Y-%m-%d').date()
                appeal_days = 40 if case['نوع'] == 'دعوى' else 60
                notify_on = judgment_date + timedelta(days=appeal_days - 15)
                days_left = (notify_on - today).days
                if days_left == 0:
                    case['appeal_days'] = appeal_days
                    alerts["appeals"].append(case)
            except: pass
    return alerts
# =
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
        if st.button("📋 الحصر العام", use_container_width=True): st.session_state.page = "الحصر"; st.rerun()
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
# =======================================
# ================================================
# =========== الجزء الثاني: تسجيل القضايا ============
# ================================================
elif st.session_state.page == "تسجيل":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#D4AF37; text-align:center'>➕ تسجيل القضايا</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", key="back_add", use_container_width=True):
        st.session_state.page = "الرئيسية"
        st.rerun()

    نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"], key="case_type_add")
    with st.form("form_case_add"):
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        محكمة_اسم = st.text_input("اسم المحكمة", key="court_name_add"); مأمورية = st.text_input("المأمورية", key="mamoria_add") if نوع == "استئناف" else ""
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>2- بيانات الدعوى</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: رقم = st.text_input("رقم الدعوى / الاستئناف / الطعن", key="case_num_add")
        with col2: سنة = st.text_input("السنة القضائية", key="case_year_add")
        دائرة = st.text_input("الدائرة", key="circle_add"); st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>3- بيانات الخصوم</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: مدعي = st.text_input("اسم المدعى / المستأنف / الطاعن", key="plaintiff_add")
        with col2: مدعي_عليه = st.text_input("اسم المدعى عليه / المستأنف ضده / المطعون ضده", key="defendant_add")
        موضوع = st.text_area("موضوع الدعوى", height=100, key="subject_add"); st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>4- بيانات الجلسة</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: تاريخ_جلسة = st.date_input("تاريخ أول جلسة", value=datetime.now(), key="session_date_add")
        with col2: الرول = st.text_input("الرول", key="roll_add")
        سبب = st.text_input("سبب الجلسة", key="reason_add"); ملاحظات = st.text_area("ملاحظات", height=100, key="notes_add"); st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>5- صحيفة الدعوى</div>", unsafe_allow_html=True)
        مسندة_ل = st.selectbox("نوع الصحيفة", ["صحيفة الدعوى"], key="paper_type_add"); st.markdown("</div>", unsafe_allow_html=True)

        if st.form_submit_button("💾 حفظ القضية", use_container_width=True, type="primary"):
            if not رقم or not سنة: st.error("❌ من فضلك ادخل رقم الدعوى والسنة")
            else:
                new_case = {"id": len(data["cases"])+1, "نوع": نوع, "محكمة_اسم": محكمة_اسم, "مأمورية": مأمورية, "رقم": رقم, "سنة": سنة, "دائرة": دائرة, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة), "الرول": الرول, "سبب": سبب, "ملاحظات": ملاحظات, "جلسات": [], "مستندات": [], "حالة": "متداولة", "مسندة_ل": مسندة_ل}
                if الرول or سبب: new_case["جلسات"].append({"تاريخ":str(تاريخ_جلسة),"الرول":الرول,"سبب":سبب,"ملاحظات":ملاحظات})
                data["cases"].append(new_case); save_data(data)
                st.success(f"✅ تم حفظ القضية رقم {رقم} لسنة {سنة}")
                st.session_state.page = "الحصر"; st.rerun()
# ================================================
# ========== الجزء الثالث: الحصر العام ============
# ================================================
elif st.session_state.page == "الحصر":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📊 الحصر العام الخارجي</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()

    # ======= السطرين الجداد دول بس =======
    if st.session_state.get('open_from_search', False):
        st.session_state.open_from_search = False
        st.info("جاري فتح القضية من البحث...")
    # ======================================

    # هنا باقي كود الحصر بتاعك زي ما هو
    # ... عرض الجدول بتاع الحصر
    if not data["cases"]:
        st.info("لا توجد قضايا مسجلة")
    else:
        for i, case in enumerate(data["cases"]):
            if "id" not in case: case["id"] = i + 1
            if "مستندات" not in case: case["مستندات"] = []

        save_data(data)
        sorted_cases = sorted(data["cases"], key=lambda x: x.get("تاريخ_جلسة","9999-12-31"))
        total = len(sorted_cases)
        today = datetime.now().date()
        this_week = len([c for c in sorted_cases if c.get('تاريخ_جلسة') and datetime.strptime(c['تاريخ_جلسة'],'%Y-%m-%d').date() <= today + timedelta(days=7)])
        ended = len([c for c in sorted_cases if c.get('حالة') == 'منتهية'])
        
        st.markdown(f"<div style='background:#1E2A47; padding:20px; border-radius:15px; border:2px solid #D4AF37; text-align:center; margin-bottom:20px'>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1: st.markdown(f"<div style='font-size:28px; font-weight:900; color:#D4AF37'>📊 {total}</div><div style='font-size:18px; color:#FFF; font-weight:700'>اجمالي القضايا</div>", unsafe_allow_html=True)
        with col2: st.markdown(f"<div style='font-size:28px; font-weight:900; color:#4DA8DA'>📅 {this_week}</div><div style='font-size:18px; color:#FFF; font-weight:700'>جلسات هذا الاسبوع</div>", unsafe_allow_html=True)
        with col3: st.markdown(f"<div style='font-size:28px; font-weight:900; color:#FF5252'>🚫 {ended}</div><div style='font-size:18px; color:#FFF; font-weight:700'>عدد المنتهية</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        for idx, case in enumerate(sorted_cases, 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
            if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = f"{case.get('دائرة','')} عمال" if case.get('دائرة','') else "" # <-- مرة واحدة بس
            if دائرة_كاملة: محكمة_كاملة += f"<br>{دائرة_كاملة}"
            خصوم = f"<div style='background:#FFF3CD; padding:8px; border-radius:8px; color:#000; margin-bottom:5px; text-align:center'><b>المدعى:</b><br>{case.get('مدعي','')}</div><div style='background:#CFF4FC; padding:8px; border-radius:8px; color:#000; text-align:center'><b>المدعى عليه:</b><br>{case.get('مدعي_عليه','')}</div>"

            if case.get('حالة') == 'منتهية': row_class = "row-judgment"
            elif "الهيئة" in str(case.get('مدعي','')): row_class = "row-hey2a"
            else: row_class = "row1" if idx % 2 == 1 else "row2"

            st.markdown("<div class='table-container'>", unsafe_allow_html=True)
            table_html = f"<table class='case-table'><tr><th>م</th><th>الرقم والسنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>اخر جلسة</th><th>السبب</th><th>الحالة</th></tr><tr class='{row_class}'><td>{idx}</td><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td style='color:#FFD700; font-weight:900'>{case.get('تاريخ_جلسة','')}</td><td>{case.get('سبب','')}</td><td style='color:#4CAF50; font-weight:900'>{case.get('حالة','متداولة')}</td></tr></table></div>"
            st.markdown(table_html, unsafe_allow_html=True)

            c1, c2, c3 = st.columns([4,1,4])
            with c2:
                if st.button("فتح", key=f"open_{case['id']}", use_container_width=True): 
                    st.session_state.selected_case_id = case['id']; st.session_state.page = "تفاصيل"; st.rerun()

# ================================================
# ============ الجزء الرابع: تفاصيل القضية ============
# ================================================
elif st.session_state.page == "تفاصيل":
    data = load_data()
    case = next((c for c in data["cases"] if c["id"] == st.session_state.selected_case_id), None)
    if not case: st.error("القضية غير موجودة"); st.session_state.page = "الحصر"; st.rerun()
    if 'جلسات' not in case: case['جلسات'] = []
    if 'مستندات' not in case: case['مستندات'] = []

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown(f"<h2 style='color:#D4AF37; text-align:center'>📄 تفاصيل القضية رقم {case.get('رقم')} لسنة {case.get('سنة')}</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للحصر العام", use_container_width=True): st.session_state.page = "الحصر"; st.rerun()

    # 1- بيانات القضية كروت 3 في سطر
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:15px'>1- بيانات القضية</div>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1: 
        st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>رقم القضية</div><div style='color:#FFF; font-weight:900; font-size:22px'>{case.get('رقم')}</div></div>", unsafe_allow_html=True)
    with col2: 
        st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>السنة</div><div style='color:#FFF; font-weight:900; font-size:22px'>{case.get('سنة')}</div></div>", unsafe_allow_html=True)
    with col3: 
        دائرة_نص = f"{case.get('دائرة')} عمال" if case.get('دائرة') else "" # <-- مرة واحدة بس
        st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>الدائرة</div><div style='color:#FFF; font-weight:900; font-size:18px'>{دائرة_نص}</div></div>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1: 
        st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>النوع</div><div style='color:#FFF; font-weight:900; font-size:18px'>{case.get('نوع')}</div></div>", unsafe_allow_html=True)
    with col2: 
        محكمة_كاملة = f"{case.get('محكمة_اسم')}"
        if case.get('مأمورية'): محكمة_كاملة += f" - مأمورية {case.get('مأمورية')}"
        st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>المحكمة</div><div style='color:#FFF; font-weight:700; font-size:14px'>{محكمة_كاملة}</div></div>", unsafe_allow_html=True)
    with col3: 
        st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>الحالة</div><div style='color:#4CAF50; font-weight:900; font-size:18px'>{case.get('حالة')}</div></div>", unsafe_allow_html=True)
    
    st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>الموضوع</div><div style='color:#FFF; font-weight:700; font-size:16px'>{case.get('موضوع')}</div></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # 2- بيانات الخصوم كارت
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>2- بيانات الخصوم</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1: st.markdown(f"<div style='background:#FFF3CD; padding:10px; border-radius:10px; color:#000; text-align:center'><b>المدعى:</b><br>{case.get('مدعي')}</div>", unsafe_allow_html=True)
    with col2: st.markdown(f"<div style='background:#CFF4FC; padding:10px; border-radius:10px; color:#000; text-align:center'><b>المدعى عليه:</b><br>{case.get('مدعي_عليه')}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # 3- متابعة الجلسات
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>3- متابعة الجلسات</div>", unsafe_allow_html=True)
    if case.get("جلسات"):
        html = "<table style='width:100%; border:2px solid #D4AF37; background:#0A1428; border-radius:12px'><tr style='background:#D4AF37; color:#000'><th>م</th><th>التاريخ</th><th>الرول</th><th>السبب</th><th>ملاحظات</th></tr>"
        for i, ج in enumerate(case["جلسات"], 1):
            لون = "#1E2A47" if i % 2 == 0 else "#142038"
            html += f"<tr style='background:{لون}; color:#FFF'><td>{i}</td><td>{ج.get('تاريخ')}</td><td>{ج.get('الرول')}</td><td>{ج.get('سبب')}</td><td>{ج.get('ملاحظات')}</td></tr>"
        html += "</table>"; st.markdown(html, unsafe_allow_html=True)
    else: st.info("لا توجد جلسات مسجلة")
    
    with st.expander("➕ اضافة جلسة جديدة"): # <-- ظهرتها
        with st.form("add_session"):
            تاريخ_جديد = st.date_input("تاريخ الجلسة", value=datetime.now())
            رول_جديد = st.text_input("الرول"); سبب_جديد = st.text_input("سبب الجلسة"); ملاحظات_جديدة = st.text_area("ملاحظات")
            if st.form_submit_button("حفظ الجلسة"):
                case["جلسات"].append({"تاريخ":str(تاريخ_جديد),"الرول":رول_جديد,"سبب":سبب_جديد,"ملاحظات":ملاحظات_جديدة})
                case["تاريخ_جلسة"] = str(تاريخ_جديد); save_data(data); st.success("تم اضافة الجلسة"); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    # 4- المستندات
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>4- المستندات</div>", unsafe_allow_html=True)
    with st.form("upload_form"):
        نوع_المستند = st.selectbox("نوع المستند", ANWA3_MOSTANDAT)
        uploaded_file = st.file_uploader("اختر الملف")
        if st.form_submit_button("رفع المستند"):
            if uploaded_file:
                file_path = os.path.join(UPLOAD_FOLDER, f"{case['id']}_{uploaded_file.name}")
    st.markdown("</div>", unsafe_allow_html=True)

    # 5- جلسة الحكم
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #FF5252; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#FF5252; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>5- جلسة الحكم</div>", unsafe_allow_html=True)
    
    if case.get('حالة') != 'منتهية':
        with st.form("judgment_form"):
            st.markdown("<div style='background:#142038; padding:10px; border-radius:10px; margin-bottom:10px'>", unsafe_allow_html=True)
            st.markdown("<label style='color:#FFD700; font-weight:900; font-size:16px'>1- تاريخ الجلسة</label>", unsafe_allow_html=True)
            تاريخ_حكم = st.date_input("تاريخ الجلسة", value=datetime.now().date(), label_visibility="collapsed")
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<div style='background:#142038; padding:10px; border-radius:10px; margin-bottom:10px'>", unsafe_allow_html=True)
            st.markdown("<label style='color:#FFD700; font-weight:900; font-size:16px'>2- منطوق الحكم</label>", unsafe_allow_html=True)
            منطوق_الحكم = st.text_area("منطوق الحكم", height=150, placeholder="اكتب منطوق الحكم هنا...", label_visibility="collapsed")
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<div style='background:#142038; padding:10px; border-radius:10px; margin-bottom:10px'>", unsafe_allow_html=True)
            st.markdown("<label style='color:#FFD700; font-weight:900; font-size:16px'>3- مسندة لـ</label>", unsafe_allow_html=True)
            مسندة_ل = st.selectbox("مسندة لـ", ["الصالح", "الضد"], label_visibility="collapsed")
            st.markdown("</div>", unsafe_allow_html=True)
            
            if st.form_submit_button("💾 حفظ الحكم", use_container_width=True, type="primary"):
                if not منطوق_الحكم:
                    st.error("❌ لازم تكتب منطوق الحكم")
                else:
                    case['حالة'] = 'منتهية'
                    case['تاريخ_الحكم'] = str(تاريخ_حكم)
                    case['منطوق_الحكم'] = منطوق_الحكم
                    case['مسندة_ل_الحكم'] = مسندة_ل
                    case['جلسات'].append({'تاريخ':str(تاريخ_حكم),'الرول':'-','سبب':f'الحكم - مسندة لـ {مسندة_ل}','ملاحظات':منطوق_الحكم})
                    case['تاريخ_جلسة'] = str(تاريخ_حكم)
                    case['سبب'] = f'الحكم - مسندة لـ {مسندة_ل}'
                    save_data(data)
                    st.success(f"✅ تم حفظ الحكم واغلاق القضية. تم نقلها للارشيف")
                    st.session_state.page = "الأرشيف"
                    st.rerun()
    else:
        st.success(f"✅ تم الحكم بتاريخ: {case.get('تاريخ_الحكم')}")
        st.info(f"**مسندة لـ:** {case.get('مسندة_ل_الحكم')}")
        st.warning(f"**المنطوق:** {case.get('منطوق_الحكم')}")
        
        if st.button("↩️ ارجاع القضية للتداول", use_container_width=True):
            case['حالة'] = 'متداولة'
            case['تاريخ_الحكم'] = ""
            case['منطوق_الحكم'] = ""
            case['مسندة_ل_الحكم'] = ""
            save_data(data)
            st.session_state.page = "الحصر"
            st.rerun()
            
    st.markdown("</div>", unsafe_allow_html=True)
    # ================================================
# ================================================
# ============ الجزء الخامس: الأرشيف ============
# ================================================
elif st.session_state.page == "الأرشيف":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FF5252; text-align:center'>📁 أرشيف الأحكام النهائية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()

    # ======= ده الجديد عشان يفتح القضية من البحث =======
    if st.session_state.get('selected_case_id'):
        case_id = st.session_state.selected_case_id
        st.session_state.selected_case_id = None # امسحها
        st.session_state.open_from_search = False
        
        case = next((c for c in data["cases"] if c['id'] == case_id), None)
        if case:
            st.success(f"تم فتح القضية رقم {case.get('رقم','')} لسنة {case.get('سنة','')}")
            st.markdown("---")
            st.markdown(f"### 📂 تفاصيل القضية")
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**المدعي:** {case.get('مدعي','')}")
                st.markdown(f"**نوع الدعوى:** {case.get('نوع_الدعوى','')}")
                st.markdown(f"**المحكمة:** {case.get('محكمة_اسم','')}")
                if case.get('مأمورية'): st.markdown(f"**المأمورية:** {case.get('مأمورية','')}")
                if case.get('دائرة'): st.markdown(f"**الدائرة:** {case.get('دائرة','')}")
            with col2:
                st.markdown(f"**المدعى عليه:** {case.get('مدعي_عليه','')}")
                st.markdown(f"**الموضوع:** {case.get('موضوع','')}")
                st.markdown(f"**اخر جلسة:** {case.get('تاريخ_جلسة','')}")
                st.markdown(f"**الحالة:** {case.get('حالة','')}")
            
            st.markdown("---")
            if st.button("⬅️ الرجوع لجدول الارشيف", use_container_width=True):
                st.rerun()
            st.stop() # مهم عشان ميعرضش الجدول تحت
    # =====================================
    # بحث
    search_query = st.text_input("🔍 ابحث برقم القضية او الخصوم او رقم الطعن", key="search_archive")
    
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()

    # نجيب القضايا اللي اتحكم فيها بس للصالح او للضد
    ended_cases = [c for c in data["cases"] if c.get('حالة') == 'منتهية' and c.get('مسندة_ل_الحكم') in ['الصالح', 'الضد']]
    
    # فلتر البحث
    if search_query:
        ended_cases = [c for c in ended_cases if 
                       search_query in str(c.get('رقم','')) or 
                       search_query in str(c.get('مدعي','')) or
                       search_query in str(c.get('مدعي_عليه','')) or
                       search_query in str(c.get('رقم_الطعن',''))]

    if not ended_cases:
        st.info("لا توجد أحكام نهائية في الارشيف")
    else:
        # ترتيب من الاحدث للاقدم
        sorted_ended = sorted(ended_cases, key=lambda x: x.get("تاريخ_الحكم","0000-00-00"), reverse=True)
        total_ended = len(sorted_ended)
        
        st.markdown(f"<div style='background:#1E2A47; padding:20px; border-radius:15px; border:2px solid #FF5252; text-align:center; margin-bottom:20px'>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:28px; font-weight:900; color:#FF5252'>📁 {total_ended}</div><div style='font-size:18px; color:#FFF; font-weight:700'>عدد الأحكام النهائية</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        for idx, case in enumerate(sorted_ended, 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            if case.get('رقم_الطعن'): رقم_كامل += f"<br><b style='color:#FFD700'>طعن رقم:</b> {case.get('رقم_الطعن')}"
            
            محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
            if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = f"{case.get('دائرة','')} عمال" if case.get('دائرة','') else ""
            if دائرة_كاملة: محكمة_كاملة += f"<br>{دائرة_كاملة}"
            
            خصوم = f"<div style='background:#FFF3CD; padding:8px; border-radius:8px; color:#000; margin-bottom:5px; text-align:center'><b>المدعى:</b><br>{case.get('مدعي','')}</div><div style='background:#CFF4FC; padding:8px; border-radius:8px; color:#000; text-align:center'><b>المدعى عليه:</b><br>{case.get('مدعي_عليه','')}</div>"
            
            # ملخص بيانات الحكم
            مسندة = case.get('مسندة_ل_الحكم','')
            لون_مسندة = "#4CAF50" if مسندة == "الصالح" else "#FF5252"
            بيانات_الحكم = f"<div style='background:#142038; padding:8px; border-radius:8px; border:2px solid {لون_مسندة}; text-align:center'><b style='color:{لون_مسندة}; font-size:16px'>تاريخ الحكم:</b><br>{case.get('تاريخ_الحكم','')}<br><b style='color:{لون_مسندة}'>مسندة لـ:</b> {مسندة}<br><b style='color:{لون_مسندة}'>المنطوق:</b><br>{case.get('منطوق_الحكم','')}</div>"
            
            # بيانات الحفظ لو موجودة
            بيانات_الحفظ = f"<div style='background:#2A3A5F; padding:8px; border-radius:8px; color:#FFF; text-align:center'><b>سبب الحفظ:</b><br>{case.get('سبب_الحفظ','-')}</div>" if case.get('سبب_الحفظ') else "-"
            
            # بيانات اعادة التداول لو موجودة
            بيانات_العودة = f"<div style='background:#5F2A2A; padding:8px; border-radius:8px; color:#FFF; text-align:center'><b>سبب العودة:</b><br>{case.get('سبب_العودة','-')}</div>" if case.get('سبب_العودة') else "-"

            st.markdown("<div class='table-container'>", unsafe_allow_html=True)
            st.markdown(f"<table class='case-table'><tr><th>م</th><th>الرقم والسنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>بيانات الحكم</th><th>الحفظ</th><th>عودة للتداول</th></tr><tr class='row-judgment'><td>{idx}</td><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td>{بيانات_الحكم}</td>", unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            with col1:
                سبب_الحفظ = st.text_input("سبب الحفظ", key=f"save_reason_{case['id']}", placeholder="اكتب سبب الحفظ...")
                if st.button("💾 حفظ", key=f"save_btn_{case['id']}", use_container_width=True):
                    if سبب_الحفظ:
                        case['سبب_الحفظ'] = سبب_الحفظ
                        save_data(data)
                        st.success("تم حفظ سبب الحفظ")
                        st.rerun()
                    else:
                        st.error("اكتب سبب الحفظ الاول")
                st.markdown(بيانات_الحفظ, unsafe_allow_html=True)
                
            with col2:
                سبب_العودة = st.text_input("سبب العودة للتداول", key=f"return_reason_{case['id']}", placeholder="ليه بترجعها؟")
                if st.button("↩️ عودة للتداول", key=f"return_{case['id']}", use_container_width=True):
                    if سبب_العودة:
                        case['حالة'] = 'متداولة'
                        case['سبب_العودة'] = سبب_العودة
                        case['تاريخ_الحكم'] = ""
                        case['منطوق_الحكم'] = ""
                        case['مسندة_ل_الحكم'] = ""
                        save_data(data)
                        st.success("تم ارجاع القضية للتداول")
                        st.session_state.page = "الحصر"
                        st.rerun()
                    else:
                        st.error("لازم تكتب سبب العودة")
                st.markdown(بيانات_العودة, unsafe_allow_html=True)
            
            st.markdown("</tr></table></div>", unsafe_allow_html=True)
            # ================================================
# ============ الجزء السادس: البحث ============
# ================================================
elif st.session_state.page == "بحث":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#4DA8DA; text-align:center'>🔍 البحث عن دعوى</h2>", unsafe_allow_html=True)

    if st.button("⬅️ العودة للرئيسية", use_container_width=True, key="back_from_search"):
        st.session_state.page = "الرئيسية"
        st.rerun()

    st.markdown("<div style='background:#1E2A47; padding:20px; border-radius:15px; border:2px solid #4DA8DA; margin-bottom:20px'>", unsafe_allow_html=True)
    search_type = st.selectbox("نوع البحث", [
        "اسم المدعي او المدعى عليه",
        "رقم الدعوى وسنتها - اول درجة",
        "رقم الاستئناف وسنته - ثاني درجة",
        "رقم الطعن بالنقض وسنته",
        "رقم الدعوى وسنتها - المحكمة الادارية",
        "رقم الدعوى وسنتها - محكمة القضاء الاداري",
        "رقم الطعن وسنته - القضاء الاداري بهيئة استئنافية",
        "رقم الدعوى وسنتها - المحكمة الادارية العليا"
    ], key="search_type_select")

    search_value = st.text_input("اكتب كلمة البحث", placeholder="للاسم: اكتب الاسم | للرقم: اكتب الرقم والسنة مثال 1234 2024", key="search_input")
    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🔍 بحث", use_container_width=True, type="primary", key="do_search"):
        if not search_value.strip():
            st.error("اكتب حاجة تبحث بيها")
        else:
            results = []
            parts = search_value.split()
            رقم_بحث = parts[0] if len(parts) > 0 else ""
            سنة_بحث = parts[1] if len(parts) > 1 else ""

            for case in data["cases"]:
                match = False
                نوع_الدعوى = str(case.get('نوع_الدعوى', '')).lower()

                if search_type == "اسم المدعي او المدعى عليه":
                    if search_value in str(case.get('مدعي','')) or search_value in str(case.get('مدعي_عليه','')):
                        match = True

                elif "اول درجة" in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and ("دعوى" in نوع_الدعوى or "عمال" in نوع_الدعوى):
                        match = True

                elif "استئناف" in search_type and "استئنافية" not in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and "استئناف" in نوع_الدعوى:
                        match = True

                elif "الطعن بالنقض" in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and "طعن" in نوع_الدعوى and "اداري" not in نوع_الدعوى:
                        match = True

                elif "المحكمة الادارية" in search_type and "العليا" not in search_type and "استئنافية" not in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and "ادارية" in نوع_الدعوى and "عليا" not in نوع_الدعوى:
                        match = True

                elif "القضاء الاداري" in search_type and "استئنافية" not in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and "قضاء اداري" in نوع_الدعوى:
                        match = True

                elif "استئنافية" in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and "استئنافية" in نوع_الدعوى:
                        match = True

                elif "الادارية العليا" in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and "ادارية عليا" in نوع_الدعوى:
                        match = True

                if match:
                    results.append(case)

            if not results:
                st.warning("لم يتم العثور على نتائج")
            else:
                st.success(f"تم العثور على {len(results)} نتيجة")
                for idx, case in enumerate(results, 1):
                    رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
                    محكمة_كاملة = f"{case.get('نوع_الدعوى','')} - {case.get('محكمة_اسم','')}"
                    if case.get('مأمورية',''):
                        محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
                    if case.get('دائرة',''):
                        محكمة_كاملة += f"<br>دائرة {case.get('دائرة','')}"

                    خصوم = f"<div style='background:#FFF3CD; padding:8px; border-radius:8px; color:#000; margin-bottom:5px; text-align:center'><b>المدعى:</b><br>{case.get('مدعي','')}</div><div style='background:#CFF4FC; padding:8px; border-radius:8px; color:#000; text-align:center'><b>المدعى عليه:</b><br>{case.get('مدعي_عليه','')}</div>"

                    if case.get('حالة') == 'منتهية':
                        row_class = "row-judgment"
                        حالة_لون = "#FF5252"
                        مكان = "📁 الأرشيف"
                        الصفحة_المطلوبة = "الأرشيف" # <-- اتعدل هنا بالالف واللام
                    else:
                        row_class = "row1" if idx % 2 == 1 else "row2"
                        حالة_لون = "#4CAF50"
                        مكان = "📋 الحصر العام"
                        الصفحة_المطلوبة = "الحصر"

                    st.markdown("<div class='table-container'>", unsafe_allow_html=True)
                    st.markdown(f"""
                    <table class='case-table'>
                    <tr><th>م</th><th>الرقم والسنة</th><th>نوع الدعوى والمحكمة</th><th>الخصوم</th><th>الموضوع</th><th>اخر جلسة</th><th>الحالة</th><th>المكان</th></tr>
                    <tr class='{row_class}'>
                        <td>{idx}</td>
                        <td>{رقم_كامل}</td>
                        <td>{محكمة_كاملة}</td>
                        <td>{خصوم}</td>
                        <td>{case.get('موضوع','')}</td>
                        <td style='color:#FFD700; font-weight:900'>{case.get('تاريخ_جلسة','-')}</td>
                        <td style='color:{حالة_لون}; font-weight:900'>{case.get('حالة','متداولة')}</td>
                        <td style='color:#4DA8DA; font-weight:900'>{مكان}</td>
                    </tr>
                    </table>
                    """, unsafe_allow_html=True)

                    c1, c2, c3 = st.columns([4,1,4])
                    with c2:
                        if st.button(f"📂 فتح في {مكان}", key=f"open_smart_btn_{case['id']}", use_container_width=True):
                            st.session_state.selected_case_id = case['id']
                            st.session_state.open_from_search = True # العلامة عشان الصفحة التانية تعرف
                            st.session_state.page = الصفحة_المطلوبة
                            st.rerun()
                    st.markdown("</div>", unsafe_allow_html=True)
