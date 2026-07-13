# ==============================================
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
    with open(TOKENS_FILE, "w", encoding="utf-8") as f:
        json.dump(tokens_data, f, ensure_ascii=False, indent=4)

# ============ الجزء الثالث: الحصر العام ============
if st.session_state.page == "الرئيسية":
    pass
elif st.session_state.page == "حصر":
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📊 الحصر العام الخارجي</h2>", unsafe_allow_html=True)
elif st.session_state.page == "حصر":
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📊 الحصر العام الخارجي</h2>", unsafe_allow_html=True)
    
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): 
        st.session_state.page = "الرئيسية"; st.rerun()

    if not data["cases"]:
        st.info("لا توجد قضايا مسجلة")
    else:
        # ندي id لو مش موجود
        for i, case in enumerate(data["cases"]):
            if "id" not in case: case["id"] = i + 1
            if "مستندات" not in case: case["مستندات"] = []

        save_data(data)
        
        # نرتب حسب تاريخ الجلسة
        sorted_cases = sorted(data["cases"], key=lambda x: x.get("تاريخ_جلسة","9999-12-31"))
        
        total = len(sorted_cases)
        today = datetime.now().date()
        this_week = len([c for c in sorted_cases if c.get('تاريخ_جلسة') and datetime.strptime(c['تاريخ_جلسة'],'%Y-%m-%d').date() <= today + timedelta(days=7)])
        
        st.markdown(f"<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; text-align:center; margin-bottom:20px'>", unsafe_allow_html=True)
        st.markdown(f"<h3 style='color:#D4AF37'>📊 اجمالي القضايا: {total} | جلسات هذا الاسبوع: {this_week}</h3>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        for idx, case in enumerate(sorted_cases, 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
            if case.get('مأمورية',''): 
                محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = f"{case.get('دائرة','')} عمال" if case.get('دائرة','') else ""
            محكمة_كاملة += f"<br>{دائرة_كاملة}"
            خصوم = f"{case.get('مدعي','')}<br>ضد<br>{case.get('مدعي_عليه','')}"

            # الوان الصفوف
            if case.get('حالة') == 'منتهية': 
                bg_color = "#2C2F33"
            elif "الهيئة" in str(case.get('مدعي','')): 
                bg_color = "#1E3A6B"
            else: 
                bg_color = "#142038" if idx % 2 == 1 else "#1E2A47"

            st.markdown(f"<div style='background:{bg_color}; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
            
            # الجدول
            table_html = f"""
            <table style='width:100%; color:#FFFFFF; text-align:center'>
                <tr style='background:#D4AF37; color:#0A1428'>
                    <th style='padding:8px; border-radius:8px'>م</th>
                    <th style='padding:8px; border-radius:8px'>الرقم والسنة</th>
                    <th style='padding:8px; border-radius:8px'>المحكمة والدائرة</th>
                    <th style='padding:8px; border-radius:8px'>الخصوم</th>
                    <th style='padding:8px; border-radius:8px'>الموضوع</th>
                    <th style='padding:8px; border-radius:8px'>اخر جلسة</th>
                    <th style='padding:8px; border-radius:8px'>السبب</th>
                    <th style='padding:8px; border-radius:8px'>الحالة</th>
                </tr>
                <tr>
                    <td style='padding:8px'>{idx}</td>
                    <td style='padding:8px'>{رقم_كامل}</td>
                    <td style='padding:8px'>{محكمة_كاملة}</td>
                    <td style='padding:8px'>{خصوم}</td>
                    <td style='padding:8px'>{case.get('موضوع','')}</td>
                    <td style='padding:8px; color:#FFD700'>{case.get('تاريخ_جلسة','')}</td>
                    <td style='padding:8px'>{case.get('سبب','')}</td>
                    <td style='padding:8px; color:#4CAF50; font-weight:900'>{case.get('حالة','متداولة')}</td>
                </tr>
            </table>
            """
            st.markdown(table_html, unsafe_allow_html=True)

            c1, c2, c3 = st.columns([4,1,4])
            with c2:
                if st.button("فتح", key=f"open_{case['id']}", use_container_width=True): 
                    st.session_state.selected_case_id = case['id']
                    st.session_state.page = "تفاصيل"
                    st.rerun()
            
            st.markdown("</div>", unsafe_allow_html=True)

# ================================================
# ========== نهاية الجزء الثالث ==========
# ================================================

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
# ==========================================
# ================================================
# ============ الجزء الثاني: تسجيل القضايا ============
# ================================================

elif st.session_state.page == "تسجيل":
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#D4AF37; text-align:center'>➕ تسجيل القضايا</h2>", unsafe_allow_html=True)
    
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): 
        st.session_state.page = "الرئيسية"; st.rerun()

    نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
    
    with st.form("form_case"):
        # ===== 1- بيانات المحكمة =====
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        محكمة_اسم = st.text_input("اسم المحكمة")
        مأمورية = st.text_input("المأمورية") if نوع == "استئناف" else ""
        st.markdown("</div>", unsafe_allow_html=True)

        # ===== 2- بيانات الدعوى =====
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>2- بيانات الدعوى</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: رقم = st.text_input("رقم الدعوى / الاستئناف / الطعن")
        with col2: سنة = st.text_input("السنة القضائية")
        دائرة = st.text_input("الدائرة")
        st.markdown("</div>", unsafe_allow_html=True)

        # ===== 3- بيانات الخصوم =====
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>3- بيانات الخصوم</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: مدعي = st.text_input("اسم المدعى / المستأنف / الطاعن")
        with col2: مدعي_عليه = st.text_input("اسم المدعى عليه / المستأنف ضده / المطعون ضده")
        موضوع = st.text_area("موضوع الدعوى", height=100)
        st.markdown("</div>", unsafe_allow_html=True)

        # ===== 4- بيانات الجلسة =====
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>4- بيانات الجلسة</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: تاريخ_جلسة = st.date_input("تاريخ أول جلسة", value=datetime.now())
        with col2: الرول = st.text_input("الرول")
        سبب = st.text_input("سبب الجلسة")
        ملاحظات = st.text_area("ملاحظات", height=100)
        st.markdown("</div>", unsafe_allow_html=True)

        # ===== 5- صحيفة الدعوى =====
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>5- صحيفة الدعوى</div>", unsafe_allow_html=True)
        مسندة_ل = st.selectbox("نوع الصحيفة", ["صحيفة الدعوى"]) # <-- ثابتة
        st.markdown("</div>", unsafe_allow_html=True)

        # ===== زرار الحفظ =====
        st.markdown("<div style='text-align:center'>", unsafe_allow_html=True)
        if st.form_submit_button("💾 حفظ القضية", use_container_width=True, type="primary"):
            if not رقم or not سنة: 
                st.error("❌ من فضلك ادخل رقم الدعوى والسنة")
            else:
                new_case = {
                    "id": len(data["cases"])+1, 
                    "نوع": نوع, 
                    "محكمة_اسم": محكمة_اسم, 
                    "مأمورية": مأمورية, 
                    "رقم": رقم, 
                    "سنة": سنة, 
                    "دائرة": دائرة, 
                    "مدعي": مدعي, 
                    "مدعي_عليه": مدعي_عليه, 
                    "موضوع": موضوع, 
                    "تاريخ_جلسة": str(تاريخ_جلسة), 
                    "الرول": الرول, 
                    "سبب": سبب, 
                    "ملاحظات": ملاحظات, 
                    "جلسات": [], 
                    "مستندات": [],  # فاضية عشان شلنا الرفع
                    "حالة": "متداولة",
                    "مسندة_ل": مسندة_ل
                }
                if الرول or سبب: 
                    new_case["جلسات"].append({
                        "تاريخ":str(تاريخ_جلسة),
                        "الرول":الرول,
                        "سبب":سبب,
                        "ملاحظات":ملاحظات
                    })
                data["cases"].append(new_case)
                save_data(data)
                st.success(f"✅ تم حفظ القضية رقم {رقم} لسنة {سنة}")
                st.session_state.page = "حصر"; st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

# ================================================
# ========== نهاية الجزء الثاني ==========
# ================================================
elif st.session_state.page == "الحصر":
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📊 الحصر العام الخارجي</h2>", unsafe_allow_html=True)
    if st.button("العودة للرئيسية", use_container_width=True): 
        st.session_state.page = "الرئيسية"; st.rerun()

    data = load_data()
    if not data["cases"]:
        st.info("لا توجد قضايا مسجلة")
    else:
        for i, case in enumerate(data["cases"]):
            if "id" not in case: case["id"] = i + 1
            if "مستندات" not in case: case["مستندات"] = []

        save_data(data)
        sorted_cases = sorted(data["cases"], key=lambda x: x.get("تاريخ_جلسة","9999"))
        total = len(sorted_cases)
        today = datetime.now()
        week_later = today + timedelta(days=7)
        this_week = len([c for c in sorted_cases if c.get('تاريخ_جلسة') and datetime.strptime(c['تاريخ_جلسة'],'%Y-%m-%d') <= week_later])
        
        # الاحصائيات فوق
        col1, col2, col3 = st.columns(3)
        col1.metric("📊 اجمالي القضايا", total)
        col2.metric("📅 جلسات هذا الاسبوع", this_week)
        col3.metric("🚫 عدد الحظر", len([c for c in sorted_cases if c.get('حالة') == 'منتهية']))

        st.markdown("---")

        for idx, case in enumerate(sorted_cases, 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
            if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = f"{case.get('دائرة','')} عمال" if case.get('دائرة','') else ""
            محكمة_كاملة += f"<br>{دائرة_كاملة}"
            خصوم = f"{case.get('مدعي','')}<br>ضد<br>{case.get('مدعي_عليه','')}"

            if case.get('حالة') == 'منتهية': row_class = "row-judgment" # ده الحظر
            elif "الهيئة" in str(case.get('مدعي','')): row_class = "row-hey2a"
            else: row_class = "row1" if idx % 2 == 1 else "row2"

            st.markdown("<div class='table-container'>", unsafe_allow_html=True)
            table_html = f"<table class='case-table'><tr><th>م</th><th>الرقم والسنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>اخر جلسة</th><th>السبب</th><th>الحالة</th></tr><tr class='{row_class}'><td>{idx}</td><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td>{case.get('تاريخ_جلسة','')}</td><td>{case.get('سبب','')}</td><td>{case.get('حالة','متداولة')}</td></tr></table></div>"
            st.markdown(table_html, unsafe_allow_html=True)

            c1, c2, c3 = st.columns([4,1,4])
            with c2:
                if st.button("فتح", key=f"open_{case['id']}"): 
                    st.session_state.selected_case_id = case['id']; st.session_state.page = "تفاصيل"; st.rerun()
