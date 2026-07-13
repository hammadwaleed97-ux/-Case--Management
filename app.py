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
# ==========================================
# ================================================
# ============ الجزء الثاني: تسجيل القضايا ============
# ================================================

elif st.session_state.page == "تسجيل":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#D4AF37; text-align:center'>➕ تسجيل القضايا</h2>", unsafe_allow_html=True)
    
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): 
        st.session_state.page = "الرئيسية"; st.rerun()

    نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
    
    with st.form("form_case"):
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        محكمة_اسم = st.text_input("اسم المحكمة")
        مأمورية = st.text_input("المأمورية") if نوع == "استئناف" else ""
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>2- بيانات الدعوى</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: رقم = st.text_input("رقم الدعوى / الاستئناف / الطعن")
        with col2: سنة = st.text_input("السنة القضائية")
        دائرة = st.text_input("الدائرة")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>3- بيانات الخصوم</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: مدعي = st.text_input("اسم المدعى / المستأنف / الطاعن")
        with col2: مدعي_عليه = st.text_input("اسم المدعى عليه / المستأنف ضده / المطعون ضده")
        موضوع = st.text_area("موضوع الدعوى", height=100)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>4- بيانات الجلسة</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: تاريخ_جلسة = st.date_input("تاريخ أول جلسة", value=datetime.now())
        with col2: الرول = st.text_input("الرول")
        سبب = st.text_input("سبب الجلسة")
        ملاحظات = st.text_area("ملاحظات", height=100)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>5- صحيفة الدعوى</div>", unsafe_allow_html=True)
        مسندة_ل = st.selectbox("نوع الصحيفة", ["صحيفة الدعوى"])
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='text-align:center'>", unsafe_allow_html=True)
        if st.form_submit_button("💾 حفظ القضية", use_container_width=True, type="primary"):
            if not رقم or not سنة: 
                st.error("❌ من فضلك ادخل رقم الدعوى والسنة")
            else:
                new_case = {
                    "id": len(data["cases"])+1, "نوع": نوع, "محكمة_اسم": محكمة_اسم, "مأمورية": مأمورية, 
                    "رقم": رقم, "سنة": سنة, "دائرة": دائرة, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, 
                    "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة), "الرول": الرول, "سبب": سبب, 
                    "ملاحظات": ملاحظات, "جلسات": [], "مستندات": [], "حالة": "متداولة", "مسندة_ل": مسندة_ل
                }
                if الرول or سبب: 
                    new_case["جلسات"].append({"تاريخ":str(تاريخ_جلسة),"الرول":الرول,"سبب":سبب,"ملاحظات":ملاحظات})
                data["cases"].append(new_case)
                save_data(data)
                st.success(f"✅ تم حفظ القضية رقم {رقم} لسنة {سنة}")
                st.session_state.page = "الحصر"; st.rerun() # <-- صلحتها هنا
        st.markdown("</div>", unsafe_allow_html=True)

# ================================================
# ========== نهاية الجزء الثاني ==========
# ==============================================
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

    /* جدول الحصر */
    .case-table { width:100%; color:#FFFFFF; text-align:center; border-collapse: collapse; font-size:14px }
    .case-table th { background:#D4AF37; color:#0A1428; padding:10px; font-weight:900; border:1px solid #0A1428 }
    .case-table td { padding:10px; border:1px solid #D4AF37; }
    .table-container { background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:20px; box-shadow:0 4px 12px rgba(0,0,0,0.4) }
    .row1 { background: #142038; }
    .row2 { background: #1E2A47; }
    .row-hey2a { background: #1E3A6B; }
    .row-judgment { background: #2C2F33; color:#AAA }
    
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
    with open(TOKENS_FILE, "w", encoding="utf-8") as f: json.dump(tokens_data, f, ensure_ascii=False, indent=4)

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
        if st.button("🔴 مركز التنبيهات", use_container_width=True): st.session_state.page = "التنبيهات"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="btn-report">', unsafe_allow_html=True)
        if st.button("📊 التقارير", use_container_width=True): st.session_state.page = "التقارير"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="btn-lib">', unsafe_allow_html=True)
        if st.button("📚 المكتبة القانونية", use_container_width=True): st.session_state.page = "المكتبة"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown('<div class="btn-arch">', unsafe_allow_html=True)
        if st.button("🗄️ الأرشيف", use_container_width=True): st.session_state.page = "الأرشيف"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown('<div class="btn-search">', unsafe_allow_html=True)
        if st.button("🔍 البحث عن دعوى", use_container_width=True): st.session_state.page = "البحث"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# ================================================
# ============ الجزء الثاني: تسجيل القضايا ============
# ================================================
elif st.session_state.page == "تسجيل":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#D4AF37; text-align:center'>➕ تسجيل القضايا</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()

    نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
    with st.form("form_case"):
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        محكمة_اسم = st.text_input("اسم المحكمة"); مأمورية = st.text_input("المأمورية") if نوع == "استئناف" else ""
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>2- بيانات الدعوى</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: رقم = st.text_input("رقم الدعوى / الاستئناف / الطعن")
        with col2: سنة = st.text_input("السنة القضائية")
        دائرة = st.text_input("الدائرة"); st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>3- بيانات الخصوم</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: مدعي = st.text_input("اسم المدعى / المستأنف / الطاعن")
        with col2: مدعي_عليه = st.text_input("اسم المدعى عليه / المستأنف ضده / المطعون ضده")
        موضوع = st.text_area("موضوع الدعوى", height=100); st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>4- بيانات الجلسة</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: تاريخ_جلسة = st.date_input("تاريخ أول جلسة", value=datetime.now())
        with col2: الرول = st.text_input("الرول")
        سبب = st.text_input("سبب الجلسة"); ملاحظات = st.text_area("ملاحظات", height=100); st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>5- صحيفة الدعوى</div>", unsafe_allow_html=True)
        مسندة_ل = st.selectbox("نوع الصحيفة", ["صحيفة الدعوى"]); st.markdown("</div>", unsafe_allow_html=True)

        if st.form_submit_button("💾 حفظ القضية", use_container_width=True, type="primary"):
            if not رقم or not سنة: st.error("❌ من فضلك ادخل رقم الدعوى والسنة")
            else:
                new_case = {"id": len(data["cases"])+1, "نوع": نوع, "محكمة_اسم": محكمة_اسم, "مأمورية": مأمورية, "رقم": رقم, "سنة": سنة, "دائرة": دائرة, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة), "الرول": الرول, "سبب": سبب, "ملاحظات": ملاحظات, "جلسات": [], "مستندات": [], "حالة": "متداولة", "مسندة_ل": مسندة_ل}
                if الرول or سبب: new_case["جلسات"].append({"تاريخ":str(تاريخ_جلسة),"الرول":الرول,"سبب":سبب,"ملاحظات":ملاحظات})
                data["cases"].append(new_case); save_data(data)
                st.success(f"✅ تم حفظ القضية رقم {رقم} لسنة {سنة}")
                st.session_state.page = "الحصر"; st.rerun()

# ================================================
# ============ الجزء الثالث: الحصر العام ============
# ================================================
elif st.session_state.page == "الحصر":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📊 الحصر العام الخارجي</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()

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
        
        st.markdown(f"<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; text-align:center; margin-bottom:20px'>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        col1.metric("📊 اجمالي القضايا", total)
        col2.metric("📅 جلسات هذا الاسبوع", this_week)
        col3.metric("🚫 عدد المنتهية", ended)
        st.markdown("</div>", unsafe_allow_html=True)

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

    # 1- بيانات القضية كارت
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>1- بيانات القضية</div>", unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("رقم القضية", case.get('رقم'))
    col2.metric("السنة", case.get('سنة'))
    col3.metric("الدائرة", f"{case.get('دائرة')} عمال")
    col4.metric("النوع", case.get('نوع'))
    st.write(f"**المحكمة:** {case.get('محكمة_اسم')} {case.get('مأمورية')}")
    st.write(f"**الموضوع:** {case.get('موضوع')}")
    st.write(f"**الحالة:** {case.get('حالة')}")
    st.markdown("</div>", unsafe_allow_html=True)

    # 2- بيانات الخصوم كارت
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>2- بيانات الخصوم</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1: st.markdown(f"<div style='background:#FFF3CD; padding:10px; border-radius:10px; color:#000'><b>المدعى:</b><br>{case.get('مدعي')}</div>", unsafe_allow_html=True)
    with col2: st.markdown(f"<div style='background:#CFF4FC; padding:10px; border-radius:10px; color:#000'><b>المدعى عليه:</b><br>{case.get('مدعي_عليه')}</div>", unsafe_allow_html=True)
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
    
    with st.expander("➕ اضافة جلسة جديدة"):
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
              
