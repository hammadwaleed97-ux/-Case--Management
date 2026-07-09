import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import json, os, io
from fpdf import FPDF
from docx import Document

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

DB_FILE = "cases_db.json"
def load_data():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, 'r', encoding='utf-8') as f: return json.load(f)
    return {"cases": [], "archive": [], "library": []}
def save_data():
    with open(DB_FILE, 'w', encoding='utf-8') as f:
        json.dump({"cases": st.session_state.cases, "archive": st.session_state.archive, "library": st.session_state.library}, f, ensure_ascii=False, indent=2)

data = load_data()
if 'cases' not in st.session_state: st.session_state.cases = data["cases"]
if 'archive' not in st.session_state: st.session_state.archive = data["archive"]
if 'library' not in st.session_state: st.session_state.library = data["library"]

# التصميم الفخم مطابق للصورة
st.markdown("""
<style>
    .stApp {background: #0a1a3a;}
    [data-testid="stSidebar"] {background: #0d2347; border-right: 3px solid #DAA520;}
    h1 {color: #DAA520 !important; text-align: center; font-size: 40px;}
    .logo-text {animation: glow 2s infinite alternate; color: #00BFFF; font-weight: bold; text-align: center; position: fixed; bottom: 10px; width: 100%; font-size: 14px;}
    @keyframes glow {from{text-shadow: 0 0 5px #00BFFF} to{text-shadow: 0 0 20px #FFD700}}
    .card {background: #122c5e; padding: 15px; border-radius: 10px; border: 1px solid #DAA520; margin: 5px;}
    .metric-card {background: linear-gradient(135deg, #1e3a5f, #2a4a7c); padding: 20px; border-radius: 10px; text-align: center; border: 2px solid #DAA520;}
    .alert {background: #8B0000; color: white; padding: 10px; border-radius: 8px; border-left: 5px solid yellow;}
    label, .stSelectbox label, .stTextInput label {color: white !important; font-weight: bold;}
    .watermark {position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); font-size: 400px; opacity: 0.03; color: #DAA520; z-index: -1;}
    .sidebar-btn {width: 100%; margin: 5px 0;}
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='watermark'>⚖️</div>", unsafe_allow_html=True)

# الهيدر
st.markdown("<h1>⚖️</h1>", unsafe_allow_html=True)
st.markdown("<h1>إدارة القضايا</h1>", unsafe_allow_html=True)

# السايدبار
with st.sidebar:
    st.markdown("<h2 style='text-align:center; color:#DAA520'>القائمة الرئيسية</h2>", unsafe_allow_html=True)
    menu = st.radio("", ["تسجيل القضايا", "الحصر العام", "البحث", "التنبيهات", "التقارير", "الأرشيف", "المكتبة القانونية", "الإحصائيات"])

# كروت الإحصائيات فوق - زي الصورة
c1,c2,c3,c4 = st.columns(4)
c1.markdown(f"<div class='metric-card'><h3>📁 عدد القضايا المتداولة</h3><h2>{len(st.session_state.cases)}</h2></div>", unsafe_allow_html=True)
c2.markdown(f"<div class='metric-card'><h3>📅 عدد الجلسات القادمة</h3><h2>0</h2></div>", unsafe_allow_html=True)
c3.markdown(f"<div class='metric-card'><h3>✅ أحكام لصالح الهيئة</h3><h2>0</h2></div>", unsafe_allow_html=True)
c4.markdown(f"<div class='metric-card'><h3>❌ أحكام ضد الهيئة</h3><h2>0</h2></div>", unsafe_allow_html=True)
st.markdown("---")

# 1. تسجيل القضايا
if menu == "تسجيل القضايا":
    st.subheader("📝 تسجيل قضية جديدة")
    with st.form("form_case", clear_on_submit=True):
        نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
        محكمة = st.selectbox("نوع المحكمة", ["الابتدائية","الاستئناف","النقض","الإدارية","القضاء الإداري","الإدارية العليا"])
        اسم_المحكمة = st.text_input("اسم المحكمة")
        if نوع == "استئناف": مأمورية = st.text_input("المأمورية")
        c1,c2,c3 = st.columns(3)
        رقم = c1.text_input("رقم الدعوى")
        سنة = c2.text_input("السنة القضائية")
        دائرة = c3.text_input("الدائرة")
        مدعي = st.text_input("اسم المدعي / المستأنف / الطاعن")
        مدعي_عليه = st.text_input("اسم المدعي عليه / المستأنف ضده")
        موضوع = st.text_area("موضوع الدعوى")
        c1,c2,c3 = st.columns(3)
        تاريخ = c1.date_input("تاريخ أول جلسة")
        رول = c2.text_input("الرول")
        سبب = c3.text_input("سبب الجلسة")
        ملاحظات = st.text_area("ملاحظات")
        واتس = st.checkbox("تفعيل التنبيهات عبر الواتس اب")
        رقم_واتس = st.text_input("رقم هاتف واتس اب") if واتس else ""
        نوع_المستند = st.selectbox("تحميل المستندات", ["صحيفة الدعوى","صحيفة الاستئناف","صحيفة الطعن"])
        ملف = st.file_uploader("اختر الملف")
        
        col1,col2 = st.columns(2)
        if col1.form_submit_button("💾 حفظ القضية"):
            قضية = {"id": str(datetime.now().timestamp()), "رقم": f"{رقم} لسنة {سنة}", "نوع": نوع, "المحكمة": اسم_المحكمة, "مأمورية": مأمورية if نوع=="استئناف" else "", "الدائرة": دائرة, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, "موضوع": موضوع, "تاريخ": str(تاريخ), "رول": رول, "سبب": سبب, "ملاحظات": ملاحظات, "الحالة": "متداولة", "الجلسات": [{"الرول": رول, "التاريخ": str(تاريخ), "الإجراءات": سبب}], "المستندات": [], "حكم": {}}
            st.session_state.cases.append(قضية)
            save_data()
            st.success("تم الحفظ وانتقلت للحصر العام")
            st.rerun()
        col2.form_submit_button("🗑️ حذف")

# 2. الحصر العام - كروت زي الصورة
elif menu == "الحصر العام":
    st.subheader("📊 الحصر العام")
    for i, ق in enumerate(sorted(st.session_state.cases, key=lambda x: x['تاريخ'])):
        with st.container():
            st.markdown(f"<div class='card'>", unsafe_allow_html=True)
            st.markdown(f"**رقم:** {ق['رقم']} | **الدائرة:** {ق['الدائرة']} | **المحكمة:** {ق['المحكمة']}")
            st.markdown(f"**الخصوم:** {ق['مدعي']} ضد {ق['مدعي_عليه']}")
            st.markdown(f"**الموضوع:** {ق['موضوع']} | **آخر جلسة:** {ق['تاريخ']} | **السبب:** {ق['سبب']}")
            
            with st.expander("عرض التفاصيل وإضافة جلسة"):
                st.dataframe(pd.DataFrame(ق["الجلسات"]), use_container_width=True)
                c1,c2,c3 = st.columns(3)
                ت_جديد = c1.date_input("تاريخ الجلسة القادمة", key=f"d{i}")
                ر_جديد = c2.text_input("الرول", key=f"r{i}")
                س_جديد = c3.text_input("سبب التأجيل", key=f"s{i}")
                if st.button("حفظ الجلسة", key=f"save{i}"):
                    ق["الجلسات"].append({"الرول": ر_جديد, "التاريخ": str(ت_جديد), "الإجراءات": س_جديد})
                    ق["تاريخ"] = str(ت_جديد); ق["سبب"] = س_جديد; save_data(); st.rerun()
                
                st.markdown("### رفع المستندات")
                نوع_م = st.selectbox("نوع المستند", ["مذكرة دفاع","حافظة مستندات","تقرير خبير","صورة حكم تمهيدي","أخرى"], key=f"doc{i}")
                بيان_م = st.text_input("بيان المستند", key=f"bayn{i}")
                ملف_م = st.file_uploader("المستند", key=f"file{i}")
                if st.button("حفظ المستند", key=f"savdoc{i}"):
                    ق["المستندات"].append({"النوع": نوع_م, "البيان": بيان_م}); save_data()
                
                st.markdown("### بيانات الحكم")
                ت_حكم = st.date_input("تاريخ جلسة الحكم", key=f"th{i}")
                منطوق = st.text_area("منطوق الحكم", key=f"mn{i}")
                النتيجة = st.selectbox("النتيجة", ["لصالح الهيئة","ضد الهيئة"], key=f"nt{i}")
                if st.button("حفظ الحكم وأرشفة", key=f"hk{i}"):
                    ق["حكم"] = {"التاريخ": str(ت_حكم), "المنطوق": منطوق, "النتيجة": النتيجة}
                    st.session_state.archive.append(ق); st.session_state.cases.pop(i); save_data(); st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

# 3. التنبيهات
elif menu == "التنبيهات":
    st.subheader("🔔 التنبيهات")
    تفعيل = st.checkbox("تفعيل إرسال تنبيهات واتس الساعة 9 صباحاً")
    if تفعيل: st.text_input("رقم الواتس")
    اليوم = datetime.now().date()
    for ق in st.session_state.cases:
        if (datetime.strptime(ق['تاريخ'], '%Y-%m-%d').date() - اليوم).days <= 7:
            st.markdown(f"<div class='alert'>⚠️ جلسة {ق['تاريخ']} - {ق['رقم']}</div>", unsafe_allow_html=True)

# 4. التقارير
elif menu == "التقارير":
    نوع = st.selectbox("نوع التقرير", ["الدعاوى المتداولة","الاحكام"])
    df = pd.DataFrame(st.session_state.cases if نوع=="الدعاوى المتداولة" else st.session_state.archive)
    st.dataframe(df, use_container_width=True)
    col1,col2,col3,col4 = st.columns(4)
    col1.button("📄 فتح PDF")
    col2.button("📝 فتح Word")
    col3.download_button("📥 تحميل PDF", df.to_csv().encode('utf-8'), "report.csv")
    col4.download_button("📥 تحميل Word", df.to_csv().encode('utf-8'), "report.docx")

# 5. المكتبة
elif menu == "المكتبة القانونية":
    بحث = st.text_input("🔍 بحث في المكتبة")
    اقسام = ["القوانين","القرارات الوزارية","احكام النقض","احكام الدستورية","أخرى"]
    for قسم in اقسام:
        with st.expander(f"📁 {قسم}"):
            st.file_uploader(f"رفع مستند", key=قسم)

# 6. الأرشيف
elif menu == "الأرشيف":
    st.dataframe(pd.DataFrame(st.session_state.archive), use_container_width=True)

# 7. الإحصائيات
elif menu == "الإحصائيات":
    st.metric("المتداولة", len(st.session_state.cases))
    st.metric("المؤرشفة", len(st.session_state.archive))

# الفوتر المتحرك
st.markdown("<p class='logo-text'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</p>", unsafe_allow_html=True)
