import streamlit as st
import pandas as pd
from datetime import datetime
import json, os

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

# التصميم مطابق للصورة 100%
st.markdown("""
<style>
    .stApp {background: #0a1a3a;}
    [data-testid="stSidebar"] {background: #0d2347;}
    h1 {color: #DAA520 !important; text-align: center; font-size: 45px; font-weight: bold;}
    .card {background: #122c5e; padding: 25px; border-radius: 15px; border: 3px solid #DAA520; margin: 15px 0; text-align: center;}
    .card h3 {color: #E0E0E0; font-size: 22px;}
    .card h2 {color: #DAA520; font-size: 50px; font-weight: bold;}
    label {color: white !important; font-weight: bold; font-size: 16px;}
    .logo-text {animation: glow 2s infinite alternate; color: #00BFFF; font-weight: bold; text-align: center; position: fixed; bottom: 5px; width: 100%; font-size: 13px;}
    @keyframes glow {from{text-shadow: 0 0 5px #00BFFF} to{text-shadow: 0 0 15px #FFD700}}
    .watermark {position: fixed; top: 45%; left: 50%; transform: translate(-50%, -50%); font-size: 350px; opacity: 0.04; color: #DAA520; z-index: -1;}
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='watermark'>⚖️</div>", unsafe_allow_html=True)

# الهيدر + الميزان
st.markdown("<h1>⚖️</h1>", unsafe_allow_html=True)
st.markdown("<h1>إدارة القضايا</h1>", unsafe_allow_html=True)

# السايدبار زي الصورة
with st.sidebar:
    st.image("https://i.imgur.com/8Qf2ZQp.png", width=50) # ايقونة الميزان
    menu = st.radio("", ["📝 تسجيل القضايا", "📊 الحصر العام", "🔍 البحث", "🔔 التنبيهات", "📑 التقارير", "🗃️ الأرشيف", "📚 المكتبة القانونية", "📈 الإحصائيات"])

# الكروت ال 4 زي الصورة
c1,c2,c3,c4 = st.columns(4)
c1.markdown(f"<div class='card'>📁 <h3>عدد القضايا المتداولة</h3><h2>{len(st.session_state.cases)}</h2></div>", unsafe_allow_html=True)
c2.markdown(f"<div class='card'>📅 <h3>عدد الجلسات القادمة</h3><h2>0</h2></div>", unsafe_allow_html=True)
c3.markdown(f"<div class='card'>✅ <h3>أحكام لصالح الهيئة</h3><h2>0</h2></div>", unsafe_allow_html=True)
c4.markdown(f"<div class='card'>❌ <h3>أحكام ضد الهيئة</h3><h2>0</h2></div>", unsafe_allow_html=True)

st.markdown("---")
st.markdown("<h2 style='color:#DAA520; text-align:center'>الحصر العام</h2>", unsafe_allow_html=True)

# 1. تسجيل القضايا
if menu == "📝 تسجيل القضايا":
    with st.form("form_case", clear_on_submit=True):
        st.subheader("تسجيل قضية جديدة")
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
        
        if st.form_submit_button("💾 حفظ القضية"):
            قضية = {"رقم": f"{رقم} لسنة {سنة}", "نوع": نوع, "المحكمة": اسم_المحكمة, "الدائرة": دائرة, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, "موضوع": موضوع, "تاريخ": str(تاريخ), "رول": رول, "سبب": سبب}
            st.session_state.cases.append(قضية)
            save_data()
            st.success("تم الحفظ وانتقلت للحصر العام")
            st.rerun()

# الفوتر المتحرك
st.markdown("<p class='logo-text'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</p>", unsafe_allow_html=True)
