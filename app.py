import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

# ملف الحفظ الدائم
DATA_FILE = "cases_data.json"

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"cases": []}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

data = load_data()
today = datetime.now().strftime("%A, %d %B %Y")

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');
    html, body {{font-family: 'Cairo', sans-serif;}}
    .stApp {{background: linear-gradient(180deg, #0a1a3a 0%, #122c5e 100%);}}
    
    /* الشريط العلوي المتحرك */
    .marquee {{
        background: linear-gradient(90deg, #DAA520, #FFD700, #DAA520);
        color: #0a1a3a; padding: 10px; font-weight: 900; font-size: 14px;
        border-bottom: 3px solid #FFD700; position: sticky; top: 0; z-index: 999;
        animation: glow 2s infinite alternate;
    }}
    @keyframes glow {{
        from {{box-shadow: 0 0 10px #FFD700;}}
        to {{box-shadow: 0 0 25px #FFD700;}}
    }}
    
    /* الهيدر */
    .header {{
        background: linear-gradient(135deg, #DAA520, #FFD700);
        padding: 15px; border-radius: 15px; text-align: center;
        border: 3px solid #DAA520; margin: 15px 0;
        box-shadow: 0 0 20px rgba(218,165,32,0.5);
    }}
    .header h1 {{color: #0a1a3a; font-size: 36px; font-weight: 900; margin: 0;}}
    .header p {{color: #0a1a3a; font-size: 16px; font-weight: 700; margin: 5px 0 0 0;}}
    
    /* ازرار الاقسام + الاسماء ظاهرة */
    .section-box {{
        text-align: center; margin: 8px;
    }}
    div[data-testid="stButton"] > button {{
        width: 100%; padding: 18px; border-radius: 12px; 
        border: 3px solid #DAA520; font-weight: 900; font-size: 16px;
        color: white !important; height: 80px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.3);
    }}
    div[data-testid="stButton"] > button:hover {{
        transform: translateY(-3px);
        box-shadow: 0 6px 15px rgba(218,165,32,0.6);
    }}
    .section-name {{
        color: #FFD700; font-size: 12px; font-weight: 700; margin-top: 5px;
    }}
    
    /* الوان الازرار */
    .btn-1 {{background: linear-gradient(135deg, #1e3a5f, #2a4a7c) !important;}}
    .btn-2 {{background: linear-gradient(135deg, #2F4F4F, #4a5f5f) !important;}}
    .btn-3 {{background: linear-gradient(135deg, #8B4513, #A0522D) !important;}}
    .btn-4 {{background: linear-gradient(135deg, #B8860B, #DAA520) !important; color: black !important;}}
    .btn-5 {{background: linear-gradient(135deg, #4B0082, #6A5ACD) !important;}}
    .btn-6 {{background: linear-gradient(135deg, #556B2F, #6B8E23) !important;}}
    .btn-7 {{background: linear-gradient(135deg, #8B0000, #B22222) !important;}}
    .btn-8 {{background: linear-gradient(135deg, #006400, #228B22) !important;}}
    
    /* كروت الاحصائيات */
    .stat-card {{
        padding: 20px; border-radius: 15px; text-align: center; 
        border: 3px solid #DAA520; height: 140px;
        box-shadow: 0 5px 15px rgba(218,165,32,0.3);
    }}
    .stat-card p {{color: white; font-size: 14px; margin: 0; font-weight: 700;}}
    .stat-card h2 {{color: #FFD700; font-size: 40px; font-weight: 900; margin: 8px 0 0 0;}}
    .c1 {{background: linear-gradient(135deg, #2F4F4F, #4a5f5f);}}
    .c2 {{background: linear-gradient(135deg, #1e3a5f, #2a4a7c);}}
    .c3 {{background: linear-gradient(135deg, #006400, #228B22);}}
    .c4 {{background: linear-gradient(135deg, #8B0000, #B22222);}}
</style>
""", unsafe_allow_html=True)

if 'page' not in st.session_state: st.session_state.page = "الرئيسية"

# الشريط العلوي المتحرك
st.markdown("<div class='marquee'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>", unsafe_allow_html=True)

# الهيدر + التاريخ
st.markdown(f"""
<div class='header'>
    <h1>⚖️ إدارة القضايا</h1>
    <p>📅 {today}</p>
</div>
""", unsafe_allow_html=True)

# الصفحة الرئيسية
if st.session_state.page == "الرئيسية":
    st.markdown("<h3 style='color:#FFD700; text-align:center'>📊 ملخص الإحصائيات</h3>", unsafe_allow_html=True)
    
    total_cases = len(data["cases"])
    col1,col2,col3,col4 = st.columns(4)
    with col1: st.markdown(f"<div class='stat-card c1'><p>📁 القضايا المتداولة</p><h2>{total_cases}</h2></div>", unsafe_allow_html=True)
    with col2: st.markdown("<div class='stat-card c2'><p>📅 الجلسات القادمة</p><h2>18</h2></div>", unsafe_allow_html=True)
    with col3: st.markdown("<div class='stat-card c3'><p>✅ أحكام لصالح</p><h2>42</h2></div>", unsafe_allow_html=True)
    with col4: st.markdown("<div class='stat-card c4'><p>❌ أحكام ضد</p><h2>7</h2></div>", unsafe_allow_html=True)
    
    st.markdown("<hr style='border:1px solid #DAA520; margin:20px 0'>", unsafe_allow_html=True)
    st.markdown("<h3 style='color:#FFD700; text-align:center'>الأقسام</h3>", unsafe_allow_html=True)
    
    # 2 فوق
    r1c1, r1c2 = st.columns(2)
    with r1c1:
        st.markdown("<div class='section-box'>", unsafe_allow_html=True)
        if st.button("📝", use_container_width=True, key="b1"): st.session_state.page = "تسجيل"
        st.markdown("<div class='section-name'>تسجيل القضايا</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
    with r1c2:
        st.markdown("<div class='section-box'>", unsafe_allow_html=True)
        if st.button("📊", use_container_width=True, key="b2"): st.session_state.page = "حصر"
        st.markdown("<div class='section-name'>الحصر العام</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    # 4 في النص
    r2c1, r2c2, r2c3, r2c4 = st.columns(4)
    with r2c1:
        st.markdown("<div class='section-box'>", unsafe_allow_html=True)
        if st.button("🔍", use_container_width=True, key="b3"): st.session_state.page = "بحث"
        st.markdown("<div class='section-name'>البحث</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
    with r2c2:
        st.markdown("<div class='section-box'>", unsafe_allow_html=True)
        if st.button("🔔", use_container_width=True, key="b4"): st.session_state.page = "تنبيهات"
        st.markdown("<div class='section-name'>التنبيهات</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
    with r2c3:
        st.markdown("<div class='section-box'>", unsafe_allow_html=True)
        if st.button("📑", use_container_width=True, key="b5"): st.session_state.page = "تقارير"
        st.markdown("<div class='section-name'>التقارير</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
    with r2c4:
        st.markdown("<div class='section-box'>", unsafe_allow_html=True)
        if st.button("🗃️", use_container_width=True, key="b6"): st.session_state.page = "ارشيف"
        st.markdown("<div class='section-name'>الأرشيف</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    # 2 تحت
    r3c1, r3c2 = st.columns(2)
    with r3c1:
        st.markdown("<div class='section-box'>", unsafe_allow_html=True)
        if st.button("📚", use_container_width=True, key="b7"): st.session_state.page = "مكتبة"
        st.markdown("<div class='section-name'>المكتبة القانونية</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
    with r3c2:
        st.markdown("<div class='section-box'>", unsafe_allow_html=True)
        if st.button("📈", use_container_width=True, key="b8"): st.session_state.page = "احصائيات"
        st.markdown("<div class='section-name'>الإحصائيات</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

# صفحة التسجيل مع الحفظ
elif st.session_state.page == "تسجيل":
    st.markdown("<h2 style='color:#FFD700'>📝 تسجيل القضايا</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    
    with st.form("form"):
        نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
        محكمة = st.text_input("اسم المحكمة")
        رقم = st.text_input("رقم القضية")
        سنة = st.text_input("السنة")
        if st.form_submit_button("💾 حفظ القضية"):
            data["cases"].append({"نوع": نوع, "محكمة": محكمة, "رقم": رقم, "سنة": سنة})
            save_data(data)
            st.success("✅ تم الحفظ بنجاح وسيتم الاحتفاظ به دائماً")

# صفحة الحصر
elif st.session_state.page == "حصر":
    st.markdown("<h2 style='color:#FFD700'>📊 الحصر العام</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    
    if data["cases"]:
        df = pd.DataFrame(data["cases"])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("لا توجد قضايا مسجلة بعد")

# باقي الصفحات...
elif st.session_state.page == "تنبيهات":
    st.markdown("<h2 style='color:#FFD700'>🔔 التنبيهات</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
