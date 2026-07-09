import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

DATA_FILE = "cases_data.json"
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return {"cases": []}
def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=4)

data = load_data()
today = datetime.now().strftime("%A, %d %B %Y")

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');
    html, body {{font-family: 'Cairo', sans-serif;}}
    .stApp {{background: #0a1a3a;}}
    
    /* الشريط المتحرك بتاع التليفزيون */
    .marquee-container {{
        background: linear-gradient(90deg, #DAA520, #FFD700);
        padding: 12px; overflow: hidden; border-radius: 8px;
        margin-bottom: 15px; box-shadow: 0 0 15px #DAA520;
    }}
    .marquee-text {{
        color: #0a1a3a; font-weight: 900; font-size: 14px;
        white-space: nowrap; display: inline-block;
        animation: scroll-left 15s linear infinite;
    }}
    @keyframes scroll-left {{
        0% {{transform: translateX(100%);}}
        100% {{transform: translateX(-100%);}}
    }}
    
    /* هيدر اصفر */
    .header-yellow {{
        background: linear-gradient(135deg, #DAA520, #FFD700);
        padding: 20px; border-radius: 15px; text-align: center;
        border: 3px solid #FFD700; margin-bottom: 20px;
    }}
    .header-yellow h1 {{color: #0a1a3a; font-size: 32px; font-weight: 900; margin: 0;}}
    .header-yellow p {{color: #0a1a3a; font-size: 14px; font-weight: 700; margin: 8px 0 0 0;}}
    
    /* عنوان الاقسام */
    .section-title {{color: #FFD700; text-align: center; font-size: 28px; font-weight: 900; margin: 20px 0;}}
    
    /* الايقونات البيضاء بحد دهبي والاسم جواها */
    div[data-testid="stButton"] > button {{
        background: white !important;
        color: #0a1a3a !important;
        width: 100%; padding: 25px 10px; border-radius: 15px; 
        border: 3px solid #DAA520; font-weight: 900; font-size: 16px;
        height: 80px; margin-bottom: 5px;
        display: flex; flex-direction: column; justify-content: center; align-items: center;
    }}
    div[data-testid="stButton"] > button:hover {{border-color: #FFD700; box-shadow: 0 0 15px #FFD700;}}
    
    .icon {{font-size: 28px; margin-bottom: 5px;}}
    .btn-text {{font-size: 14px;}}
    
    /* كروت الاحصائيات الصغيرة جنب بعض */
    .small-stat {{
        padding: 15px; border-radius: 12px; text-align: center; 
        border: 3px solid #DAA520; margin-bottom: 10px;
    }}
    .small-stat p {{color: white; font-size: 13px; margin: 0; font-weight: 700;}}
    .small-stat h2 {{color: #FFD700; font-size: 32px; font-weight: 900; margin: 5px 0 0 0;}}
    .s1 {{background: #2F4F4F;}}
    .s2 {{background: #1e3a5f;}}
    .s3 {{background: #006400;}}
    .s4 {{background: #8B0000;}}
    
    .stats-title {{color: #FFD700; font-size: 20px; font-weight: 900; margin: 15px 0;}}
</style>
""", unsafe_allow_html=True)

if 'page' not in st.session_state: st.session_state.page = "الرئيسية"

# الشريط المتحرك
st.markdown("""
<div class='marquee-container'>
    <div class='marquee-text'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>
</div>
""", unsafe_allow_html=True)

# الهيدر الاصفر
st.markdown(f"""
<div class='header-yellow'>
    <h1>⚖️ إدارة القضايا</h1>
    <p>📅 {today}</p>
</div>
""", unsafe_allow_html=True)

# الصفحة الرئيسية
if st.session_state.page == "الرئيسية":
    
    st.markdown("<div class='stats-title'>📊 ملخص الإحصائيات</div>", unsafe_allow_html=True)
    
    # الكروت الصغيرة تحت بعض
    total_cases = len(data["cases"])
    st.markdown(f"<div class='small-stat s1'><p>📁 القضايا المتداولة</p><h2>{total_cases}</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='small-stat s2'><p>📅 الجلسات القادمة</p><h2>18</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='small-stat s3'><p>✅ أحكام لصالح</p><h2>42</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='small-stat s4'><p>❌ أحكام ضد</p><h2>7</h2></div>", unsafe_allow_html=True)
    
    st.markdown("<div class='section-title'>الأقسام</div>", unsafe_allow_html=True)
    
    # الاقسام كل واحد في سطر والاسم جوه
    if st.button("📝\nتسجيل القضايا", use_container_width=True, key="b1"): st.session_state.page = "تسجيل"
    if st.button("📊\nالحصر العام", use_container_width=True, key="b2"): st.session_state.page = "حصر"
    if st.button("🔍\nالبحث", use_container_width=True, key="b3"): st.session_state.page = "بحث"
    if st.button("🔔\nالتنبيهات", use_container_width=True, key="b4"): st.session_state.page = "تنبيهات"
    if st.button("📑\nالتقارير", use_container_width=True, key="b5"): st.session_state.page = "تقارير"
    if st.button("🗃️\nالأرشيف", use_container_width=True, key="b6"): st.session_state.page = "ارشيف"
    if st.button("📚\nالمكتبة القانونية", use_container_width=True, key="b7"): st.session_state.page = "مكتبة"
    if st.button("📈\nالإحصائيات", use_container_width=True, key="b8"): st.session_state.page = "احصائيات"

# صفحة التسجيل
elif st.session_state.page == "تسجيل":
    st.markdown("<h2 style='color:#FFD700'>📝 تسجيل القضايا</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    with st.form("form"):
        نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
        محكمة = st.text_input("اسم المحكمة")
        رقم = st.text_input("رقم القضية")
        if st.form_submit_button("💾 حفظ"):
            data["cases"].append({"نوع": نوع, "محكمة": محكمة, "رقم": رقم})
            save_data(data)
            st.success("✅ تم الحفظ")

# صفحة الحصر
elif st.session_state.page == "حصر":
    st.markdown("<h2 style='color:#FFD700'>📊 الحصر العام</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    if data["cases"]:
        st.dataframe(pd.DataFrame(data["cases"]), use_container_width=True, hide_index=True)
    else:
        st.info("لا توجد قضايا مسجلة")

# باقي الصفحات...
elif st.session_state.page == "تنبيهات":
    st.markdown("<h2 style='color:#FFD700'>🔔 التنبيهات</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
