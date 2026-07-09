import streamlit as st
import pandas as pd

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    html, body {font-family: 'Cairo', sans-serif;}
    .stApp {background: #0a1a3a;}
    
    h1 {color: #FFD700 !important; text-align: center; font-size: 32px;}
    
    /* ازرار الاقسام ملونة */
    div[data-testid="stButton"] > button {
        width: 100%;
        padding: 14px; 
        border-radius: 10px; 
        border: 2px solid #DAA520;
        font-weight: bold; 
        font-size: 13px;
        color: white !important;
        height: 60px;
    }
    /* الوان كل زرار */
    div[data-testid="stButton"] > button[kind="primary"] {background: linear-gradient(135deg, #1e3a5f, #2a4a7c) !important;}
    div[data-testid="stButton"] > button[kind="secondary"] {background: linear-gradient(135deg, #2F4F4F, #4a5f5f) !important;}
    
    /* هنلون بالترتيب */
    .stButton:nth-child(1) button {background: linear-gradient(135deg, #1e3a5f, #2a4a7c) !important;} /* تسجيل */
    .stButton:nth-child(2) button {background: linear-gradient(135deg, #2F4F4F, #4a5f5f) !important;} /* الحصر */
    .stButton:nth-child(3) button {background: linear-gradient(135deg, #8B4513, #A0522D) !important;} /* البحث */
    .stButton:nth-child(4) button {background: linear-gradient(135deg, #B8860B, #DAA520) !important; color: black !important;} /* التنبيهات */
    .stButton:nth-child(5) button {background: linear-gradient(135deg, #4B0082, #6A5ACD) !important;} /* التقارير */
    .stButton:nth-child(6) button {background: linear-gradient(135deg, #556B2F, #6B8E23) !important;} /* الارشيف */
    .stButton:nth-child(7) button {background: linear-gradient(135deg, #8B0000, #B22222) !important;} /* المكتبة */
    .stButton:nth-child(8) button {background: linear-gradient(135deg, #006400, #228B22) !important;} /* الاحصائيات */
    
    /* الكروت الصغيرة */
    .small-card {
        padding: 10px; border-radius: 10px; text-align: center; 
        height: 100px; border: 2px solid #DAA520;
    }
    .small-card p {color: white; font-size: 11px; margin: 0;}
    .small-card h2 {color: white; font-size: 24px; font-weight: 700; margin: 3px 0 0 0;}
    .c1 {background: linear-gradient(135deg, #2F4F4F, #4a5f5f);}
    .c2 {background: linear-gradient(135deg, #1e3a5f, #2a4a7c);}
    .c3 {background: linear-gradient(135deg, #006400, #228B22);}
    .c4 {background: linear-gradient(135deg, #8B0000, #B22222);}
    
    .footer {
        position: fixed; bottom: 0; width: 100%; text-align: center;
        background: rgba(7,20,38,0.95); padding: 8px;
        color: #00BFFF; font-weight: 700; font-size: 11px;
        border-top: 2px solid #DAA520;
    }
</style>
""", unsafe_allow_html=True)

if 'page' not in st.session_state: st.session_state.page = "none"

st.markdown("<h1>⚖️ إدارة القضايا</h1>", unsafe_allow_html=True)

# الصف الاول: 2 ازرار
r1c1, r1c2 = st.columns(2)
with r1c1:
    if st.button("📝 تسجيل القضايا", use_container_width=True, key="b1"): st.session_state.page = "تسجيل القضايا"
with r1c2:
    if st.button("📊 الحصر العام", use_container_width=True, key="b2"): st.session_state.page = "الحصر العام"

# الصف التاني: 4 ازرار
r2c1, r2c2, r2c3, r2c4 = st.columns(4)
with r2c1:
    if st.button("🔍 البحث", use_container_width=True, key="b3"): st.session_state.page = "البحث"
with r2c2:
    if st.button("🔔 التنبيهات", use_container_width=True, key="b4"): st.session_state.page = "التنبيهات"
with r2c3:
    if st.button("📑 التقارير", use_container_width=True, key="b5"): st.session_state.page = "التقارير"
with r2c4:
    if st.button("🗃️ الأرشيف", use_container_width=True, key="b6"): st.session_state.page = "الأرشيف"

# الصف التالت: 2 ازرار
r3c1, r3c2 = st.columns(2)
with r3c1:
    if st.button("📚 المكتبة", use_container_width=True, key="b7"): st.session_state.page = "المكتبة"
with r3c2:
    if st.button("📈 الإحصائيات", use_container_width=True, key="b8"): st.session_state.page = "الإحصائيات"

st.markdown("---")

# المحتوى
if st.session_state.page == "none":
    st.markdown("<h3 style='text-align:center; color:#DAA520'>اختر قسم من الأعلى للبدء</h3>", unsafe_allow_html=True)

elif st.session_state.page == "الحصر العام":
    st.subheader("📊 الحصر العام")
    
    # الكروت الصغيرة الملونة
    col1,col2,col3,col4 = st.columns(4)
    with col1: st.markdown("<div class='small-card c1'><p>عدد القضايا</p><h2>105</h2></div>", unsafe_allow_html=True)
    with col2: st.markdown("<div class='small-card c2'><p>الجلسات القادمة</p><h2>18</h2></div>", unsafe_allow_html=True)
    with col3: st.markdown("<div class='small-card c3'><p>أحكام لصالح</p><h2>42</h2></div>", unsafe_allow_html=True)
    with col4: st.markdown("<div class='small-card c4'><p>أحكام ضد</p><h2>7</h2></div>", unsafe_allow_html=True)
    
    st.markdown("<h4 style='color:#DAA520'>جدول الحصر</h4>", unsafe_allow_html=True)
    df = pd.DataFrame({
        'م': [1,2,3,4],
        'رقم القضية': ['1234/2023', '5678/2024', '9101/2023', '1121/2024'],
        'الخصوم': ['الهيئة ضد أحمد', 'محمد ضد الهيئة', 'الهيئة ضد شركة', 'سعيد ضد الهيئة'],
        'المحكمة': ['ابتدائية', 'استئناف', 'ادارية', 'قضاء اداري'],
    })
    st.dataframe(df, use_container_width=True, height=250, hide_index=True)

elif st.session_state.page == "تسجيل القضايا":
    st.subheader("📝 تسجيل القضايا")
    with st.form("form"):
        st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
        st.text_input("اسم المحكمة")
        st.text_input("رقم القضية")
        if st.form_submit_button("💾 حفظ القضية"): st.success("✅ تم الحفظ")

elif st.session_state.page == "التنبيهات":
    st.subheader("🔔 التنبيهات")
    st.warning("⚠️ جلسة 1234/2023 يوم 20/1")

elif st.session_state.page == "التقارير":
    st.subheader("📑 التقارير")
elif st.session_state.page == "الأرشيف":
    st.subheader("🗃️ الأرشيف")
elif st.session_state.page == "المكتبة":
    st.subheader("📚 المكتبة القانونية")
elif st.session_state.page == "الإحصائيات":
    st.subheader("📈 الإحصائيات")
elif st.session_state.page == "البحث":
    st.subheader("🔍 البحث")
    st.text_input("ابحث برقم القضية")

st.markdown("<div class='footer'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>", unsafe_allow_html=True)
