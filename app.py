import streamlit as st
import pandas as pd

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️", initial_sidebar_state="expanded")

# CSS للسايدبار الثابت والكروت الصغيرة
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    html, body {font-family: 'Cairo', sans-serif;}
    .stApp {background: #0a1a3a;}
    
    /* السايدبار ثابت وازرق غامق */
    [data-testid="stSidebar"] {
        background: #071426 !important; 
        border-right: 3px solid #DAA520;
        min-width: 260px !important;
    }
    .sidebar-title {color: #FFD700; text-align: center; font-size: 22px; font-weight: 700; padding: 15px;}
    div[data-testid="stSidebar"] button {
        background: #0d2347 !important; 
        color: white !important; 
        width: 100%; 
        margin: 6px 0; 
        border: 1px solid #DAA520;
        text-align: right;
        font-weight: bold;
        border-radius: 8px;
        padding: 10px;
    }
    div[data-testid="stSidebar"] button:hover {
        background: #1e3a5f !important;
        border-right: 4px solid #FFD700;
    }
    
    /* الهيدر */
    h1 {color: #FFD700 !important; text-align: center; font-size: 32px;}
    
    /* الكروت الصغيرة جوة كل قسم */
    .small-card {
        padding: 10px; border-radius: 10px; text-align: center; 
        height: 100px; border: 2px solid #DAA520; margin: 5px;
    }
    .small-card p {color: white; font-size: 11px; margin: 0;}
    .small-card h2 {color: white; font-size: 24px; font-weight: 700; margin: 3px 0 0 0;}
    .c1 {background: #2F4F4F;}
    .c2 {background: #1e3a5f;}
    .c3 {background: #006400;}
    .c4 {background: #8B0000;}
    
    /* الفوتر */
    .footer {
        position: fixed; bottom: 0; width: 100%; text-align: center;
        background: rgba(7,20,38,0.95); padding: 8px;
        color: #00BFFF; font-weight: 700; font-size: 11px;
        border-top: 2px solid #DAA520; z-index: 999;
    }
</style>
""", unsafe_allow_html=True)

# تهيئة الصفحة
if 'page' not in st.session_state: st.session_state.page = "الحصر العام"

# السايدبار الثابت
with st.sidebar:
    st.markdown("<div class='sidebar-title'>⚖️ إدارة القضايا</div>", unsafe_allow_html=True)
    
    if st.button("📝 تسجيل القضايا"): st.session_state.page = "تسجيل القضايا"
    if st.button("📊 الحصر العام"): st.session_state.page = "الحصر العام"
    if st.button("🔍 البحث"): st.session_state.page = "البحث"
    if st.button("🔔 التنبيهات"): st.session_state.page = "التنبيهات"
    if st.button("📑 التقارير"): st.session_state.page = "التقارير"
    if st.button("🗃️ الأرشيف"): st.session_state.page = "الأرشيف"
    if st.button("📚 المكتبة القانونية"): st.session_state.page = "المكتبة"
    if st.button("📈 الإحصائيات"): st.session_state.page = "الإحصائيات"

# العنوان
st.markdown("<h1>⚖️ إدارة القضايا</h1>", unsafe_allow_html=True)

# المحتوى حسب القسم اللي ضغطت عليه
if st.session_state.page == "الحصر العام":
    st.subheader("📊 الحصر العام")
    
    # الكروت الصغيرة
    col1,col2,col3,col4 = st.columns(4)
    with col1: st.markdown("<div class='small-card c1'><p>عدد القضايا</p><h2>105</h2></div>", unsafe_allow_html=True)
    with col2: st.markdown("<div class='small-card c2'><p>الجلسات القادمة</p><h2>18</h2></div>", unsafe_allow_html=True)
    with col3: st.markdown("<div class='small-card c3'><p>أحكام لصالح</p><h2>42</h2></div>", unsafe_allow_html=True)
    with col4: st.markdown("<div class='small-card c4'><p>أحكام ضد</p><h2>7</h2></div>", unsafe_allow_html=True)
    
    df = pd.DataFrame({
        'م': [1,2,3,4],
        'رقم القضية': ['1234/2023', '5678/2024', '9101/2023', '1121/2024'],
        'الخصوم': ['الهيئة ضد أحمد', 'محمد ضد الهيئة', 'الهيئة ضد شركة', 'سعيد ضد الهيئة'],
    })
    st.dataframe(df, use_container_width=True, height=250)

elif st.session_state.page == "تسجيل القضايا":
    st.subheader("📝 تسجيل القضايا")
    with st.form("form"):
        st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
        st.text_input("اسم المحكمة")
        st.text_input("رقم القضية")
        if st.form_submit_button("💾 حفظ القضية"): st.success("تم الحفظ")

elif st.session_state.page == "التنبيهات":
    st.subheader("🔔 التنبيهات")
    st.warning("⚠️ جلسة 1234/2023 بعد 3 ايام")

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

# الفوتر
st.markdown("<div class='footer'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>", unsafe_allow_html=True)
