import streamlit as st
import pandas as pd

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

# CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    html, body {font-family: 'Cairo', sans-serif;}
    .stApp {background: #0a1a3a;}
    
    /* الهيدر */
    h1 {color: #FFD700 !important; text-align: center; font-size: 32px;}
    
    /* ازرار الاقسام فوق */
    .section-btn {
        background: #0d2347; 
        color: white; 
        padding: 10px 15px; 
        border-radius: 8px; 
        border: 2px solid #DAA520;
        margin: 3px;
        font-weight: bold;
        text-align: center;
    }
    .section-btn:hover {background: #1e3a5f; border-color: #FFD700;}
    .section-btn.active {background: #1e3a5f; border-color: #FFD700; color: #FFD700;}
    
    /* الكروت الصغيرة */
    .small-card {
        padding: 10px; border-radius: 10px; text-align: center; 
        height: 100px; border: 2px solid #DAA520;
    }
    .small-card p {color: white; font-size: 11px; margin: 0;}
    .small-card h2 {color: white; font-size: 24px; font-weight: 700; margin: 3px 0 0 0;}
    .c1 {background: #2F4F4F;}
    .c2 {background: #1e3a5f;}
    .c3 {background: #006400;}
    .c4 {background: #8B0000;}
    
    .footer {
        position: fixed; bottom: 0; width: 100%; text-align: center;
        background: rgba(7,20,38,0.95); padding: 8px;
        color: #00BFFF; font-weight: 700; font-size: 11px;
        border-top: 2px solid #DAA520;
    }
</style>
""", unsafe_allow_html=True)

# تهيئة
if 'page' not in st.session_state: st.session_state.page = "الحصر العام"

# العنوان
st.markdown("<h1>⚖️ إدارة القضايا</h1>", unsafe_allow_html=True)

# الاقسام فوق كأزرار
c1,c2,c3,c4,c5,c6,c7,c8 = st.columns(8)
if c1.button("📝 تسجيل القضايا", use_container_width=True): st.session_state.page = "تسجيل القضايا"
if c2.button("📊 الحصر العام", use_container_width=True): st.session_state.page = "الحصر العام"
if c3.button("🔍 البحث", use_container_width=True): st.session_state.page = "البحث"
if c4.button("🔔 التنبيهات", use_container_width=True): st.session_state.page = "التنبيهات"
if c5.button("📑 التقارير", use_container_width=True): st.session_state.page = "التقارير"
if c6.button("🗃️ الأرشيف", use_container_width=True): st.session_state.page = "الأرشيف"
if c7.button("📚 المكتبة", use_container_width=True): st.session_state.page = "المكتبة"
if c8.button("📈 الإحصائيات", use_container_width=True): st.session_state.page = "الإحصائيات"

st.markdown("---")

# المحتوى حسب القسم
if st.session_state.page == "الحصر العام":
    st.subheader("📊 الحصر العام")
    
    # الكروت الصغيرة
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
    st.dataframe(df, use_container_width=True, height=250)

elif st.session_state.page == "تسجيل القضايا":
    st.subheader("📝 تسجيل القضايا")
    with st.form("form"):
        نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
        محكمة = st.text_input("اسم المحكمة")
        رقم = st.text_input("رقم القضية")
        سنة = st.text_input("السنة")
        if st.form_submit_button("💾 حفظ القضية"): st.success("✅ تم الحفظ")

elif st.session_state.page == "التنبيهات":
    st.subheader("🔔 التنبيهات")
    st.warning("⚠️ جلسة 1234/2023 يوم 20/1")
    st.info("📅 جلسة 5678/2024 يوم 25/1")

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

# الفوتر
st.markdown("<div class='footer'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>", unsafe_allow_html=True)
