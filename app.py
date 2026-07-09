import streamlit as st
import pandas as pd

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

# CSS مطابق للصورة
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    html, body {font-family: 'Cairo', sans-serif;}
    .stApp {background: #0a1a3a;}
    
    /* الهيدر */
    h1 {color: #FFD700 !important; text-align: center; font-size: 35px;}
    
    /* التابس فوق */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: #071426;
        padding: 10px;
        border-radius: 10px;
        border: 2px solid #DAA520;
    }
    .stTabs [data-baseweb="tab"] {
        background: #0d2347;
        color: white !important;
        border-radius: 8px;
        padding: 8px 15px;
        font-weight: bold;
        border: 1px solid #DAA520;
    }
    .stTabs [aria-selected="true"] {
        background: #1e3a5f !important;
        color: #FFD700 !important;
    }
    
    /* الكروت مضغوطة */
    .metric-card {
        padding: 12px; border-radius: 12px; text-align: center; 
        height: 115px; border: 2px solid #DAA520;
    }
    .metric-card p {color: white; font-size: 13px; margin: 0;}
    .metric-card h2 {color: white; font-size: 30px; font-weight: 700; margin: 5px 0 0 0;}
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

# العنوان
st.markdown("<h1>⚖️ إدارة القضايا</h1>", unsafe_allow_html=True)

# الاقسام ظاهرة فوق كـ Tabs
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "📝 تسجيل القضايا", 
    "📊 الحصر العام", 
    "🔍 البحث", 
    "🔔 التنبيهات", 
    "📑 التقارير", 
    "🗃️ الأرشيف", 
    "📚 المكتبة", 
    "📈 الإحصائيات"
])

# الكروت في كل الصفحات
with tab1:
    st.subheader("تسجيل قضية جديدة")
    with st.form("form"):
        نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
        محكمة = st.text_input("اسم المحكمة")
        رقم = st.text_input("رقم القضية")
        if st.form_submit_button("💾 حفظ"):
            st.success("تم الحفظ")

with tab2:
    # الكروت ال 4
    col1,col2,col3,col4 = st.columns(4)
    with col1: st.markdown("<div class='metric-card c1'><p>📊 عدد القضايا المتداولة</p><h2>105</h2></div>", unsafe_allow_html=True)
    with col2: st.markdown("<div class='metric-card c2'><p>📅 عدد الجلسات القادمة</p><h2>18</h2></div>", unsafe_allow_html=True)
    with col3: st.markdown("<div class='metric-card c3'><p>✅ أحكام لصالح الهيئة</p><h2>42</h2></div>", unsafe_allow_html=True)
    with col4: st.markdown("<div class='metric-card c4'><p>❌ أحكام ضد الهيئة</p><h2>7</h2></div>", unsafe_allow_html=True)
    
    st.markdown("<h3 style='color:#DAA520; margin-top:15px'>الحصر العام</h3>", unsafe_allow_html=True)
    df = pd.DataFrame({
        'م': [1,2,3,4],
        'رقم القضية': ['1234/2023', '5678/2024', '9101/2023', '1121/2024'],
        'الخصوم': ['الهيئة ضد أحمد', 'محمد ضد الهيئة', 'الهيئة ضد شركة', 'سعيد ضد الهيئة'],
    })
    st.dataframe(df, use_container_width=True, height=250)

with tab3:
    st.subheader("🔍 البحث عن دعوى")
    st.text_input("ابحث برقم القضية او اسم المدعي")

with tab4:
    st.subheader("🔔 التنبيهات")
    st.warning("جلسة 1234/2023 بعد 3 ايام")

with tab5:
    st.subheader("📑 التقارير")

with tab6:
    st.subheader("🗃️ الأرشيف")

with tab7:
    st.subheader("📚 المكتبة القانونية")

with tab8:
    st.subheader("📈 الإحصائيات")

# الفوتر
st.markdown("<div class='footer'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>", unsafe_allow_html=True)
