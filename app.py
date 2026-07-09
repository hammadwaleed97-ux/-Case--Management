import streamlit as st
import pandas as pd

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️", initial_sidebar_state="expanded")

# CSS مضغوط ومحترف
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    html, body {font-family: 'Cairo', sans-serif;}
    
    .stApp {background: #0a1a3a;}
    
    /* السايدبار جنب */
    [data-testid="stSidebar"] {
        background: #071426; 
        border-right: 3px solid #DAA520;
        width: 250px !important;
    }
    .sidebar-title {color: #FFD700; text-align: center; font-size: 20px; font-weight: 700; padding: 10px;}
    .sidebar-item {
        background: #0d2347; color: white; padding: 10px; margin: 5px 10px; 
        border-radius: 6px; border-right: 3px solid #DAA520; font-size: 14px;
    }
    .sidebar-item:hover {background: #1e3a5f;}
    
    /* الهيدر */
    h1 {color: #FFD700 !important; text-align: center; font-size: 32px;}
    
    /* الكروت جنب بعض وصغيرة */
    .metric-card {
        padding: 12px; border-radius: 10px; text-align: center; 
        height: 120px; display: flex; flex-direction: column; justify-content: center;
        border: 2px solid #DAA520;
    }
    .metric-card h4 {color: white; font-size: 13px; margin: 0;}
    .metric-card h2 {color: white; font-size: 32px; font-weight: 700; margin: 5px 0 0 0;}
    .c1 {background: linear-gradient(135deg, #2F4F4F, #4a5f5f);}
    .c2 {background: linear-gradient(135deg, #1e3a5f, #2a4a7c);}
    .c3 {background: linear-gradient(135deg, #006400, #228B22);}
    .c4 {background: linear-gradient(135deg, #8B0000, #B22222);}
    
    /* الجدول */
    .dataframe {font-size: 12px;}
    
    /* الفوتر */
    .footer {
        position: fixed; bottom: 0; width: 100%; text-align: center;
        background: rgba(7,20,38,0.9); padding: 8px;
        color: #00BFFF; font-weight: 700; font-size: 12px;
        border-top: 2px solid #DAA520; z-index: 999;
    }
</style>
""", unsafe_allow_html=True)

# السايدبار بالجنب
with st.sidebar:
    st.markdown("<div class='sidebar-title'>⚖️ إدارة القضايا</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-item'>📝 تسجيل القضايا</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-item'>📊 الحصر العام</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-item'>🔍 البحث</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-item'>🔔 التنبيهات</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-item'>📑 التقارير</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-item'>🗃️ الأرشيف</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-item'>📚 المكتبة القانونية</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-item'>📈 الإحصائيات</div>", unsafe_allow_html=True)

# العنوان
st.markdown("<h1>⚖️ إدارة القضايا</h1>", unsafe_allow_html=True)

# ال 4 كروت جنب بعض
col1,col2,col3,col4 = st.columns(4)
with col1:
    st.markdown("<div class='metric-card c1'>📊 <h4>عدد القضايا المتداولة</h4><h2>105</h2></div>", unsafe_allow_html=True)
with col2:
    st.markdown("<div class='metric-card c2'>📅 <h4>عدد الجلسات القادمة</h4><h2>18</h2></div>", unsafe_allow_html=True)
with col3:
    st.markdown("<div class='metric-card c3'>✅ <h4>أحكام لصالح الهيئة</h4><h2>42</h2></div>", unsafe_allow_html=True)
with col4:
    st.markdown("<div class='metric-card c4'>❌ <h4>أحكام ضد الهيئة</h4><h2>7</h2></div>", unsafe_allow_html=True)

st.markdown("<h3 style='color:#DAA520; margin-top:20px'>الحصر العام</h3>", unsafe_allow_html=True)

# الجدول
data = {
    'م': [1,2,3,4],
    'رقم القضية': ['1234/2023', '5678/2024', '9101/2023', '1121/2024'],
    'الخصوم': ['الهيئة ضد أحمد', 'محمد ضد الهيئة', 'الهيئة ضد شركة', 'سعيد ضد الهيئة'],
    'المحكمة': ['ابتدائية', 'استئناف', 'ادارية', 'قضاء اداري'],
    'آخر جلسة': ['2024-01-15', '2024-01-20', '2024-01-10', '2024-01-25'],
}
df = pd.DataFrame(data)
st.dataframe(df, use_container_width=True, height=250)

# الفوتر
st.markdown("<div class='footer'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>", unsafe_allow_html=True)
