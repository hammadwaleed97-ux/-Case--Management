import streamlit as st
import pandas as pd
from datetime import datetime
import json

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

# تحميل البيانات
if 'cases' not in st.session_state: st.session_state.cases = []

# CSS طبق الأصل من الصورة
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    html, body, [class*="css"] {font-family: 'Cairo', sans-serif;}
    
    .stApp {
        background-image: linear-gradient(rgba(5,15,30,0.95), rgba(5,15,30,0.95)), url('https://i.imgur.com/3ZQ3ZQp.jpg');
        background-size: cover;
    }
    
    /* السايدبار */
    [data-testid="stSidebar"] {
        background: #071426; 
        border-right: 2px solid #DAA520;
        padding-top: 20px;
    }
    .sidebar-btn {
        background: #0d2347; 
        color: white; 
        padding: 12px; 
        border-radius: 8px; 
        margin: 8px 0; 
        border-left: 4px solid #DAA520;
        font-weight: bold;
    }
    .sidebar-btn.active {background: #1e3a5f; border-left: 4px solid #00BFFF;}
    
    /* الهيدر */
    h1 {color: #FFD700 !important; text-align: center; font-size: 38px; font-weight: 700;}
    
    /* الكروت ال 4 */
    .stat-card {
        padding: 20px; border-radius: 12px; text-align: center; 
        box-shadow: 0 4px 15px rgba(0,0,0,0.3); margin: 5px;
    }
    .stat-card h4 {color: white; font-size: 15px; margin-bottom: 10px;}
    .stat-card h2 {color: white; font-size: 36px; font-weight: 700;}
    .red {background: linear-gradient(135deg, #8B0000, #B22222);}
    .green {background: linear-gradient(135deg, #006400, #228B22);}
    .blue {background: linear-gradient(135deg, #1e3a5f, #2a4a7c);}
    .dark {background: linear-gradient(135deg, #2F4F4F, #4a5f5f);}
    
    /* الجدول */
    .dataframe {background: #0d2347 !important; border: 2px solid #DAA520;}
    .dataframe th {background: #1e3a5f !important; color: #FFD700 !important;}
    .dataframe td {color: white !important;}
    
    /* الفوتر */
    .footer {
        position: fixed; bottom: 0; width: 100%; text-align: center;
        background: rgba(7,20,38,0.9); padding: 10px;
        color: #00BFFF; font-weight: 700; font-size: 13px;
        border-top: 2px solid #DAA520;
    }
</style>
""", unsafe_allow_html=True)

# السايدبار
with st.sidebar:
    st.markdown("<h2 style='text-align:center; color:#DAA520'>⚖️ إدارة القضايا</h2>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-btn active'>📝 تسجيل القضايا</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-btn'>📊 الحصر العام</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-btn'>🔔 التنبيهات</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-btn'>📑 التقارير</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-btn'>🗃️ الأرشيف</div>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-btn'>📚 المكتبة القانونية</div>", unsafe_allow_html=True)

# الهيدر
st.markdown("<h1>⚖️ إدارة القضايا</h1>", unsafe_allow_html=True)

# ال 4 كروت
c1,c2,c3,c4 = st.columns(4)
c1.markdown("<div class='stat-card dark'><h4>📊 عدد القضايا المتداولة</h4><h2>105</h2></div>", unsafe_allow_html=True)
c2.markdown("<div class='stat-card blue'><h4>📅 عدد الجلسات القادمة</h4><h2>18</h2></div>", unsafe_allow_html=True)
c3.markdown("<div class='stat-card green'><h4>✅ أحكام لصالح الهيئة</h4><h2>42</h2></div>", unsafe_allow_html=True)
c4.markdown("<div class='stat-card red'><h4>❌ أحكام ضد الهيئة</h4><h2>7</h2></div>", unsafe_allow_html=True)

st.markdown("<h3 style='color:#DAA520'>الحصر العام</h3>", unsafe_allow_html=True)

# الجدول
data = {
    'م': [1,2,3,4],
    'رقم القضية': ['1234/2023', '5678/2024', '9101/2023', '1121/2024'],
    'الخصوم': ['الهيئة ضد أحمد', 'محمد ضد الهيئة', 'الهيئة ضد شركة', 'سعيد ضد الهيئة'],
    'المحكمة': ['الابتدائية', 'الاستئناف', 'الإدارية', 'القضاء الإداري'],
    'آخر جلسة': ['2024-01-15', '2024-01-20', '2024-01-10', '2024-01-25'],
    'الحالة': ['متداولة', 'متداولة', 'متداولة', 'متداولة']
}
df = pd.DataFrame(data)
st.dataframe(df, use_container_width=True, height=300)

# الفوتر
st.markdown("<div class='footer'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>", unsafe_allow_html=True)
