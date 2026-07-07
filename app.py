import streamlit as st
import pandas as pd
from datetime import datetime

# إعداد الصفحة لتكون بملء الشاشة وبشكل فخم
st.set_page_config(page_title="نظام إدارة القضايا القانونية", layout="wide")

# التصميم (CSS المخصص للألوان والخطوط)
st.markdown("""
    <style>
    .main { background-color: #001f3f; color: white; }
    .stApp { background-color: #001f3f; }
    h1 { color: #FFD700; text-align: center; font-family: 'Arial'; }
    .footer { 
        position: fixed; bottom: 0; width: 100%; text-align: center; 
        color: #FFD700; font-size: 12px; animation: blinker 2s linear infinite;
    }
    @keyframes blinker { 50% { opacity: 0; } }
    </style>
    """, unsafe_allow_html=True)

# العنوان واللوجو
st.image("path_to_scales_logo.png", width=100) # يفضل وضع صورة الميزان هنا
st.title("إدارة القضايا")

# القائمة الجانبية (الأقسام)
menu = ["تسجيل القضايا", "الحصر العام", "التقارير", "الاحصائيات", "التنبيهات", "الارشيف", "المكتبة القانونية", "البحث عن دعوى"]
choice = st.sidebar.selectbox("القائمة الرئيسية", menu)

# الفوتر (التوقيع)
st.markdown('<div class="footer">مع تحيات / وليد شعبان حماد - الادارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>', unsafe_allow_html=True)

# المنطق الأساسي للتنقل
if choice == "تسجيل القضايا":
    st.header("تسجيل القضايا")
    # هنا سيتم وضع الفورم الخاص بتسجيل الدعاوى (الاستئناف، الطعن، إلخ)
    with st.form("case_registration"):
        col1, col2 = st.columns(2)
        with col1:
            case_type = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
            court = st.selectbox("المحكمة", ["الابتدائية", "الاستئناف", "النقض", "الإدارية", "القضاء الإدارى", "الإدارية العليا"])
        with col2:
            case_num = st.text_input("رقم الدعوى")
            year = st.number_input("السنة القضائية", min_value=2000, max_value=2050)
        
        submit = st.form_submit_button("حفظ القضية")
        if submit:
            st.success("تم حفظ القضية بنجاح!")
            
