import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

# تاريخ اليوم
today = datetime.now().strftime("%A, %d %B %Y")

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');
    html, body {{font-family: 'Cairo', sans-serif;}}
    .stApp {{background: linear-gradient(180deg, #0a1a3a 0%, #122c5e 100%);}}
    
    /* الهيدر الفخم */
    .header {{
        background: linear-gradient(135deg, #DAA520, #FFD700);
        padding: 15px; border-radius: 15px; text-align: center;
        border: 3px solid #DAA520; margin-bottom: 20px;
        box-shadow: 0 0 20px rgba(218,165,32,0.5);
    }}
    .header h1 {{color: #0a1a3a; font-size: 36px; font-weight: 900; margin: 0;}}
    .header p {{color: #0a1a3a; font-size: 16px; font-weight: 700; margin: 5px 0 0 0;}}
    
    /* ازرار الاقسام فخمة */
    div[data-testid="stButton"] > button {{
        width: 100%; padding: 16px; border-radius: 12px; 
        border: 2px solid #DAA520; font-weight: 900; font-size: 14px;
        color: white !important; height: 70px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.3);
        transition: 0.3s;
    }}
    div[data-testid="stButton"] > button:hover {{
        transform: translateY(-3px);
        box-shadow: 0 6px 15px rgba(218,165,32,0.6);
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
    
    /* كروت الاحصائيات الفخمة */
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
    
    /* كروت داخل الصفحات */
    .page-card {{
        background: #0d2347; padding: 20px; border-radius: 12px;
        border: 2px solid #DAA520; margin: 10px 0;
    }}
    
    .footer {{
        position: fixed; bottom: 0; width: 100%; text-align: center;
        background: rgba(7,20,38,0.95); padding: 10px;
        color: #FFD700; font-weight: 700; font-size: 12px;
        border-top: 3px solid #DAA520;
    }}
</style>
""", unsafe_allow_html=True)

if 'page' not in st.session_state: st.session_state.page = "الرئيسية"

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
    
    col1,col2,col3,col4 = st.columns(4)
    with col1: st.markdown("<div class='stat-card c1'><p>📁 القضايا المتداولة</p><h2>105</h2></div>", unsafe_allow_html=True)
    with col2: st.markdown("<div class='stat-card c2'><p>📅 الجلسات القادمة</p><h2>18</h2></div>", unsafe_allow_html=True)
    with col3: st.markdown("<div class='stat-card c3'><p>✅ أحكام لصالح</p><h2>42</h2></div>", unsafe_allow_html=True)
    with col4: st.markdown("<div class='stat-card c4'><p>❌ أحكام ضد</p><h2>7</h2></div>", unsafe_allow_html=True)
    
    st.markdown("<hr style='border:1px solid #DAA520; margin:20px 0'>", unsafe_allow_html=True)
    st.markdown("<h3 style='color:#FFD700; text-align:center'>اختر القسم</h3>", unsafe_allow_html=True)
    
    # 2 فوق
    r1c1, r1c2 = st.columns(2)
    with r1c1:
        if st.button("📝 تسجيل القضايا", use_container_width=True, key="b1"): st.session_state.page = "تسجيل"
    with r1c2:
        if st.button("📊 الحصر العام", use_container_width=True, key="b2"): st.session_state.page = "حصر"
    
    # 4 في النص
    r2c1, r2c2, r2c3, r2c4 = st.columns(4)
    with r2c1:
        if st.button("🔍 البحث", use_container_width=True, key="b3"): st.session_state.page = "بحث"
    with r2c2:
        if st.button("🔔 التنبيهات", use_container_width=True, key="b4"): st.session_state.page = "تنبيهات"
    with r2c3:
        if st.button("📑 التقارير", use_container_width=True, key="b5"): st.session_state.page = "تقارير"
    with r2c4:
        if st.button("🗃️ الأرشيف", use_container_width=True, key="b6"): st.session_state.page = "ارشيف"
    
    # 2 تحت
    r3c1, r3c2 = st.columns(2)
    with r3c1:
        if st.button("📚 المكتبة", use_container_width=True, key="b7"): st.session_state.page = "مكتبة"
    with r3c2:
        if st.button("📈 الإحصائيات", use_container_width=True, key="b8"): st.session_state.page = "احصائيات"

# صفحة الحصر العام
elif st.session_state.page == "حصر":
    st.markdown("<h2 style='color:#FFD700'>📊 الحصر العام</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    
    col1,col2,col3 = st.columns(3)
    with col1: st.markdown("<div class='page-card'><h4 style='color:#DAA520'>إجمالي القضايا</h4><h2 style='color:white'>105</h2></div>", unsafe_allow_html=True)
    with col2: st.markdown("<div class='page-card'><h4 style='color:#DAA520'>تم الحكم فيها</h4><h2 style='color:white'>49</h2></div>", unsafe_allow_html=True)
    with col3: st.markdown("<div class='page-card'><h4 style='color:#DAA520'>متداولة</h4><h2 style='color:white'>56</h2></div>", unsafe_allow_html=True)
    
    st.dataframe(pd.DataFrame({
        'رقم القضية': ['1234/2023', '5678/2024'],
        'الخصوم': ['الهيئة ضد أحمد', 'محمد ضد الهيئة'],
        'الحالة': ['متداولة', 'متداولة']
    }), use_container_width=True, hide_index=True)

# صفحة التنبيهات
elif st.session_state.page == "تنبيهات":
    st.markdown("<h2 style='color:#FFD700'>🔔 التنبيهات</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    
    st.markdown("<div class='page-card'><h4 style='color:#FFD700'>⚠️ جلسات هذا الاسبوع</h4><p style='color:white'>1234/2023 - يوم 20/1/2026</p></div>", unsafe_allow_html=True)

# باقي الصفحات
elif st.session_state.page == "تسجيل":
    st.markdown("<h2 style='color:#FFD700'>📝 تسجيل القضايا</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    with st.form("form"):
        st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
        if st.form_submit_button("💾 حفظ"): st.success("تم الحفظ")

elif st.session_state.page == "تقارير":
    st.markdown("<h2 style='color:#FFD700'>📑 التقارير</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"

elif st.session_state.page == "ارشيف":
    st.markdown("<h2 style='color:#FFD700'>🗃️ الأرشيف</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"

elif st.session_state.page == "مكتبة":
    st.markdown("<h2 style='color:#FFD700'>📚 المكتبة القانونية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"

elif st.session_state.page == "احصائيات":
    st.markdown("<h2 style='color:#FFD700'>📈 الإحصائيات التفصيلية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"

elif st.session_state.page == "بحث":
    st.markdown("<h2 style='color:#FFD700'>🔍 البحث</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"

# الفوتر
st.markdown("<div class='footer'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>", unsafe_allow_html=True)
