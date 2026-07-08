# ==========================================================
# إدارة القضايا
# Professional Judicial Edition
# الجزء الأول (1/15)
# ==========================================================

import streamlit as st
from PIL import Image
import sqlite3
import os

# ==========================================================
# إعداد الصفحة
# ==========================================================

st.set_page_config(
    page_title="إدارة القضايا",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================================
# إخفاء عناصر Streamlit
# ==========================================================

st.markdown("""
<style>

#MainMenu{
visibility:hidden;
}

header{
visibility:hidden;
}

footer{
visibility:hidden;
}

.block-container{
padding-top:0rem;
padding-left:0rem;
padding-right:0rem;
padding-bottom:2rem;
max-width:100%;
}

.stApp{
background:#071C3A;
}

/* ================= الأزرار ================= */

div.stButton > button{

width:100%;

height:70px;

background:linear-gradient(90deg,#0B2E63,#114C9A);

color:white;

font-size:24px;

font-weight:bold;

border-radius:15px;

border:2px solid #FFD700;

transition:.3s;

box-shadow:0px 0px 12px rgba(255,215,0,.4);

}

div.stButton > button:hover{

background:linear-gradient(90deg,#114C9A,#1565C0);

border:2px solid white;

transform:scale(1.02);

}

.title{

text-align:center;

font-size:42px;

font-weight:900;

color:#FFD700;

margin-top:20px;

margin-bottom:20px;

text-shadow:0px 0px 15px gold;

}

</style>

""", unsafe_allow_html=True)

# ==========================================================
# قاعدة البيانات
# ==========================================================

conn = sqlite3.connect("cases.db", check_same_thread=False)

# ==========================================================
# صورة الواجهة
# ==========================================================

if os.path.exists("home.png"):

    image = Image.open("home.png")

    st.image(image, use_container_width=True)

else:

    st.error("ملف home.png غير موجود")

# ==========================================================
# عنوان الصفحة
# ==========================================================

st.markdown("""
<div class="title">
⚖️ إدارة القضايا ⚖️
</div>
""", unsafe_allow_html=True)
# ==========================================================
# القائمة الرئيسية
# ==========================================================

if "page" not in st.session_state:
    st.session_state.page = "home"

col1, col2, col3 = st.columns([1,4,1])

with col2:

    if st.button("📚 تسجيل القضايا", use_container_width=True):
        st.session_state.page = "register"

    st.write("")

    if st.button("📑 الحصر العام للقضايا", use_container_width=True):
        st.session_state.page = "general"

    st.write("")

    if st.button("🔍 البحث عن دعوى", use_container_width=True):
        st.session_state.page = "search"

    st.write("")

    if st.button("📊 التقارير", use_container_width=True):
        st.session_state.page = "reports"

    st.write("")

    if st.button("🔔 التنبيهات", use_container_width=True):
        st.session_state.page = "notifications"

    st.write("")

    if st.button("🗂️ أرشيف القضايا", use_container_width=True):
        st.session_state.page = "archive"

    st.write("")

    if st.button("⚖️ المكتبة القانونية", use_container_width=True):
        st.session_state.page = "library"
        # ==========================================================
# الصفحة الرئيسية
# ==========================================================

if st.session_state.page == "home":

    st.markdown("<br>", unsafe_allow_html=True)

    c1, c2, c3 = st.columns([1,5,1])

    with c2:

        st.markdown("""
        <div style="
        background:linear-gradient(90deg,#062456,#0B3E91,#062456);
        border:2px solid #FFD700;
        border-radius:20px;
        padding:18px;
        text-align:center;
        color:white;
        font-size:26px;
        font-weight:bold;
        box-shadow:0px 0px 20px rgba(255,215,0,.45);
        ">
        مرحباً بك فى نظام إدارة القضايا
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        m1,m2,m3 = st.columns(3)

        with m1:
            st.metric("إجمالى القضايا", "0")

        with m2:
            st.metric("جلسات اليوم", "0")

        with m3:
            st.metric("التنبيهات", "0")

        st.markdown("<br>", unsafe_allow_html=True)

        st.info("اختر أحد الأقسام من الأعلى للبدء.")

# ==========================================================
# الأقسام
# ==========================================================

elif st.session_state.page == "register":

    st.title("📚 تسجيل القضايا")

elif st.session_state.page == "general":

    st.title("📑 الحصر العام للقضايا")

elif st.session_state.page == "search":

    st.title("🔍 البحث عن دعوى")

elif st.session_state.page == "reports":

    st.title("📊 التقارير")

elif st.session_state.page == "notifications":

    st.title("🔔 التنبيهات")

elif st.session_state.page == "archive":

    st.title("🗂️ أرشيف القضايا")

elif st.session_state.page == "library":

    st.title("⚖️ المكتبة القانونية")
