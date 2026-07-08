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
/* ================= إظهار أسماء الحقول ================= */

label{
    color:white !important;
    font-size:18px !important;
    font-weight:bold !important;
}

.stTextInput label,
.stTextArea label,
.stSelectbox label,
.stDateInput label,
.stCheckbox label{
    color:white !important;
    font-size:18px !important;
    font-weight:bold !important;
}

</style>

""", unsafe_allow_html=True)
# ==========================================================
# إنشاء قاعدة البيانات
# ==========================================================

conn = sqlite3.connect("cases.db", check_same_thread=False)
cur = conn.cursor()

cur.execute("""
CREATE TABLE IF NOT EXISTS cases(

id INTEGER PRIMARY KEY AUTOINCREMENT,

litigation_type TEXT,

claimant_type TEXT,

claimant TEXT,

defendant_type TEXT,

defendant TEXT,

case_number TEXT,

judicial_year TEXT,

circuit TEXT,

case_type TEXT,

court TEXT,

court_name TEXT,

appeal_office TEXT,

subject TEXT,

roll_number TEXT,

session_date TEXT,

adjournment_reason TEXT,

notes TEXT,

judgment_result TEXT,

notifications_enabled INTEGER,

mobile TEXT,

status TEXT DEFAULT 'متداولة',

created_at TEXT

)
""")

conn.commit()
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
# الصفحات
# ==========================================================

page = st.session_state.page
# ==========================================================
# الصفحة الرئيسية
# ==========================================================

if page == "home":

    st.markdown("<br>", unsafe_allow_html=True)

    c1, c2, c3 = st.columns([1,5,1])

    with c2:
        st.info("اختر أحد الأقسام من القائمة للبدء.")
        # ==========================================================
# بداية قسم تسجيل القضايا
# ==========================================================

elif page == "register":

    if st.button("🏠 العودة للصفحة الرئيسية", key="back_register"):
        st.session_state.page = "home"
        st.rerun()

    st.markdown("""
    <h2 style='text-align:center;color:#FFD700'>
    📚 تسجيل القضايا
    </h2>
    """, unsafe_allow_html=True)

    st.divider()

    col1, col2 = st.columns(2)

    with col1:

        litigation_type = st.selectbox(
            "نوع الطعن",
            ["دعوى", "استئناف", "نقض"]
        )

        claimant_type = st.selectbox(
            "صفة رافع الدعوى",
            ["المدعى", "المستأنف", "الطاعن"]
        )

        claimant = st.text_input(
            "اسم رافع الدعوى"
        )

        case_number = st.text_input(
            "رقم الدعوى"
        )

        judicial_year = st.text_input(
            "السنة القضائية"
        )

        circuit = st.text_input(
            "الدائرة"
        )

        case_type = st.text_input(
            "نوع الدعوى"
        )

    with col2:

        defendant_type = st.selectbox(
            "صفة الخصم",
            [
                "المدعى عليه",
                "المستأنف ضده",
                "المطعون ضده"
            ]
        )

        defendant = st.text_input(
            "اسم الخصم"
        )

        court = st.selectbox(
            "المحكمة",
            [
                "الابتدائية",
                "الاستئناف",
                "النقض",
                "الإدارية",
                "القضاء الإداري",
                "الإدارية العليا"
            ]
        )

        court_name = st.text_input(
            "اسم المحكمة"
        )

        appeal_office = st.text_input(
            "مكتب الاستئناف"
        )

        subject = st.text_area(
            "موضوع الدعوى"
        )

    st.divider()
