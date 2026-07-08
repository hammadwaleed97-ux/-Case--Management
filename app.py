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

# ==========================================================
# جدول القضايا
# ==========================================================

cur.execute("""
CREATE TABLE IF NOT EXISTS cases(

id INTEGER PRIMARY KEY AUTOINCREMENT,

case_type TEXT,

claimant_type TEXT,
claimant TEXT,

defendant_type TEXT,
defendant TEXT,

case_number TEXT,
judicial_year TEXT,

court TEXT,
court_name TEXT,

appeal_office TEXT,

circuit TEXT,

subject TEXT,

status TEXT DEFAULT 'متداولة',

notifications_enabled INTEGER DEFAULT 1,

mobile TEXT,

notes TEXT,

created_at TEXT

)
""")

# ==========================================================
# جدول الجلسات
# ==========================================================

cur.execute("""
CREATE TABLE IF NOT EXISTS sessions(

id INTEGER PRIMARY KEY AUTOINCREMENT,

case_id INTEGER,

session_date TEXT,

roll_number TEXT,

procedure TEXT,

adjournment_reason TEXT,

session_notes TEXT,

is_judgment INTEGER DEFAULT 0,

judgment_date TEXT,

judgment_text TEXT,

judgment_result TEXT,

created_at TEXT,

FOREIGN KEY(case_id) REFERENCES cases(id)

)
""")

# ==========================================================
# جدول القضايا المحذوفة
# ==========================================================

cur.execute("""
CREATE TABLE IF NOT EXISTS deleted_cases(

id INTEGER PRIMARY KEY AUTOINCREMENT,

case_id INTEGER,

reason TEXT,

deleted_at TEXT

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
# تسجيل القضايا
# ==========================================================

elif page == "register":

    if st.button("🏠 العودة للصفحة الرئيسية", key="back_register"):
        st.session_state.page = "home"
        st.rerun()

    st.markdown(
        """
        <h2 style="text-align:center;color:#FFD700;">
        📚 تسجيل القضايا
        </h2>
        """,
        unsafe_allow_html=True
    )

    st.divider()

    col1, col2 = st.columns(2)

    with col1:

        case_type = st.selectbox(
            "نوع الدعوى",
            [
                "دعوى",
                "استئناف",
                "نقض"
            ]
        )

        claimant_type = st.selectbox(
            "صفة رافع الدعوى",
            [
                "المدعى",
                "المستأنف",
                "الطاعن"
            ]
        )

        claimant = st.text_input("اسم رافع الدعوى")

        case_number = st.text_input("رقم الدعوى")

        judicial_year = st.text_input("السنة القضائية")

        circuit = st.text_input("الدائرة")

    with col2:

        defendant_type = st.selectbox(
            "صفة الخصم",
            [
                "المدعى عليه",
                "المستأنف ضده",
                "المطعون ضده"
            ]
        )

        defendant = st.text_input("اسم الخصم")

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

        court_name = st.text_input("اسم المحكمة")

        if case_type == "استئناف":

            appeal_office = st.text_input("مأمورية الاستئناف")

        else:

            appeal_office = ""

        subject = st.text_area("موضوع الدعوى")

    st.divider()

    col3, col4 = st.columns(2)

    with col3:

        notifications_enabled = st.checkbox(
            "تفعيل التنبيهات",
            value=True
        )

        mobile = st.text_input("رقم الهاتف")

    with col4:

        notes = st.text_area("ملاحظات")

    st.divider()
    # ==========================================================
# الجزء الثانى : حفظ القضية
# ==========================================================

    if st.button("💾 حفظ القضية", use_container_width=True):

        if (
            claimant.strip() == ""
            or defendant.strip() == ""
            or case_number.strip() == ""
            or judicial_year.strip() == ""
        ):

            st.error("يرجى استكمال البيانات الأساسية.")

        else:

            cur.execute(
                """
                SELECT id
                FROM cases
                WHERE case_number=?
                AND judicial_year=?
                """,
                (case_number, judicial_year)
            )

            if cur.fetchone():

                st.warning("هذه القضية مسجلة بالفعل.")

            else:

                from datetime import datetime

                cur.execute(
                    """
                    INSERT INTO cases(

                        case_type,
                        claimant_type,
                        claimant,
                        defendant_type,
                        defendant,
                        case_number,
                        judicial_year,
                        court,
                        court_name,
                        appeal_office,
                        circuit,
                        subject,
                        status,
                        notifications_enabled,
                        mobile,
                        notes,
                        created_at

                    )

                    VALUES(

                        ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?

                    )
                    """,

                    (

                        case_type,
                        claimant_type,
                        claimant,
                        defendant_type,
                        defendant,
                        case_number,
                        judicial_year,
                        court,
                        court_name,
                        appeal_office,
                        circuit,
                        subject,
                        "متداولة",
                        int(notifications_enabled),
                        mobile,
                        notes,
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    )

                )

                conn.commit()

                st.success("✅ تم حفظ القضية بنجاح")

                st.balloons()

# ==========================================================
# نهاية قسم تسجيل القضايا
# ==========================================================
    # ==========================================================
# الحصر العام للقضايا
# ==========================================================

elif page == "general":

    if st.button("🏠 العودة للصفحة الرئيسية", use_container_width=True):
        st.session_state.page = "home"
        st.rerun()

    st.markdown("""
    <h2 style='text-align:center;color:#FFD700'>
    📑 الحصر العام للقضايا
    </h2>
    """, unsafe_allow_html=True)

    st.divider()

    search = st.text_input("🔍 البحث داخل القضايا")

    if search.strip() == "":

        cur.execute("""
            SELECT *
            FROM cases
            ORDER BY id DESC
        """)

    else:

        cur.execute("""
            SELECT *
            FROM cases
            WHERE
            claimant LIKE ?
            OR defendant LIKE ?
            OR case_number LIKE ?
            OR subject LIKE ?
            ORDER BY id DESC
        """, (
            f"%{search}%",
            f"%{search}%",
            f"%{search}%",
            f"%{search}%"
        ))

    rows = cur.fetchall()

    st.success(f"إجمالي القضايا : {len(rows)}")

    st.divider()

    if len(rows) == 0:

        st.warning("لا توجد قضايا مسجلة.")

    else:

        for row in rows:

            with st.expander(f"📁 القضية رقم {row[6]} / {row[7]}"):

                st.write("**نوع الطعن:**", row[1])
                st.write("**رافع الدعوى:**", row[3])
                st.write("**الخصم:**", row[5])
                st.write("**المحكمة:**", row[10])
                st.write("**اسم المحكمة:**", row[11])
                st.write("**الدائرة:**", row[8])
                st.write("**نوع الدعوى:**", row[9])
                st.write("**الموضوع:**", row[13])
                st.write("**رقم الرول:**", row[14])
                st.write("**تاريخ الجلسة:**", row[15])
                st.write("**سبب التأجيل:**", row[16])
                st.write("**الحالة:**", row[18])
                st.write("**رقم الهاتف:**", row[20])
                st.write("**ملاحظات:**", row[17])

                st.divider()

                c1, c2, c3 = st.columns(3)

                with c1:
                    st.button(
                        "✏️ تعديل",
                        key=f"edit_{row[0]}",
                        use_container_width=True
                    )

                with c2:
                    st.button(
                        "📂 متابعة",
                        key=f"follow_{row[0]}",
                        use_container_width=True
                    )

                with c3:
                    st.button(
                        "🗑️ حذف",
                        key=f"delete_{row[0]}",
                        use_container_width=True
                )
