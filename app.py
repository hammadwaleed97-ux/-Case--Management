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

    st.markdown("""
    <h2 style='text-align:center;color:#FFD700'>
    📚 تسجيل القضايا
    </h2>
    """, unsafe_allow_html=True)

    st.divider()

    col1, col2 = st.columns(2)

    # ======================================================
    # العمود الأول
    # ======================================================

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

        claimant = st.text_input(
            "اسم رافع الدعوى"
        )

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

        case_number = st.text_input(
            "رقم الدعوى"
        )

        judicial_year = st.text_input(
            "السنة القضائية"
        )

        circuit = st.text_input(
            "الدائرة"
        )

    # ======================================================
    # العمود الثانى
    # ======================================================

    with col2:

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

        if case_type == "استئناف":

            appeal_office = st.text_input(
                "مأمورية الاستئناف"
            )

        else:

            appeal_office = ""

        subject = st.text_area(
            "موضوع الدعوى"
        )

        session_date = st.date_input(
            "تاريخ أول جلسة"
        )

        procedure = st.text_area(
            "الإجراء المطلوب"
        )

    st.divider()

    col3, col4 = st.columns(2)

    with col3:

        notifications_enabled = st.checkbox(
            "تفعيل التنبيهات",
            value=True
        )

        mobile = st.text_input(
            "رقم الهاتف"
        )

    with col4:

        notes = st.text_area(
            "ملاحظات"
        )

    st.info("حالة القضية عند التسجيل ستكون (متداولة) تلقائياً.")

    st.divider()
    # ==========================================================
# حفظ القضية
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

                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

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
                        now

                    )

                )

                case_id = cur.lastrowid

                cur.execute(
                    """
                    INSERT INTO sessions(

                        case_id,
                        session_date,
                        procedure,
                        created_at

                    )

                    VALUES(

                        ?,?,?,?

                    )
                    """,

                    (

                        case_id,
                        str(session_date),
                        procedure,
                        now

                    )

                )

                conn.commit()

                st.success("✅ تم تسجيل القضية وحفظ أول جلسة بنجاح.")

                st.balloons()

# ==========================================================
# نهاية قسم تسجيل القضايا
# ==========================================================
# ==========================================================
# الحصر العام للقضايا
# ==========================================================

elif page == "general":

    if st.button("🏠 العودة للصفحة الرئيسية", key="back_general"):
        st.session_state.page = "home"
        st.rerun()

    st.markdown("""
    <h2 style="text-align:center;color:#FFD700;">
    📑 الحصر العام للقضايا
    </h2>
    """, unsafe_allow_html=True)

    st.divider()

    search = st.text_input(
        "🔍 البحث فى القضايا",
        placeholder="رقم الدعوى أو الخصوم أو الموضوع..."
    )

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
        """,(
            f"%{search}%",
            f"%{search}%",
            f"%{search}%",
            f"%{search}%"
        ))

    cases = cur.fetchall()

    st.success(f"إجمالى القضايا : {len(cases)}")

    st.divider()

    if not cases:

        st.warning("لا توجد قضايا مسجلة.")

    else:

        for case in cases:

            case_id = case[0]

            with st.expander(
                f"📁 القضية رقم {case[6]} / {case[7]}",
                expanded=False
            ):

                st.subheader("بيانات القضية")

                col1, col2 = st.columns(2)

                with col1:

                    st.write("**نوع القضية:**", case[1])

                    st.write("**صفة رافع الدعوى:**", case[2])

                    st.write("**رافع الدعوى:**", case[3])

                    st.write("**صفة الخصم:**", case[4])

                    st.write("**الخصم:**", case[5])

                    st.write("**رقم الدعوى:**", case[6])

                    st.write("**السنة القضائية:**", case[7])

                    st.write("**الدائرة:**", case[11])

                with col2:

                    st.write("**المحكمة:**", case[8])

                    st.write("**اسم المحكمة:**", case[9])

                    st.write("**مأمورية الاستئناف:**", case[10])

                    st.write("**موضوع الدعوى:**", case[12])

                    st.write("**الحالة:**", case[13])

                    st.write("**رقم الهاتف:**", case[15])

                    st.write("**ملاحظات:**", case[16])

                st.divider()
                                # ============================================
                # آخر جلسة
                # ============================================

                cur.execute("""
                SELECT
                    session_date,
                    roll_number,
                    procedure,
                    adjournment_reason,
                    session_notes,
                    is_judgment,
                    judgment_date,
                    judgment_text,
                    judgment_result
                FROM sessions
                WHERE case_id=?
                ORDER BY id DESC
                LIMIT 1
                """,(case_id,))

                last_session = cur.fetchone()

                if last_session:

                    st.markdown("### 📅 آخر جلسة")

                    st.write("**تاريخ الجلسة:**", last_session[0])

                    st.write("**رقم الرول:**", last_session[1])

                    st.write("**الإجراء المطلوب:**", last_session[2])

                    if last_session[3]:
                        st.write("**سبب التأجيل:**", last_session[3])

                    if last_session[4]:
                        st.write("**ملاحظات:**", last_session[4])

                    if last_session[5] == 1:

                        st.success("⚖️ تم الفصل فى الدعوى")

                        st.write("**تاريخ الحكم:**", last_session[6])

                        st.write("**منطوق الحكم:**", last_session[7])

                        st.write("**النتيجة:**", last_session[8])

                else:

                    st.info("لا توجد جلسات حتى الآن.")

                st.divider()

                # ============================================
                # أزرار الإدارة
                # ============================================

                col_a, col_b, col_c, col_d = st.columns(4)

                with col_a:

                    add_session = st.button(
                        "➕ إضافة جلسة",
                        key=f"session_{case_id}",
                        use_container_width=True
                    )

                with col_b:

                    edit_case = st.button(
                        "✏️ تعديل",
                        key=f"edit_{case_id}",
                        use_container_width=True
                    )

                with col_c:

                    documents = st.button(
                        "📂 مستندات القضية",
                        key=f"docs_{case_id}",
                        use_container_width=True
                    )

                with col_d:

                    delete_case = st.button(
                        "🗑️ حذف",
                        key=f"delete_{case_id}",
                        use_container_width=True
                    )

                st.divider()
                                # ============================================
                # إضافة جلسة جديدة
                # ============================================

                if add_session:

                    st.markdown("## ➕ إضافة جلسة جديدة")

                    session_date = st.date_input(
                        "تاريخ الجلسة",
                        key=f"date_{case_id}"
                    )

                    roll_number = st.text_input(
                        "رقم الرول",
                        key=f"roll_{case_id}"
                    )

                    procedure = st.text_area(
                        "الإجراء المطلوب",
                        key=f"procedure_{case_id}"
                    )

                    adjournment_reason = st.text_area(
                        "سبب التأجيل",
                        key=f"reason_{case_id}"
                    )

                    session_notes = st.text_area(
                        "ملاحظات الجلسة",
                        key=f"notes_{case_id}"
                    )

                    st.divider()

                    is_judgment = st.checkbox(
                        "⚖️ هذه جلسة حكم",
                        key=f"judgment_{case_id}"
                    )

                    judgment_date = ""
                    judgment_text = ""
                    judgment_result = ""

                    if is_judgment:

                        judgment_date = st.date_input(
                            "تاريخ الحكم",
                            key=f"judgment_date_{case_id}"
                        )

                        judgment_text = st.text_area(
                            "منطوق الحكم",
                            key=f"judgment_text_{case_id}"
                        )

                        judgment_result = st.selectbox(
                            "النتيجة",
                            [
                                "لصالح الهيئة",
                                "ضد الهيئة"
                            ],
                            key=f"judgment_result_{case_id}"
                        )

                    if st.button(
                        "💾 حفظ الجلسة",
                        key=f"save_session_{case_id}",
                        use_container_width=True
                    ):

                        from datetime import datetime

                        now = datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )

                        cur.execute("""

                        INSERT INTO sessions(

                        case_id,

                        session_date,

                        roll_number,

                        procedure,

                        adjournment_reason,

                        session_notes,

                        is_judgment,

                        judgment_date,

                        judgment_text,

                        judgment_result,

                        created_at

                        )

                        VALUES(

                        ?,?,?,?,?,?,?,?,?,?,?

                        )

                        """,(

                        case_id,

                        str(session_date),

                        roll_number,

                        procedure,

                        adjournment_reason,

                        session_notes,

                        int(is_judgment),

                        str(judgment_date) if is_judgment else "",

                        judgment_text,

                        judgment_result,

                        now

                        ))

                        if is_judgment:

                            cur.execute("""

                            UPDATE cases

                            SET status=?

                            WHERE id=?

                            """,(

                            f"منتهية - {judgment_result}",

                            case_id

                            ))

                        conn.commit()

                        st.success("✅ تم حفظ الجلسة بنجاح")

                        st.rerun()
                                        # ============================================
                # جميع الجلسات السابقة
                # ============================================

                st.divider()

                st.markdown("## 📅 سجل الجلسات")

                cur.execute("""

                SELECT *

                FROM sessions

                WHERE case_id=?

                ORDER BY session_date DESC,id DESC

                """,(case_id,))

                sessions = cur.fetchall()

                if not sessions:

                    st.info("لا توجد جلسات.")

                else:

                    for s in sessions:

                        with st.expander(
                            f"📅 جلسة {s[2]}"
                        ):

                            new_date = st.date_input(
                                "تاريخ الجلسة",
                                value=s[2],
                                key=f"date_edit_{s[0]}"
                            )

                            new_roll = st.text_input(
                                "رقم الرول",
                                value=s[3],
                                key=f"roll_edit_{s[0]}"
                            )

                            new_procedure = st.text_area(
                                "الإجراء المطلوب",
                                value=s[4],
                                key=f"proc_edit_{s[0]}"
                            )

                            new_reason = st.text_area(
                                "سبب التأجيل",
                                value=s[5],
                                key=f"reason_edit_{s[0]}"
                            )

                            new_notes = st.text_area(
                                "ملاحظات",
                                value=s[6],
                                key=f"notes_edit_{s[0]}"
                            )

                            judgment = st.checkbox(
                                "جلسة حكم",
                                value=bool(s[7]),
                                key=f"judgment_edit_{s[0]}"
                            )

                            judgment_date = ""
                            judgment_text = ""
                            judgment_result = ""

                            if judgment:

                                judgment_date = st.date_input(
                                    "تاريخ الحكم",
                                    value=s[8] if s[8] else None,
                                    key=f"jdate_{s[0]}"
                                )

                                judgment_text = st.text_area(
                                    "منطوق الحكم",
                                    value=s[9],
                                    key=f"jtext_{s[0]}"
                                )

                                judgment_result = st.selectbox(
                                    "النتيجة",
                                    [
                                        "لصالح الهيئة",
                                        "ضد الهيئة"
                                    ],
                                    index=0 if s[10]!="ضد الهيئة" else 1,
                                    key=f"jresult_{s[0]}"
                                )

                            c1,c2=st.columns(2)

                            with c1:

                                if st.button(
                                    "💾 حفظ التعديل",
                                    key=f"save_edit_{s[0]}",
                                    use_container_width=True
                                ):

                                    cur.execute("""

                                    UPDATE sessions

                                    SET

                                    session_date=?,

                                    roll_number=?,

                                    procedure=?,

                                    adjournment_reason=?,

                                    session_notes=?,

                                    is_judgment=?,

                                    judgment_date=?,

                                    judgment_text=?,

                                    judgment_result=?

                                    WHERE id=?

                                    """,(

                                    str(new_date),

                                    new_roll,

                                    new_procedure,

                                    new_reason,

                                    new_notes,

                                    int(judgment),

                                    str(judgment_date) if judgment else "",

                                    judgment_text,

                                    judgment_result,

                                    s[0]

                                    ))

                                    if judgment:

                                        cur.execute("""

                                        UPDATE cases

                                        SET status=?

                                        WHERE id=?

                                        """,(

                                        f"منتهية - {judgment_result}",

                                        case_id

                                        ))

                                    conn.commit()

                                    st.success("تم تعديل الجلسة")

                                    st.rerun()

                            with c2:

                                if s[0] != sessions[-1][0]:

                                    if st.button(
                                        "🗑️ حذف الجلسة",
                                        key=f"delete_session_{s[0]}",
                                        use_container_width=True
                                    ):

                                        cur.execute(
                                            "DELETE FROM sessions WHERE id=?",
                                            (s[0],)
                                        )

                                        conn.commit()

                                        st.success("تم حذف الجلسة")

                                        st.rerun()
