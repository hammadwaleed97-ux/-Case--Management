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

    if st.button(
        "⬅️ العودة للرئيسية",
        key="back_register",
        use_container_width=True
    ):
        st.session_state.page = "home"
        st.rerun()

    st.markdown("""

    <style>

    .register-box{

        background:linear-gradient(180deg,#071B38,#06162F);

        border:2px solid #D4AF37;

        border-radius:25px;

        padding:25px;

        margin-top:10px;

        margin-bottom:25px;

        box-shadow:0 0 20px rgba(212,175,55,.35);

    }

    .register-title{

        text-align:center;

        color:#FFD700;

        font-size:42px;

        font-weight:900;

        margin-bottom:5px;

        text-shadow:0px 0px 12px gold;

    }

    .register-sub{

        text-align:center;

        color:white;

        font-size:18px;

        margin-bottom:25px;

    }

    .section-title{

        background:#0B2E63;

        color:#FFD700;

        padding:12px;

        border-radius:12px;

        font-size:24px;

        font-weight:bold;

        border:1px solid #D4AF37;

        margin-top:15px;

        margin-bottom:15px;

    }

    </style>

    <div class="register-box">

        <div class="register-title">

        ⚖️ تسجيل قضية جديدة

        </div>

        <div class="register-sub">

        أدخل جميع بيانات القضية بدقة

        </div>

    """, unsafe_allow_html=True)
    # ==========================================================
# الجزء الثانى (2/15)
# كارت بيانات القضية
# يبدأ بعد عنوان (تسجيل قضية جديدة)
# وينتهى قبل (بيانات أول جلسة)
# ==========================================================

st.markdown("""
<div class="section-title">
📋 أولاً : بيانات القضية
</div>
""", unsafe_allow_html=True)

left,right = st.columns(2)

# ======================================================
# العمود الأيمن
# ======================================================

with right:

    case_number = st.text_input(
        "رقم الدعوى *",
        placeholder="أدخل رقم الدعوى"
    )

    judicial_year = st.text_input(
        "السنة القضائية *",
        placeholder="مثال : 2026"
    )

    circuit = st.text_input(
        "الدائرة *"
    )

    case_type = st.selectbox(
        "نوع الدعوى *",
        [
            "دعوى",
            "استئناف",
            "نقض"
        ]
    )

    court = st.selectbox(
        "المحكمة *",
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
        "اسم المحكمة *"
    )

    if case_type == "استئناف":

        appeal_office = st.text_input(
            "مأمورية الاستئناف"
        )

    else:

        appeal_office = ""

# ======================================================
# العمود الأيسر
# ======================================================

with left:

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

    subject = st.text_area(
        "موضوع الدعوى",
        height=120
    )

    mobile = st.text_input(
        "رقم الهاتف"
    )

    notifications_enabled = st.checkbox(
        "تفعيل التنبيهات",
        value=True
    )

    notes = st.text_area(
        "ملاحظات"
    )

# ==========================================================
# نهاية الجزء الثانى
# الجزء الثالث يبدأ من:
# 📅 ثانياً : بيانات أول جلسة
# ==========================================================
# ==========================================================
# الحصر العام للقضايا
# ==========================================================

elif page == "general":

    if st.button("🏠 العودة للصفحة الرئيسية", key="back_general"):
        st.session_state.page = "home"
        st.rerun()

    st.markdown("""
    <h2 style='text-align:center;color:#FFD700'>
    📑 الحصر العام للقضايا
    </h2>
    """, unsafe_allow_html=True)

    st.divider()

    search = st.text_input(
        "🔍 البحث داخل القضايا",
        placeholder="رقم الدعوى - الخصوم - الموضوع"
    )

    if search.strip():

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

    else:

        cur.execute("""
        SELECT *
        FROM cases
        ORDER BY id DESC
        """)

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

                st.markdown("## 📋 بيانات القضية")

                left, right = st.columns(2)

                with left:

                    st.write("**نوع الدعوى :**", case[1])
                    st.write("**صفة رافع الدعوى :**", case[2])
                    st.write("**رافع الدعوى :**", case[3])
                    st.write("**صفة الخصم :**", case[4])
                    st.write("**الخصم :**", case[5])
                    st.write("**رقم الدعوى :**", case[6])
                    st.write("**السنة القضائية :**", case[7])
                    st.write("**الدائرة :**", case[11])

                with right:

                    st.write("**المحكمة :**", case[8])
                    st.write("**اسم المحكمة :**", case[9])

                    if case[10]:
                        st.write("**مأمورية الاستئناف :**", case[10])

                    st.write("**موضوع الدعوى :**", case[12])
                    st.write("**الحالة :**", case[13])
                    st.write("**رقم الهاتف :**", case[15])
                    st.write("**ملاحظات :**", case[16])

                st.divider()
                                # ==================================================
                # آخر جلسة
                # ==================================================

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

                    st.markdown("## 📅 آخر جلسة")

                    st.write("**تاريخ الجلسة :**", last_session[0])

                    st.write("**رقم الرول :**", last_session[1])

                    st.write("**الإجراء المطلوب :**", last_session[2])

                    if last_session[3]:
                        st.write("**سبب التأجيل :**", last_session[3])

                    if last_session[4]:
                        st.write("**ملاحظات الجلسة :**", last_session[4])

                    if last_session[5]:

                        st.success("⚖️ صدر حكم فى الدعوى")

                        st.write("**تاريخ الحكم :**", last_session[6])

                        st.write("**منطوق الحكم :**")

                        st.info(last_session[7])

                        st.write("**النتيجة :**", last_session[8])

                else:

                    st.info("لا توجد جلسات مسجلة لهذه القضية.")

                st.divider()

                # ==================================================
                # متابعة القضية
                # ==================================================

                st.markdown("## 📂 متابعة القضية")

                tab1, tab2, tab3, tab4 = st.tabs([

                    "➕ إضافة جلسة",

                    "📅 الجلسات السابقة",

                    "📎 مستندات القضية",

                    "✏️ تعديل بيانات القضية"

                ])
                                # ==================================================
                # تبويب إضافة جلسة
                # ==================================================

                with tab1:

                    session_date = st.date_input(
                        "تاريخ الجلسة",
                        key=f"session_date_{case_id}"
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

                    judgment_date = None
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

                        judgment_result = st.radio(
                            "نتيجة الحكم",
                            ["لصالح الهيئة", "ضد الهيئة"],
                            horizontal=True,
                            key=f"judgment_result_{case_id}"
                        )

                    if st.button(
                        "💾 حفظ الجلسة",
                        key=f"save_session_{case_id}",
                        use_container_width=True
                    ):

                        from datetime import datetime

                        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

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

                # ==================================================
                # تبويب الجلسات السابقة
                # ==================================================

                with tab2:

                    cur.execute("""

                    SELECT *

                    FROM sessions

                    WHERE case_id=?

                    ORDER BY session_date DESC,id DESC

                    """,(case_id,))

                    sessions = cur.fetchall()

                    if not sessions:

                        st.info("لا توجد جلسات مسجلة.")

                    else:

                        for s in sessions:

                            with st.expander(f"📅 جلسة {s[2]}"):

                                st.write("**رقم الرول:**", s[3])

                                st.write("**الإجراء المطلوب:**", s[4])

                                if s[5]:
                                    st.write("**سبب التأجيل:**", s[5])

                                if s[6]:
                                    st.write("**ملاحظات:**", s[6])

                                if s[7]:

                                    st.success("جلسة حكم")

                                    st.write("**تاريخ الحكم:**", s[8])

                                    st.write("**منطوق الحكم:**")

                                    st.info(s[9])

                                    st.write("**النتيجة:**", s[10])

                                if st.button(
                                    "✏️ تعديل هذه الجلسة",
                                    key=f"edit_session_{s[0]}",
                                    use_container_width=True
                                ):

                                    st.session_state[f"edit_session_{case_id}"] = s[0]
                # ==================================================
                # تعديل جلسة سابقة
                # ==================================================

                if st.session_state.get(f"edit_session_{case_id}"):

                    session_id = st.session_state[f"edit_session_{case_id}"]

                    cur.execute("""
                    SELECT *
                    FROM sessions
                    WHERE id=?
                    """,(session_id,))

                    s = cur.fetchone()

                    if s:

                        st.divider()

                        st.markdown("## ✏️ تعديل الجلسة")

                        edit_session_date = st.date_input(
                            "تاريخ الجلسة",
                            value=s[2],
                            key=f"edit_date_{session_id}"
                        )

                        edit_roll = st.text_input(
                            "رقم الرول",
                            value=s[3],
                            key=f"edit_roll_{session_id}"
                        )

                        edit_procedure = st.text_area(
                            "الإجراء المطلوب",
                            value=s[4],
                            key=f"edit_proc_{session_id}"
                        )

                        edit_reason = st.text_area(
                            "سبب التأجيل",
                            value=s[5],
                            key=f"edit_reason_{session_id}"
                        )

                        edit_notes = st.text_area(
                            "ملاحظات الجلسة",
                            value=s[6],
                            key=f"edit_notes_{session_id}"
                        )

                        edit_is_judgment = st.checkbox(
                            "جلسة حكم",
                            value=bool(s[7]),
                            key=f"edit_judgment_{session_id}"
                        )

                        edit_judgment_date = ""
                        edit_judgment_text = ""
                        edit_judgment_result = ""

                        if edit_is_judgment:

                            edit_judgment_date = st.date_input(
                                "تاريخ الحكم",
                                value=s[8] if s[8] else None,
                                key=f"edit_jdate_{session_id}"
                            )

                            edit_judgment_text = st.text_area(
                                "منطوق الحكم",
                                value=s[9],
                                key=f"edit_jtext_{session_id}"
                            )

                            edit_judgment_result = st.radio(
                                "نتيجة الحكم",
                                ["لصالح الهيئة","ضد الهيئة"],
                                horizontal=True,
                                index=0 if s[10] != "ضد الهيئة" else 1,
                                key=f"edit_jresult_{session_id}"
                            )

                        col_save,col_delete = st.columns(2)

                        with col_save:

                            if st.button(
                                "💾 حفظ التعديل",
                                key=f"save_edit_{session_id}",
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

                                    str(edit_session_date),
                                    edit_roll,
                                    edit_procedure,
                                    edit_reason,
                                    edit_notes,
                                    int(edit_is_judgment),
                                    str(edit_judgment_date) if edit_is_judgment else "",
                                    edit_judgment_text,
                                    edit_judgment_result,
                                    session_id

                                ))

                                if edit_is_judgment:

                                    cur.execute("""
                                    UPDATE cases
                                    SET status=?
                                    WHERE id=?
                                    """,(
                                        f"منتهية - {edit_judgment_result}",
                                        case_id
                                    ))

                                conn.commit()

                                st.success("تم تعديل الجلسة")

                                st.session_state.pop(f"edit_session_{case_id}")

                                st.rerun()

                        with col_delete:

                            if st.button(
                                "🗑️ حذف الجلسة",
                                key=f"delete_session_{session_id}",
                                use_container_width=True
                            ):

                                cur.execute("""
                                SELECT COUNT(*)
                                FROM sessions
                                WHERE case_id=?
                                """,(case_id,))

                                total_sessions = cur.fetchone()[0]

                                if total_sessions <= 1:

                                    st.error("لا يمكن حذف آخر جلسة بالقضية.")

                                else:

                                    cur.execute(
                                        "DELETE FROM sessions WHERE id=?",
                                        (session_id,)
                                    )

                                    conn.commit()

                                    st.success("تم حذف الجلسة")

                                    st.session_state.pop(f"edit_session_{case_id}")

                                    st.rerun()
                                                    # ==================================================
                # تبويب مستندات القضية
                # ==================================================

                with tab3:

                    st.markdown("## 📎 مستندات القضية")

                    cur.execute("""
                    CREATE TABLE IF NOT EXISTS documents(

                        id INTEGER PRIMARY KEY AUTOINCREMENT,

                        case_id INTEGER,

                        document_name TEXT,

                        file_name TEXT,

                        file_path TEXT,

                        uploaded_at TEXT

                    )
                    """)

                    uploaded_file = st.file_uploader(

                        "إرفاق مستند",

                        type=[
                            "pdf",
                            "doc",
                            "docx",
                            "xls",
                            "xlsx",
                            "jpg",
                            "jpeg",
                            "png"
                        ],

                        key=f"upload_{case_id}"

                    )

                    document_name = st.text_input(

                        "اسم المستند",

                        key=f"doc_name_{case_id}"

                    )

                    if st.button(

                        "💾 حفظ المستند",

                        key=f"save_doc_{case_id}",

                        use_container_width=True

                    ):

                        if uploaded_file is None:

                            st.warning("اختر ملف أولاً.")

                        else:

                            os.makedirs("documents", exist_ok=True)

                            file_path = os.path.join(

                                "documents",

                                f"{case_id}_{uploaded_file.name}"

                            )

                            with open(file_path, "wb") as f:

                                f.write(uploaded_file.getbuffer())

                            from datetime import datetime

                            cur.execute("""

                            INSERT INTO documents(

                                case_id,

                                document_name,

                                file_name,

                                file_path,

                                uploaded_at

                            )

                            VALUES(

                                ?,?,?,?,?,?

                            )

                            """,(

                                case_id,

                                document_name,

                                uploaded_file.name,

                                file_path,

                                datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                            ))

                            conn.commit()

                            st.success("✅ تم حفظ المستند")

                            st.rerun()

                    st.divider()

                    cur.execute("""

                    SELECT

                    id,

                    document_name,

                    file_name,

                    uploaded_at

                    FROM documents

                    WHERE case_id=?

                    ORDER BY id DESC

                    """,(case_id,))

                    docs = cur.fetchall()

                    if docs:

                        for d in docs:

                            c1,c2 = st.columns([5,1])

                            with c1:

                                st.write(
                                    f"📄 {d[1]} - {d[2]}"
                                )

                            with c2:

                                if st.button(

                                    "🗑️",

                                    key=f"delete_doc_{d[0]}"

                                ):

                                    cur.execute(

                                        "DELETE FROM documents WHERE id=?",

                                        (d[0],)

                                    )

                                    conn.commit()

                                    st.rerun()

                    else:

                        st.info("لا توجد مستندات مرفقة.")
                                        # ==================================================
                # تبويب تعديل بيانات القضية
                # ==================================================

                with tab4:

                    st.markdown("## ✏️ تعديل بيانات القضية")

                    edit_case_type = st.selectbox(
                        "نوع الدعوى",
                        ["دعوى","استئناف","نقض"],
                        index=["دعوى","استئناف","نقض"].index(case[1]),
                        key=f"case_type_{case_id}"
                    )

                    edit_claimant_type = st.selectbox(
                        "صفة رافع الدعوى",
                        ["المدعى","المستأنف","الطاعن"],
                        index=["المدعى","المستأنف","الطاعن"].index(case[2]),
                        key=f"claimant_type_{case_id}"
                    )

                    edit_claimant = st.text_input(
                        "اسم رافع الدعوى",
                        value=case[3],
                        key=f"claimant_{case_id}"
                    )

                    edit_defendant_type = st.selectbox(
                        "صفة الخصم",
                        ["المدعى عليه","المستأنف ضده","المطعون ضده"],
                        index=["المدعى عليه","المستأنف ضده","المطعون ضده"].index(case[4]),
                        key=f"defendant_type_{case_id}"
                    )

                    edit_defendant = st.text_input(
                        "اسم الخصم",
                        value=case[5],
                        key=f"defendant_{case_id}"
                    )

                    edit_court = st.selectbox(
                        "المحكمة",
                        [
                            "الابتدائية",
                            "الاستئناف",
                            "النقض",
                            "الإدارية",
                            "القضاء الإداري",
                            "الإدارية العليا"
                        ],
                        index=[
                            "الابتدائية",
                            "الاستئناف",
                            "النقض",
                            "الإدارية",
                            "القضاء الإداري",
                            "الإدارية العليا"
                        ].index(case[8]),
                        key=f"court_{case_id}"
                    )

                    edit_court_name = st.text_input(
                        "اسم المحكمة",
                        value=case[9],
                        key=f"court_name_{case_id}"
                    )

                    if edit_case_type == "استئناف":

                        edit_appeal = st.text_input(
                            "مأمورية الاستئناف",
                            value=case[10],
                            key=f"appeal_{case_id}"
                        )

                    else:

                        edit_appeal = ""

                    edit_subject = st.text_area(
                        "موضوع الدعوى",
                        value=case[12],
                        key=f"subject_{case_id}"
                    )

                    edit_mobile = st.text_input(
                        "رقم الهاتف",
                        value=case[15],
                        key=f"mobile_{case_id}"
                    )

                    edit_notes = st.text_area(
                        "ملاحظات",
                        value=case[16],
                        key=f"notes_{case_id}"
                    )

                    if st.button(
                        "💾 حفظ تعديل بيانات القضية",
                        key=f"save_case_{case_id}",
                        use_container_width=True
                    ):

                        cur.execute("""

                        UPDATE cases

                        SET

                        case_type=?,
                        claimant_type=?,
                        claimant=?,
                        defendant_type=?,
                        defendant=?,
                        court=?,
                        court_name=?,
                        appeal_office=?,
                        subject=?,
                        mobile=?,
                        notes=?

                        WHERE id=?

                        """,(

                            edit_case_type,
                            edit_claimant_type,
                            edit_claimant,
                            edit_defendant_type,
                            edit_defendant,
                            edit_court,
                            edit_court_name,
                            edit_appeal,
                            edit_subject,
                            edit_mobile,
                            edit_notes,
                            case_id

                        ))

                        conn.commit()

                        st.success("✅ تم حفظ التعديلات")

                        st.rerun()

                    st.divider()

                    if st.button(
                        "🗑️ حذف القضية نهائياً",
                        key=f"delete_case_{case_id}",
                        use_container_width=True
                    ):

                        cur.execute(
                            "DELETE FROM sessions WHERE case_id=?",
                            (case_id,)
                        )

                        cur.execute(
                            "DELETE FROM documents WHERE case_id=?",
                            (case_id,)
                        )

                        cur.execute(
                            "DELETE FROM cases WHERE id=?",
                            (case_id,)
                        )

                        conn.commit()

                        st.success("✅ تم حذف القضية نهائياً")

                        st.rerun()
