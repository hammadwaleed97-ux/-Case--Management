import streamlit as st
import sqlite3
from datetime import datetime
import os

# =====================================
# مكتبات تصدير التقارير
# ====================================

from io import BytesIO

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Mm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENTATION
from docx.enum.table import WD_TABLE_ALIGNMENT

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# =====================================
# إنشاء تقرير Word
# =====================================

def create_word_report(
    report_title,
    office,
    lawyer,
    from_date,
    to_date,
    headers,
    rows
):

    file = BytesIO()

    doc = Document()

    # =================================
    # إعداد الصفحة
    # =================================

    section = doc.sections[0]

    section.orientation = WD_ORIENTATION.LANDSCAPE

    section.page_width, section.page_height = (
        section.page_height,
        section.page_width
    )

    section.top_margin = Cm(1)
    section.bottom_margin = Cm(1)
    section.left_margin = Cm(1)
    section.right_margin = Cm(1)

    style = doc.styles["Normal"]

    style.font.name = "Arial"

    style._element.rPr.rFonts.set(
        qn("w:eastAsia"),
        "Arial"
    )

    style.font.size = Pt(9)

    COLOR_BLUE = RGBColor(17, 52, 120)
    COLOR_GOLD = RGBColor(184, 134, 11)

    # =================================
    # رأس الصفحة
    # =================================

    head = doc.add_table(rows=1, cols=3)
    head.alignment = WD_TABLE_ALIGNMENT.CENTER
    head.autofit = False

    head.columns[0].width = Cm(3)
    head.columns[1].width = Cm(3)
    head.columns[2].width = Cm(16)

    # إزالة الحدود

    for cell in head.rows[0].cells:

        tcPr = cell._tc.get_or_add_tcPr()

        borders = OxmlElement("w:tcBorders")

        for edge in ("top", "left", "bottom", "right"):

            b = OxmlElement(f"w:{edge}")
            b.set(qn("w:val"), "nil")
            borders.append(b)

        tcPr.append(borders)

    # =================================
    # شعار الهيئة (أعلى يسار)
    # =================================

    p = head.cell(0,0).paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    try:
        run = p.add_run()
        run.add_picture("logo.png", width=Mm(22))
    except:
        pass

    # =================================
    # بيانات الهيئة
    # =================================

    p = head.cell(0,2).paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    r = p.add_run("الهيئة القومية للتأمين الاجتماعى\n")
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = COLOR_BLUE

    r = p.add_run("الإدارة المركزية للادارات القانونية\n")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = COLOR_BLUE

    r = p.add_run("الإدارة العامة للشئون القانونية\n")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = COLOR_BLUE

    r = p.add_run(f"ديوان عام منطقة : {office}")
    r.bold = True
    r.font.size = Pt(12)

    doc.add_paragraph()

    # =================================
    # عنوان التقرير
    # =================================

    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    r = p.add_run(report_title)
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = COLOR_BLUE

    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    r = p.add_run(
        f"خلال الفترة من {from_date.strftime('%d/%m/%Y')} حتى {to_date.strftime('%d/%m/%Y')}"
    )
    r.bold = True
    r.font.size = Pt(11)

    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)

    r = p.add_run(
        f"طرف الأستاذ / {lawyer} - المحامى"
    )
    r.bold = True
    r.font.size = Pt(11)

    doc.add_paragraph()

    # =================================
    # إنشاء الجدول
    # =================================

    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.allow_autofit = False

    rtl_headers = list(reversed(headers))

    # عرض الأعمدة (سم)

    widths = [
        0.20,   # م
        0.60,   # رقم القضية
        0.35,   # السنة
        0.50,   # الدائرة
        0.80,   # نوع القضية
        1.40,   # المحكمة
        0.80,   # المأمورية
        10.50,  # الخصوم
        5.30,   # موضوع الدعوى
        1.00,   # آخر جلسة
        3.80,   # الإجراء
        1.00    # الملاحظات
    ]

    widths = list(reversed(widths))

    # ==========================
    # رؤوس الجدول
    # ==========================

    head_cells = table.rows[0].cells

    for i, h in enumerate(rtl_headers):

        cell = head_cells[i]

        tcPr = cell._tc.get_or_add_tcPr()

        tcW = OxmlElement("w:tcW")
        tcW.set(qn("w:type"), "dxa")
        tcW.set(qn("w:w"), str(int(widths[i] * 567)))
        tcPr.append(tcW)

        # خلفية أصفر فاتح
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "FFF2CC")
        tcPr.append(shd)

        p = cell.paragraphs[0]
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        run = p.add_run(str(h))
        run.bold = True
        run.font.size = Pt(6)
        run.font.color.rgb = RGBColor(0, 0, 0)

    # ==========================
    # بيانات الجدول
    # ==========================

    for row_data in rows:

        row_data = list(reversed(row_data))

        cells = table.add_row().cells

        for i, value in enumerate(row_data):

            cell = cells[i]

            tcPr = cell._tc.get_or_add_tcPr()
            tcW = OxmlElement("w:tcW")
            tcW.set(qn("w:type"), "dxa")
            tcW.set(qn("w:w"), str(int(widths[i] * 567)))
            tcPr.append(tcW)

            p = cell.paragraphs[0]

            if rtl_headers[i] in (
                "الخصوم",
                "موضوع الدعوى",
                "الإجراء",
                "سبب آخر جلسة",
                "ملاحظات"
            ):
                p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            else:
                p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            run = p.add_run("" if value is None else str(value))
            run.font.size = Pt(5.5)

    doc.add_paragraph()

    # =================================
    # التوقيعات
    # =================================

    doc.add_paragraph()
    doc.add_paragraph()

    sign = doc.add_table(rows=3, cols=2)
    sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    sign.autofit = False

    sign.columns[0].width = Cm(11)
    sign.columns[1].width = Cm(11)

    # إزالة الحدود
    for row in sign.rows:
        for cell in row.cells:

            tcPr = cell._tc.get_or_add_tcPr()

            tcBorders = OxmlElement("w:tcBorders")

            for edge in ("top", "left", "bottom", "right"):

                e = OxmlElement(f"w:{edge}")
                e.set(qn("w:val"), "nil")
                tcBorders.append(e)

            tcPr.append(tcBorders)

    # --------------------------
    # السطر الأول
    # --------------------------

    # اليمين
    p = sign.cell(0,1).paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    r = p.add_run("عضو الإدارة")
    r.bold = True
    r.font.size = Pt(11)

    # اليسار
    p = sign.cell(0,0).paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    r = p.add_run("مدير الإدارة")
    r.bold = True
    r.font.size = Pt(11)

    # --------------------------
    # السطر الثانى
    # --------------------------

    p = sign.cell(1,1).paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    p.add_run(".......................")

    p = sign.cell(1,0).paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p.add_run(".......................")

    # --------------------------
    # السطر الثالث
    # --------------------------

    sign.cell(2,0).merge(sign.cell(2,1))

    p = sign.cell(2,0).paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    r = p.add_run("مدير عام الإدارات القانونية")
    r.bold = True
    r.font.size = Pt(11)

    # ==========================
    # التاريخ
    # ==========================

    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    r = p.add_run(
        f"تحريراً فى : {datetime.now().strftime('%d/%m/%Y')}"
    )

    r.bold = True
    r.font.size = Pt(10)

    # ==========================
    # ترقيم الصفحات (Page X of Y)
    # ==========================

    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    p.add_run("صفحة ")

    # PAGE
    r = p.add_run()

    fldBegin = OxmlElement("w:fldChar")
    fldBegin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "

    fldSeparate = OxmlElement("w:fldChar")
    fldSeparate.set(qn("w:fldCharType"), "separate")

    text = OxmlElement("w:t")
    text.text = "1"

    fldEnd = OxmlElement("w:fldChar")
    fldEnd.set(qn("w:fldCharType"), "end")

    r._r.append(fldBegin)
    r._r.append(instr)
    r._r.append(fldSeparate)
    r._r.append(text)
    r._r.append(fldEnd)

    p.add_run(" من ")

    # NUMPAGES
    r = p.add_run()

    fldBegin = OxmlElement("w:fldChar")
    fldBegin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " NUMPAGES "

    fldSeparate = OxmlElement("w:fldChar")
    fldSeparate.set(qn("w:fldCharType"), "separate")

    text = OxmlElement("w:t")
    text.text = "1"

    fldEnd = OxmlElement("w:fldChar")
    fldEnd.set(qn("w:fldCharType"), "end")

    r._r.append(fldBegin)
    r._r.append(instr)
    r._r.append(fldSeparate)
    r._r.append(text)
    r._r.append(fldEnd)

    # تحديث الحقول عند فتح Word
    settings = doc.settings._element
    update = OxmlElement("w:updateFields")
    update.set(qn("w:val"), "true")
    settings.append(update)

    # ==========================
    # حفظ الملف
    # ==========================

    doc.save(file)
    file.seek(0)
    return file

# =====================================
# إعداد الصفحة
# =====================================

st.set_page_config(
    page_title="إدارة القضايا",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =====================================
# إعداد قاعدة البيانات والمستندات المحلية الدائمة
# =====================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "local_data")
DOCUMENTS_DIR = os.path.join(DATA_DIR, "Documents")
os.makedirs(DOCUMENTS_DIR, exist_ok=True)
LOCAL_DB_PATH = os.path.join(DATA_DIR, "cases.db")

@st.cache_resource
def get_local_connection(db_path):
    connection = sqlite3.connect(db_path, check_same_thread=False)
    connection.row_factory = sqlite3.Row
    return connection

conn = get_local_connection(LOCAL_DB_PATH)
cur = conn.cursor()

# =====================================
# Session State
# =====================================

if "page" not in st.session_state:
    st.session_state.page = "home"

if "selected_case" not in st.session_state:
    st.session_state.selected_case = None

if "delete_case_id" not in st.session_state:
    st.session_state.delete_case_id = None

if "form_reset" not in st.session_state:
    st.session_state.form_reset = 0

if "report_reset" not in st.session_state:
    st.session_state.report_reset = 0

if "show_report_flag" not in st.session_state:
    st.session_state.show_report_flag = False

# =====================================
# قاعدة البيانات المحلية
# =====================================

def init_local_db():
    cur.execute("""
    CREATE TABLE IF NOT EXISTS cases(

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    case_type TEXT,
    court_type TEXT,
    court_name TEXT,
    mission TEXT,

    case_number TEXT,
    judicial_year TEXT,
    circuit TEXT,
    case_category TEXT,

    plaintiff TEXT,
    defendant TEXT,

    subject TEXT,

    notes TEXT,

    whatsapp_enabled INTEGER,
    whatsapp_number TEXT,

    created_at TEXT

    )
    """)

    cur.execute("""

    CREATE TABLE IF NOT EXISTS sessions(

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    case_id INTEGER,

    session_date TEXT,

    roll_number TEXT,

    procedure TEXT,

    judgment_result TEXT,

    created_at TEXT

    )

    """)

    try:
        cur.execute("""
            ALTER TABLE sessions
            ADD COLUMN judgment_result TEXT
        """)
    except sqlite3.OperationalError:
        pass

    cur.execute("""

    CREATE TABLE IF NOT EXISTS documents(

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    case_id INTEGER,

    document_type TEXT,

    document_description TEXT,

    file_name TEXT,

    file_path TEXT,

    uploaded_at TEXT

    )

    """)

    conn.commit()

init_local_db()

# =====================================
# التصميم العام (CSS)
# =====================================

st.markdown("""
<style>

#MainMenu{visibility:hidden;}
header{visibility:hidden;}
footer{visibility:hidden;}

html,body,[class*="css"]{
    direction:rtl;
}

/* الخلفية */

.stApp{

background:
linear-gradient(
180deg,
#5A3A22 0%,
#4A2F1C 35%,
#352114 70%,
#24160D 100%
);

}

/* ميزان الخلفية */

.watermark{

position:fixed;

top:50%;

left:50%;

transform:translate(-50%,-50%);

font-size:320px;

opacity:.05;

z-index:0;

pointer-events:none;

}

/* رأس البرنامج */

.header{

text-align:center;

position:relative;

z-index:999;

}

.logo{

font-size:90px;

margin-bottom:-15px;

}

.title1{

font-size:34px;

font-weight:bold;

color:#FFD700;

}

.title2{

font-size:27px;

font-weight:bold;

color:white;

}

.title3{

font-size:27px;

font-weight:bold;

color:white;

}

.title4{

font-size:27px;

font-weight:bold;

color:white;

}

/* الأزرار */

.stButton>button{

width:100%;

height:85px;

font-size:24px;

font-weight:bold;

border-radius:18px;

background:#6D4C41;

color:white;

border:2px solid gold;

transition:.3s;

}

.stButton>button:hover{

background:#8D6E63;

transform:scale(1.02);

box-shadow:0 0 18px gold;

}

/* الكتابة */

label,p,h1,h2,h3,h4,h5,h6,span{

color:white !important;

font-weight:bold;

}

/* الشريط السفلى */

.news-bar{

position:fixed;

bottom:0;

right:0;

width:100%;

height:45px;

background:#1A120B;

border-top:2px solid gold;

overflow:hidden;

z-index:999999;

}

.news-text{

position:absolute;

right:-100%;

white-space:nowrap;

line-height:45px;

font-size:20px;

font-weight:bold;

animation:scrollText 35s linear infinite;

}

@keyframes scrollText{

from{

right:-100%;

}

to{

right:100%;

}

}

</style>

""",unsafe_allow_html=True)

# =====================================
# الميزان بالخلفية
# =====================================

st.markdown("""

<div class="watermark">

⚖️

</div>

""",unsafe_allow_html=True)

# ======================================
# رأس البرنامج
# =====================================

st.markdown("""

<div class="header">

<div class="logo">
⚖️
</div>

<div class="title1">
           الهيئة القومية للتأمين الاجتماعى
</div>

<div class="title2">
الإدارة المركزية للشئون القانونية
</div>

<div class="title3">
الإدارة العامة للقضايا
</div>

<div class="title4">
ديوان عام منطقة البحيرة
</div>

</div>

<br>

""",unsafe_allow_html=True)

# =====================================
# الشريط السفلى
# =====================================

st.markdown("""

<div class="news-bar">

<div class="news-text">

<span style="color:#FFD700;">
✨ مع تحيات / وليد حماد
</span>

<span style="color:white;">
&nbsp;&nbsp;|&nbsp;&nbsp;
</span>

<span style="color:#00FFFF;">
الإدارة العامة للقضايا
</span>

<span style="color:white;">
&nbsp;&nbsp;|&nbsp;&nbsp;
</span>

<span style="color:#7CFC00;">
الإدارة المركزية للشئون القانونية
</span>

<span style="color:white;">
&nbsp;&nbsp;|&nbsp;&nbsp;
</span>

<span style="color:#FFA500;">
ديوان عام منطقة البحيرة
</span>

<span style="color:white;">
&nbsp;&nbsp;|&nbsp;&nbsp;
</span>

<span style="color:#FF4444;">
الهيئة القومية للتأمين الاجتماعى
</span>

</div>

</div>

""",unsafe_allow_html=True)

# =====================================
# الصفحة الرئيسية
# =====================================

if st.session_state.page == "home":

    st.markdown("<br>", unsafe_allow_html=True)

    c1, c2, c3 = st.columns([1,3,1])

    with c2:

        if st.button(
            "⚖️ تسجيل القضايا",
            use_container_width=True
        ):
            st.session_state.page = "cases"
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button(
            "📋 الحصر العام",
            use_container_width=True
        ):
            st.session_state.page = "inventory"
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button(
            "🔔 التنبيهات",
            use_container_width=True
        ):
            st.session_state.page = "notifications"
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button(
            "📊 التقارير",
            use_container_width=True
        ):
            st.session_state.page = "reports"
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button(
            "🗄️ الأرشيف",
            use_container_width=True
        ):
            st.session_state.page = "archive"
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button(
            "📚 المكتبة القانونية",
            use_container_width=True
        ):
            st.session_state.page = "library"
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button(
            "🔍 البحث عن دعوى",
            use_container_width=True
        ):
            st.session_state.page = "search"
            st.rerun()


# =====================================
# صفحات تحت الإنشاء
# =====================================

elif st.session_state.page == "notifications":

    st.info("🚧 سيتم تنفيذ قسم التنبيهات فى المرحلة القادمة.")

    if st.button("⬅ العودة للرئيسية"):
        st.session_state.page = "home"
        st.rerun()

elif st.session_state.page == "archive":

    st.info("🚧 سيتم تنفيذ قسم الأرشيف فى المرحلة القادمة.")

    if st.button("⬅ العودة للرئيسية"):
        st.session_state.page = "home"
        st.rerun()


elif st.session_state.page == "library":

    st.info("🚧 سيتم تنفيذ قسم المكتبة القانونية فى المرحلة القادمة.")

    if st.button("⬅ العودة للرئيسية"):
        st.session_state.page = "home"
        st.rerun()


elif st.session_state.page == "search":

    st.info("🚧 سيتم تنفيذ البحث عن الدعاوى فى المرحلة القادمة.")

    if st.button("⬅ العودة للرئيسية"):
        st.session_state.page = "home"
        st.rerun()

# ==========================================================
# تسجيل القضايا
# ==========================================================

elif st.session_state.page == "cases":

    rk = st.session_state.form_reset

    st.markdown("""

    <div style="
    background:#4E342E;
    padding:15px;
    border-radius:15px;
    border:2px solid gold;
    text-align:center;
    color:white;
    font-size:30px;
    font-weight:bold;
    box-shadow:0 0 15px rgba(255,215,0,.4);
    ">

    ⚖️ تسجيل قضية جديدة

    </div>

    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    # =====================================
    # العمود الأول
    # =====================================

    with col1:

        case_type = st.selectbox(
            "نوع الدعوى",
            [
                "دعوى",
                "استئناف",
                "طعن"
            ],
            key=f"c_type_{rk}"
        )

        court_type = st.selectbox(
            "المحكمة",
            [
                "الابتدائية",
                "الاستئناف",
                "النقض",
                "الإدارية",
                "القضاء الإدارى",
                "الإدارية العليا"
            ],
            key=f"ct_type_{rk}"
        )

        court_name = st.text_input(
            "اسم المحكمة",
            key=f"ct_name_{rk}"
        )

        mission = ""

        if case_type == "استئناف":

            mission = st.text_input(
                "المأمورية",
                key=f"mission_{rk}"
            )

        case_number = st.text_input(
            "رقم الدعوى / الاستئناف / الطعن",
            key=f"c_num_{rk}"
        )

        judicial_year = st.text_input(
            "السنة القضائية",
            key=f"j_year_{rk}"
        )

        circuit = st.text_input(
            "الدائرة",
            key=f"circuit_{rk}"
        )

        case_category = st.text_input(
            "النوع",
            key=f"c_cat_{rk}"
        )

    # =====================================
    # العمود الثانى
    # =====================================

    with col2:

        plaintiff = st.text_input(
            "اسم المدعى / المستأنف / الطاعن",
            key=f"pl_{rk}"
        )

        defendant = st.text_input(
            "اسم المدعى عليه / المستأنف ضده / المطعون ضده",
            key=f"def_{rk}"
        )

        subject = st.text_area(
            "موضوع الدعوى",
            height=130,
            key=f"subj_{rk}"
        )

        first_session_date = st.date_input(
            "تاريخ أول جلسة",
            key=f"fs_date_{rk}"
        )

        roll_number = st.text_input(
            "الرول",
            key=f"roll_{rk}"
        )

        first_procedure = st.text_area(
            "سبب الجلسة",
            height=120,
            key=f"f_proc_{rk}"
        )

        notes = st.text_area(
            "ملاحظات",
            height=120,
            key=f"notes_{rk}"
        )

    st.markdown("<hr>", unsafe_allow_html=True)

    # =====================================
    # تنبيهات الواتساب
    # =====================================

    whatsapp_enabled = st.checkbox(
        "📲 تفعيل التنبيهات عبر واتساب",
        key=f"wa_en_{rk}"
    )

    whatsapp_number = ""

    if whatsapp_enabled:

        whatsapp_number = st.text_input(
            "رقم هاتف واتساب",
            key=f"wa_num_{rk}"
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # =====================================
    # رفع المستندات
    # =====================================

    if case_type == "دعوى":
        default_doc = "صحيفة الدعوى"
    elif case_type == "استئناف":
        default_doc = "صحيفة الاستئناف"
    else:
        default_doc = "صحيفة الطعن"

    doc_choices = ["صحيفة الدعوى", "صحيفة الاستئناف", "صحيفة الطعن"]

    document_type = st.selectbox(

        "نوع المستند",

        doc_choices,

        index=doc_choices.index(default_doc),

        key=f"doc_type_{rk}"

    )

    uploaded_file = st.file_uploader(

        "تحميل المستند",

        type=[
            "pdf",
            "doc",
            "docx"
        ],

        key=f"up_file_{rk}"

    )

    st.markdown("<br>", unsafe_allow_html=True)

    # =====================================
    # أزرار الحفظ والحذف
    # =====================================

    col_save, col_delete = st.columns(2)

    # =====================================
    # حفظ القضية محلياً
    # =====================================

    with col_save:

        if st.button(
            "💾 حفظ القضية",
            use_container_width=True
        ):

            if case_number.strip() == "":
                st.error("برجاء إدخال رقم القضية")
                st.stop()

            if court_name.strip() == "":
                st.error("برجاء إدخال اسم المحكمة")
                st.stop()

            if plaintiff.strip() == "":
                st.error("برجاء إدخال اسم المدعى")
                st.stop()

            if whatsapp_enabled and whatsapp_number.strip() != "":
                valid_prefix = (
                    whatsapp_number.startswith("010") or
                    whatsapp_number.startswith("011") or
                    whatsapp_number.startswith("012") or
                    whatsapp_number.startswith("015")
                )
                if len(whatsapp_number.strip()) != 11 or not valid_prefix:
                    st.error(
                        "رقم واتساب غير صحيح (يجب أن يكون 11 رقم ويبدأ بـ 010 أو 011 أو 012 أو 015)"
                    )
                    st.stop()

            try:
                with conn:
                    cur.execute("""

                    INSERT INTO cases(

                    case_type,
                    court_type,
                    court_name,
                    mission,
                    case_number,
                    judicial_year,
                    circuit,
                    case_category,
                    plaintiff,
                    defendant,
                    subject,
                    notes,
                    whatsapp_enabled,
                    whatsapp_number,
                    created_at

                    )

                    VALUES(

                    ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?

                    )

                    """,

                    (

                    case_type.strip(),
                    court_type.strip(),
                    court_name.strip(),
                    mission.strip(),
                    case_number.strip(),
                    judicial_year.strip(),
                    circuit.strip(),
                    case_category.strip(),
                    plaintiff.strip(),
                    defendant.strip(),
                    subject.strip(),
                    notes.strip(),
                    int(whatsapp_enabled),
                    whatsapp_number.strip(),
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )

                    )

                    )

                    case_id = cur.lastrowid

                    cur.execute("""

                    INSERT INTO sessions(

                    case_id,
                    session_date,
                    roll_number,
                    procedure,
                    created_at

                    )

                    VALUES(

                    ?,?,?,?,?

                    )

                    """,

                    (

                    case_id,
                    str(first_session_date),
                    roll_number.strip(),
                    first_procedure.strip(),
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )

                    )

                    )

                    if uploaded_file is not None:

                        safe_folder = f"{case_number.strip()}-{judicial_year.strip() or 'عام'}".replace("/", "_")
                        case_folder = os.path.join(DOCUMENTS_DIR, safe_folder)
                        os.makedirs(case_folder, exist_ok=True)

                        file_path = os.path.join(
                            case_folder,
                            uploaded_file.name
                        )

                        with open(file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())

                        cur.execute("""

                        INSERT INTO documents(

                        case_id,
                        document_type,
                        document_description,
                        file_name,
                        file_path,
                        uploaded_at

                        )

                        VALUES(

                        ?,?,?,?,?,?

                        )

                        """,

                        (

                        case_id,

                        document_type,

                        f"{document_type} رقم {case_number.strip()}",

                        uploaded_file.name,

                        file_path,

                        datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )

                        )

                        )

                st.success("✅ تم حفظ القضية محلياً بنجاح")
                st.balloons()
                st.session_state.page = "inventory"
                st.rerun()

            except sqlite3.Error as e:
                st.error(f"❌ حدث خطأ أثناء الحفظ المحلي: {e}")

    # =====================================
    # حذف البيانات المدخلة
    # =====================================

    with col_delete:

        if st.button(
            "🗑 حذف البيانات",
            use_container_width=True
        ):

            st.session_state.form_reset += 1
            st.rerun()

    # =====================================
    # أزرار أسفل الصفحة
    # =====================================

    st.markdown("<br>", unsafe_allow_html=True)

    b1, b2 = st.columns(2)

    with b1:

        if st.button(
            "⬅ العودة للرئيسية",
            use_container_width=True,
            key="back_home_cases"
        ):

            st.session_state.page = "home"
            st.rerun()

    with b2:

        if st.button(
            "🧹 مسح جميع البيانات",
            use_container_width=True,
            key="clear_form"
        ):

            st.session_state.form_reset += 1
            st.warning("تم مسح البيانات المدخلة.")
            st.rerun()

# ==========================================================
# الحصر العام
# ==========================================================

elif st.session_state.page == "inventory":

    st.markdown("""
    <div style="
    background:#4E342E;
    border:2px solid gold;
    border-radius:15px;
    padding:15px;
    text-align:center;
    color:white;
    font-size:30px;
    font-weight:bold;
    box-shadow:0 0 15px rgba(255,215,0,.4);
    ">
    📋 الحصر العام للقضايا
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # =====================================
    # الإحصائيات
    # =====================================

    cur.execute("SELECT COUNT(*) FROM cases")
    total_cases = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM sessions")
    total_sessions = cur.fetchone()[0]

    c1, c2 = st.columns(2)

    with c1:

        st.metric(
            "⚖️ إجمالى القضايا",
            total_cases
        )

    with c2:

        st.metric(
            "📅 إجمالى الجلسات",
            total_sessions
        )

    st.markdown("---")

    # =====================================
    # البحث
    # =====================================

    search = st.text_input(
        "🔍 البحث فى القضايا"
    )

    # =====================================
    # الفلاتر
    # =====================================

    f1, f2 = st.columns(2)

    with f1:

        case_filter = st.selectbox(

            "نوع الدعوى",

            [

                "الكل",

                "دعوى",

                "استئناف",

                "طعن"

            ]

        )

    with f2:

        court_filter = st.text_input(
            "المحكمة"
        )

    # =====================================
    # قراءة القضايا
    # =====================================

    cur.execute("""

    SELECT

    c.id,
    c.case_type,
    c.case_number,
    c.judicial_year,
    c.circuit,
    c.case_category,
    c.court_name,
    c.mission,
    c.plaintiff,
    c.defendant,
    c.subject,

    (

        SELECT session_date

        FROM sessions s

        WHERE s.case_id=c.id

        ORDER BY session_date DESC, id DESC

        LIMIT 1

    ) last_session,

    (

        SELECT procedure

        FROM sessions s

        WHERE s.case_id=c.id

        ORDER BY session_date DESC, id DESC

        LIMIT 1

    ) last_procedure

    FROM cases c

    ORDER BY c.created_at DESC, c.id DESC

    """)

    rows = cur.fetchall()

    # =====================================
    # تطبيق البحث والفلاتر
    # =====================================

    filtered = []

    for row in rows:

        text = " ".join(
            [str(x) if x else "" for x in tuple(row)]
        ).lower()

        if search.strip():

            if search.lower() not in text:
                continue

        if case_filter != "الكل":

            if row["case_type"] != case_filter:
                continue

        if court_filter.strip():

            if court_filter.lower() not in (row["court_name"] or "").lower():
                continue

        filtered.append(row)

    st.markdown("<br>", unsafe_allow_html=True)

    if len(filtered) == 0:

        st.warning("لا توجد نتائج")

    else:

        st.success(f"عدد النتائج : {len(filtered)}")

        st.markdown("<br>", unsafe_allow_html=True)

    # =====================================
    # عرض القضايا
    # =====================================

    for row in filtered:

        case_id = row["id"]

        plaintiff = (
            row["plaintiff"].replace(
                "الهيئة القومية للتأمين الاجتماعى",
                "الهيئة"
            )
            if row["plaintiff"] else ""
        )

        defendant = (
            row["defendant"].replace(
                "الهيئة القومية للتأمين الاجتماعى",
                "الهيئة"
            )
            if row["defendant"] else ""
        )

        st.markdown(f"""
        <div style="
        background:#4E342E;
        border:2px solid #D4AF37;
        border-radius:18px;
        padding:18px;
        margin-bottom:18px;
        box-shadow:0 0 10px rgba(212,175,55,.35);
        ">

        <div style="
        color:#FFD700;
        font-size:24px;
        font-weight:bold;
        ">

        ⚖️ {row["case_type"]} رقم {row["case_number"]} لسنة {row["judicial_year"]}

        </div>

        <hr style="border:1px solid #8B7355;">

        <div style="color:white;font-size:18px;">

        🏛 <b>المحكمة :</b> {row["court_name"]}

        </div>

        <div style="color:white;font-size:18px;">

        ⚖️ <b>الدائرة :</b> {row["circuit"]}

        </div>

        <div style="color:white;font-size:18px;">

        📂 <b>النوع :</b> {row["case_category"]}

        </div>

        <div style="color:white;font-size:18px;">

        👤 <b>الخصوم :</b><br>

        {plaintiff}<br>

        ضد<br>

        {defendant}

        </div>

        <div style="color:#00FFFF;font-size:18px;">

        📄 <b>الموضوع :</b><br>

        {row["subject"]}

        </div>

        <div style="color:#7CFC00;font-size:18px;">

        📅 <b>آخر جلسة :</b>

        {row["last_session"] if row["last_session"] else "لا يوجد"}

        </div>

        <div style="color:#FFA500;font-size:18px;">

        ⚖️ <b>آخر إجراء :</b><br>

        {row["last_procedure"] if row["last_procedure"] else "لا يوجد"}

        </div>

        </div>

        """, unsafe_allow_html=True)

        b1, b2, b3 = st.columns(3)

        with b1:

            if st.button(
                "📂 فتح القضية",
                key=f"open_{case_id}",
                use_container_width=True
            ):

                st.session_state.selected_case = case_id
                st.session_state.page = "case_details"
                st.rerun()

        with b2:

            if st.button(
                "✏️ تعديل",
                key=f"edit_{case_id}",
                use_container_width=True
            ):

                st.session_state.selected_case = case_id
                st.session_state.page = "edit_case"
                st.rerun()

        with b3:

            if st.button(
                "🗑 حذف",
                key=f"delete_{case_id}",
                use_container_width=True
            ):

                st.session_state.delete_case_id = case_id
                st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

    # =====================================
    # تأكيد حذف القضية محلياً
    # =====================================

    if st.session_state.delete_case_id:

        st.markdown("---")

        st.error("⚠️ هل تريد حذف القضية نهائياً؟")

        d1, d2 = st.columns(2)

        with d1:

            if st.button(
                "✅ نعم - حذف القضية",
                use_container_width=True,
                key="confirm_delete_case"
            ):

                delete_id = st.session_state.delete_case_id

                cur.execute("SELECT file_path FROM documents WHERE case_id=?", (delete_id,))
                for d_row in cur.fetchall():
                    if d_row["file_path"] and os.path.exists(d_row["file_path"]):
                        try:
                            os.remove(d_row["file_path"])
                        except OSError:
                            pass

                with conn:
                    cur.execute(
                        "DELETE FROM sessions WHERE case_id=?",
                        (delete_id,)
                    )

                    cur.execute(
                        "DELETE FROM documents WHERE case_id=?",
                        (delete_id,)
                    )

                    cur.execute(
                        "DELETE FROM cases WHERE id=?",
                        (delete_id,)
                    )

                st.success("تم حذف القضية محلياً بنجاح")

                st.session_state.delete_case_id = None

                st.rerun()

        with d2:

            if st.button(
                "❌ إلغاء",
                use_container_width=True,
                key="cancel_delete_case"
            ):

                st.session_state.delete_case_id = None

                st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # =====================================
    # العودة للرئيسية
    # =====================================

    if st.button(
        "⬅ العودة للرئيسية",
        use_container_width=True,
        key="inventory_home"
    ):

        st.session_state.page = "home"

        st.rerun()

# ==========================================================
# ملف القضية
# ==========================================================

elif st.session_state.page == "case_details":

    case_id = st.session_state.selected_case

    cur.execute("""

    SELECT *

    FROM cases

    WHERE id=?

    """,(case_id,))

    case = cur.fetchone()

    if not case:

        st.error("❌ القضية غير موجودة")

        if st.button("⬅ العودة للحصر العام"):

            st.session_state.page="inventory"

            st.rerun()

    else:

        st.markdown("""

        <div style="
        background:#4E342E;
        border:2px solid gold;
        border-radius:18px;
        padding:18px;
        text-align:center;
        color:white;
        font-size:32px;
        font-weight:bold;
        ">

        ⚖️ ملف القضية

        </div>

        """,unsafe_allow_html=True)

        st.markdown("<br>",unsafe_allow_html=True)

        # =====================================
        # بيانات القضية
        # =====================================

        st.markdown("""
        <h3 style='color:#FFD700'>
        📋 بيانات القضية
        </h3>
        """,unsafe_allow_html=True)

        t1,t2=st.columns(2)

        with t1:

            st.info(f"""
⚖️ نوع الدعوى

{case["case_type"]}
""")

            st.info(f"""
🏛 المحكمة

{case["court_name"]}
""")

            st.info(f"""
📄 رقم الدعوى

{case["case_number"]}
""")

            st.info(f"""
📅 السنة القضائية

{case["judicial_year"]}
""")

            st.info(f"""
⚖️ الدائرة

{case["circuit"]}
""")

            if case["mission"]:

                st.info(f"""
📍 المأمورية

{case["mission"]}
""")

        with t2:

            st.success(f"""
👤 المدعى

{case["plaintiff"]}
""")

            st.error(f"""
👤 المدعى عليه

{case["defendant"]}
""")

            st.warning(f"""
📑 موضوع الدعوى

{case["subject"]}
""")

            if case["notes"]:

                st.info(f"""
📝 الملاحظات

{case["notes"]}
""")

        st.markdown("<hr>",unsafe_allow_html=True)

        st.markdown("""
        <h3 style='color:#FFD700'>
        📅 جدول الجلسات
        </h3>
        """,unsafe_allow_html=True)

        # =====================================
        # قراءة الجلسات
        # =====================================

        cur.execute("""

        SELECT

        id,
        roll_number,
        session_date,
        procedure

        FROM sessions

        WHERE case_id=?

        ORDER BY session_date ASC, id ASC

        """,(case_id,))

        sessions=cur.fetchall()

        if sessions:

            st.markdown("""
            <style>

            table{
                width:100%;
                border-collapse:collapse;
            }

            th{
                background:#5D4037;
                color:#FFD700;
                text-align:center;
                padding:10px;
                border:1px solid gold;
            }

            td{
                background:#3E2723;
                color:white;
                text-align:center;
                padding:8px;
                border:1px solid #8D6E63;
            }

            </style>
            """,unsafe_allow_html=True)

            table="""

            <table>

            <tr>

            <th>م</th>

            <th>الرول</th>

            <th>تاريخ الجلسة</th>

            <th>الإجراءات</th>

            </tr>

            """

            for i,s in enumerate(sessions,start=1):

                table+=f"""

                <tr>

                <td>{i}</td>

                <td>{s["roll_number"]}</td>

                <td>{s["session_date"]}</td>

                <td>{s["procedure"]}</td>

                </tr>

                """

            table+="</table>"

            st.markdown(
                table,
                unsafe_allow_html=True
            )

        else:

            st.warning("لا توجد جلسات مسجلة")

        st.markdown("<br>",unsafe_allow_html=True)

        st.markdown("""
        <h3 style='color:#FFD700'>
        ⚖️ متابعة الجلسات
        </h3>
        """,unsafe_allow_html=True)

        c1,c2=st.columns(2)

        with c1:

            next_session_date=st.date_input(
                "تاريخ الجلسة القادمة"
            )

            next_roll=st.text_input(
                "الرول الجديد"
            )

        with c2:

            next_procedure=st.text_area(
                "سبب التأجيل / الإجراء",
                height=120
            )

        if st.button(
            "💾 حفظ الجلسة الجديدة",
            use_container_width=True
        ):

            with conn:
                cur.execute("""

                INSERT INTO sessions(

                case_id,
                session_date,
                roll_number,
                procedure,
                created_at

                )

                VALUES(

                ?,?,?,?,?

                )

                """,

                (

                case_id,

                str(next_session_date),

                next_roll.strip(),

                next_procedure.strip(),

                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

                )

                )

            st.success("تم حفظ الجلسة محلياً بنجاح")

            st.rerun()

        # ====================================
        # تعديل جلسة سابقة
        # =====================================

        st.markdown("---")

        st.markdown("""
        <h3 style='color:#FFD700'>
        ✏️ تعديل جلسة
        </h3>
        """, unsafe_allow_html=True)

        cur.execute("""

        SELECT

        id,
        session_date,
        roll_number

        FROM sessions

        WHERE case_id=?

        ORDER BY session_date ASC, id ASC

        """, (case_id,))

        session_list = cur.fetchall()

        if session_list:

            session_labels = {
                s["id"]: f"{s['session_date']} - {s['roll_number'] or 'بدون رول'}"
                for s in session_list
            }

            selected_session = st.selectbox(
                "اختر الجلسة",
                options=list(session_labels.keys()),
                format_func=lambda sid: session_labels[sid]
            )

            cur.execute("""

            SELECT

            roll_number,
            session_date,
            procedure

            FROM sessions

            WHERE id=?

            """, (selected_session,))

            data = cur.fetchone()

            edit_roll = st.text_input(
                "الرول",
                value=data["roll_number"] or "",
                key=f"edit_roll_{selected_session}"
            )

            edit_date = st.date_input(
                "تاريخ الجلسة",
                value=datetime.strptime(
                    data["session_date"],
                    "%Y-%m-%d"
                ),
                key=f"edit_date_{selected_session}"
            )

            edit_procedure = st.text_area(
                "الإجراءات",
                value=data["procedure"] or "",
                height=120,
                key=f"edit_proc_{selected_session}"
            )

            if st.button(
                "💾 حفظ التعديل",
                use_container_width=True
            ):

                with conn:
                    cur.execute("""

                    UPDATE sessions

                    SET

                    roll_number=?,
                    session_date=?,
                    procedure=?

                    WHERE id=?

                    """,

                    (

                    edit_roll.strip(),

                    str(edit_date),

                    edit_procedure.strip(),

                    selected_session

                    )

                    )

                st.success("تم تعديل الجلسة محلياً بنجاح")

                st.rerun()

        else:

            st.info("لا توجد جلسات للتعديل")

        st.markdown("---")

        # =====================================
        # مستندات القضية
        # =====================================

        st.markdown("""
        <h3 style='color:#FFD700'>
        📂 مستندات القضية
        </h3>
        """, unsafe_allow_html=True)

        document_type = st.selectbox(

            "نوع المستند",

            [

                "صحيفة دعوى",

                "صحيفة استئناف",

                "صحيفة طعن",

                "مذكرة دفاع",

                "حافظة مستندات",

                "تقرير خبير",

                "تقرير طب شرعى",

                "تقرير لجنة طبية",

                "صحيفة تجديد من الشطب",

                "صحيفة تعجيل من الوقف",

                "صورة حكم تمهيدى",

                "أخرى"

            ]

        )

        document_description = st.text_input(
            "بيان المستند"
        )

        uploaded_document = st.file_uploader(
            "رفع المستند",
            key="case_details_doc"
        )

        # =====================================
        # حفظ المستند محلياً
        # =====================================

        if st.button(
            "💾 حفظ المستند",
            use_container_width=True
        ):

            if uploaded_document is None:

                st.error("يرجى اختيار مستند")

            else:

                safe_folder = f"{case['case_number']}-{case['judicial_year'] or 'عام'}".replace("/", "_")
                case_folder = os.path.join(DOCUMENTS_DIR, safe_folder)
                os.makedirs(case_folder, exist_ok=True)

                file_path = os.path.join(
                    case_folder,
                    uploaded_document.name
                )

                with open(
                    file_path,
                    "wb"
                ) as f:

                    f.write(
                        uploaded_document.getbuffer()
                    )

                with conn:
                    cur.execute("""

                    INSERT INTO documents(

                    case_id,
                    document_type,
                    document_description,
                    file_name,
                    file_path,
                    uploaded_at

                    )

                    VALUES(?,?,?,?,?,?)

                    """,

                    (

                    case_id,

                    document_type,

                    document_description.strip(),

                    uploaded_document.name,

                    file_path,

                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )

                    )

                    )

                st.success("تم حفظ المستند محلياً بنجاح")

                st.rerun()

        st.markdown("---")

        # =====================================
        # عرض مستندات القضية
        # =====================================

        st.markdown("""
        <h3 style='color:#FFD700'>
        📂 المستندات المحفوظة
        </h3>
        """, unsafe_allow_html=True)

        cur.execute("""

        SELECT

        id,
        document_type,
        document_description,
        file_name,
        file_path

        FROM documents

        WHERE case_id=?

        ORDER BY uploaded_at DESC

        """,(case_id,))

        docs = cur.fetchall()

        if docs:

            for doc in docs:

                with st.expander(
                    f"📄 {doc['document_type']}"
                ):

                    st.write(
                        f"**البيان :** {doc['document_description']}"
                    )

                    st.write(
                        f"**اسم الملف :** {doc['file_name']}"
                    )

                    try:

                        with open(
                            doc["file_path"],
                            "rb"
                        ) as f:

                            st.download_button(

                                "⬇ تحميل المستند",

                                f,

                                file_name=doc["file_name"],

                                use_container_width=True,

                                key=f"download_{doc['id']}"

                            )

                    except:

                        st.error(
                            "الملف غير موجود"
                        )

        else:

            st.info(
                "لا توجد مستندات"
            )

        st.markdown("---")

        # =====================================
        # جلسة الحكم
        # =====================================

        st.markdown("""
        <h3 style='color:#FFD700'>
        ⚖️ بيانات الحكم
        </h3>
        """, unsafe_allow_html=True)

        judgment_date = st.date_input(
            "تاريخ جلسة الحكم"
        )

        judgment_text = st.text_area(
            "منطوق الحكم",
            height=150
        )

        judgment_result = st.selectbox(
            "نتيجة الحكم",
            [
                "لصالح الهيئة",
                "ضد الهيئة"
            ]
        )

        if st.button(
            "💾 حفظ الحكم",
            use_container_width=True
        ):

            with conn:
                cur.execute("""
                INSERT INTO sessions(
                    case_id,
                    session_date,
                    roll_number,
                    procedure,
                    judgment_result,
                    created_at
                )
                VALUES(
                    ?,?,?,?,?,?
                )
                """,
                (
                    case_id,
                    str(judgment_date),
                    "حكم",
                    judgment_text.strip(),
                    judgment_result,
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                )
                )

            st.success("تم حفظ بيانات الحكم محلياً")

            st.rerun()

# ==========================================================
# قسم التقارير
# ==========================================================

elif st.session_state.page == "reports":

    rrk = st.session_state.report_reset

    st.markdown("""
    <div class="report-card">

    <div style="display:flex;
    justify-content:space-between;
    align-items:center;">

    <div class="logo-box">

    ⚖️

    </div>

    <div>

    <div class="report-title">

    📊 التقارير

    </div>

    <div class="gov-title">

    الهيئة القومية للتأمين الاجتماعى<br>

    الإدارة المركزية للشئون القانونية<br>

    الإدارة العامة للقضايا<br>

    ديوان عام منطقة ............

    </div>

    </div>

    <div style="width:90px;"></div>

    </div>

    <div class="gold-line"></div>

    </div>

    """,unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("""
    <div class="report-card">
    <div style="color:#FFD56A;
    font-size:24px;
    font-weight:bold;
    margin-bottom:20px;">
    🔎 خيارات التقرير
    </div>
    """, unsafe_allow_html=True)

    r1, r2 = st.columns(2)

    with r1:

        report_type = st.selectbox(
            "نوع التقرير",
            [

                "بيان بجميع الدعاوى المتداولة",

                "بيان الدعاوى حسب موضوع الدعوى",

                "بيان بجميع الأحكام الصادرة",

                "بيان بالأحكام الصادرة للصالح",

                "بيان بالأحكام الصادرة للضد",

                "بيان الأحكام الصادرة حسب موضوع الدعوى",

                "بيان عددى بجميع الدعاوى المتداولة",

                "بيان عددى بالدعاوى حسب موضوع الدعوى",

                "بيان عددى بجميع الأحكام الصادرة",

                "بيان عددى بالأحكام الصادرة للصالح",

                "بيان عددى بالأحكام الصادرة للضد",

                "بيان عددى بالأحكام الصادرة حسب موضوع الدعوى"

            ],
            key=f"rep_type_{rrk}"
        )

    with r2:

        subject_search = ""

        if report_type in (

            "بيان الدعاوى حسب موضوع الدعوى",

            "بيان الأحكام الصادرة حسب موضوع الدعوى",

            "بيان عددى بالدعاوى حسب موضوع الدعوى",

            "بيان عددى بالأحكام الصادرة حسب موضوع الدعوى"

        ):

            subject_search = st.text_input(
                "موضوع الدعوى",
                key=f"subj_search_{rrk}"
            )

        else:

            st.empty()

    st.markdown("<br>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:

        office = st.text_input(
            "ديوان عام المنطقة",
            value="",
            key=f"rep_office_{rrk}"
        )

    with c2:

        lawyer = st.text_input(
            "طرف الأستاذ / المحامى",
            value="",
            key=f"rep_lawyer_{rrk}"
        )

    st.markdown("<br>", unsafe_allow_html=True)

    d1, d2 = st.columns(2)

    with d1:

        from_date = st.date_input(
            "من تاريخ",
            key=f"rep_from_{rrk}"
        )

    with d2:

        to_date = st.date_input(
            "إلى تاريخ",
            key=f"rep_to_{rrk}"
        )

    st.markdown("<br>", unsafe_allow_html=True)

    b1, b2, b3 = st.columns([2,2,1])

    with b1:

        if st.button(
            "📊 عرض التقرير",
            use_container_width=True,
            type="primary",
            key=f"show_rep_{rrk}"
        ):
            st.session_state.show_report_flag = True

    with b2:

        if st.button(
            "🧹 مسح البيانات",
            use_container_width=True,
            key=f"clear_rep_{rrk}"
        ):
            st.session_state.report_reset += 1
            st.session_state.show_report_flag = False
            st.rerun()

    with b3:

        if st.button(
            "🔄",
            use_container_width=True,
            key=f"ref_rep_{rrk}"
        ):
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.show_report_flag:

        # ===========================================
        # حساب الإحصائيات من قاعدة البيانات المحلية
        # ===========================================

        total_cases = 0
        total_judgments = 0
        total_for = 0
        total_against = 0

        try:

            cur.execute("SELECT COUNT(*) FROM cases")
            total_cases = cur.fetchone()[0]

            cur.execute("""
            SELECT COUNT(*)
            FROM sessions
            WHERE roll_number='حكم'
            """)
            total_judgments = cur.fetchone()[0]

            cur.execute("""
            SELECT COUNT(*)
            FROM sessions
            WHERE roll_number='حكم' AND judgment_result='لصالح الهيئة'
            """)
            total_for = cur.fetchone()[0]

            cur.execute("""
            SELECT COUNT(*)
            FROM sessions
            WHERE roll_number='حكم' AND judgment_result='ضد الهيئة'
            """)
            total_against = cur.fetchone()[0]

        except:
            pass

        # ===========================================
        # الكروت
        # ===========================================

        a, b, c, d = st.columns(4)

        with a:

            st.markdown(f"""

            <div class="stat-box">

            <div class="stat-number">

            {total_cases}

            </div>

            <div class="stat-title">

            📂 إجمالى القضايا

            </div>

            </div>

            """, unsafe_allow_html=True)

        with b:

            st.markdown(f"""

            <div class="stat-box">

            <div class="stat-number">

            {total_judgments}

            </div>

            <div class="stat-title">

            ⚖️ الأحكام

            </div>

            </div>

            """, unsafe_allow_html=True)

        with c:

            st.markdown(f"""

            <div class="stat-box">

            <div class="stat-number">

            {total_for}

            </div>

            <div class="stat-title">

            ✅ للصالح

            </div>

            </div>

            """, unsafe_allow_html=True)

        with d:

            st.markdown(f"""

            <div class="stat-box">

            <div class="stat-number">

            {total_against}

            </div>

            <div class="stat-title">

            ❌ للضد

            </div>

            </div>

            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ===========================================
        # عنوان التقرير
        # ===========================================

        st.markdown(f"""

        <div class="report-card">

        <div style="font-size:28px;
        color:#FFD700;
        font-weight:bold;
        text-align:center;">

        {report_type}

        </div>

        <br>

        <div style="display:flex;
        justify-content:space-between;
        color:white;
        font-size:16px;">

        <span>

        📅 من :
        {from_date.strftime('%d/%m/%Y')}

        </span>

        <span>

        📅 إلى :
        {to_date.strftime('%d/%m/%Y')}

        </span>

        </div>

        <br>

        <div style="display:flex;
        justify-content:space-between;
        color:white;
        font-size:16px;">

        <span>

        🏢 ديوان :
        {office if office else '...............'}

        </span>

        <span>

        👤 الأستاذ :
        {lawyer if lawyer else '...............'}

        </span>

        </div>

        </div>

        """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
