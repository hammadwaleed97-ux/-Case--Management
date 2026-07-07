# ============================================================
# إدارة القضايا
# Streamlit Professional Edition
# الجزء الأول : الأساس + التصميم + قاعدة البيانات
# ============================================================

import streamlit as st
import sqlite3
import os
from datetime import datetime


# ============================================================
# إعداد الصفحة
# ============================================================

st.set_page_config(
    page_title="إدارة القضايا",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ============================================================
# المسارات
# ============================================================

DB = "legal_cases.db"

DOC_FOLDER = "documents"

LIB_FOLDER = "library"


for folder in [DOC_FOLDER, LIB_FOLDER]:

    if not os.path.exists(folder):
        os.makedirs(folder)



# ============================================================
# قاعدة البيانات
# ============================================================

def db():

    return sqlite3.connect(DB)



def create_tables():

    conn=db()

    cur=conn.cursor()



    # القضايا

    cur.execute("""
    CREATE TABLE IF NOT EXISTS cases(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        case_type TEXT,

        court_type TEXT,

        court_name TEXT,

        department TEXT,

        case_number TEXT,

        case_year TEXT,

        circle TEXT,

        category TEXT,

        plaintiff TEXT,

        defendant TEXT,

        subject TEXT,

        first_session TEXT,

        roll TEXT,

        required_action TEXT,

        notes TEXT,

        whatsapp INTEGER,

        whatsapp_number TEXT,

        status TEXT DEFAULT 'متداولة',

        judgment_date TEXT,

        judgment_text TEXT,

        judgment_result TEXT,

        created_at TEXT

    )
    """)



    # الجلسات

    cur.execute("""
    CREATE TABLE IF NOT EXISTS sessions(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        case_id INTEGER,

        session_date TEXT,

        roll TEXT,

        action TEXT

    )
    """)



    # المستندات

    cur.execute("""
    CREATE TABLE IF NOT EXISTS documents(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        case_id INTEGER,

        doc_type TEXT,

        doc_name TEXT,

        file_path TEXT

    )
    """)



    # الأرشيف

    cur.execute("""
    CREATE TABLE IF NOT EXISTS archive(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        case_id INTEGER,

        judgment_date TEXT,

        judgment_text TEXT,

        result TEXT,

        appeal_number TEXT,

        save_note TEXT

    )
    """)



    # المكتبة

    cur.execute("""
    CREATE TABLE IF NOT EXISTS library(

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        section TEXT,

        title TEXT,

        file_path TEXT,

        created TEXT

    )
    """)



    conn.commit()

    conn.close()



create_tables()



# ============================================================
# التصميم الفخم
# ============================================================

st.markdown("""

<style>


body{

direction:rtl;

}



.stApp{


background:

linear-gradient(

rgba(4,18,40,.94),

rgba(4,18,40,.94)

),

url("https://upload.wikimedia.org/wikipedia/commons/thumb/8/8f/Scales_of_justice.svg/512px-Scales_of_justice.svg.png");


background-repeat:no-repeat;

background-position:center;

background-size:500px;


}



h1,h2,h3,h4{


color:#FFD700 !important;


}



.card{


background:

linear-gradient(

135deg,

#09203f,

#1b4f8a

);


border:

2px solid #D4AF37;


border-radius:25px;


padding:25px;


text-align:center;


height:170px;


box-shadow:

0 0 25px rgba(212,175,55,.45);


}



.card:hover{


transform:scale(1.04);


}



.icon{


font-size:55px;


}



.card-title{


font-size:22px;

color:white;

font-weight:bold;


}



.footer{


position:fixed;


bottom:10px;


width:100%;


text-align:center;


font-size:18px;


font-weight:bold;


color:#FFD700;


animation:flash 4s infinite;


}



@keyframes flash{


0%{opacity:.2}

50%{opacity:1}

100%{opacity:.2}


}



.stButton button{


background:

linear-gradient(

90deg,

#0b2345,

#D4AF37

);


color:white;


border-radius:15px;


font-weight:bold;


}



</style>

""",
unsafe_allow_html=True)



# ============================================================
# رأس البرنامج
# ============================================================


st.markdown(

"""

<div style="text-align:center">

<div style="font-size:75px;color:#FFD700">

⚖️

</div>


<h1>

إدارة القضايا

</h1>


</div>

""",

unsafe_allow_html=True

)



# ============================================================
# الصفحة الرئيسية
# ============================================================


if "page" not in st.session_state:

    st.session_state.page="الرئيسية"



pages=[

("📝","تسجيل القضايا"),

("📚","الحصر العام"),

("🔔","التنبيهات"),

("📊","التقارير"),

("🗄️","الأرشيف"),

("⚖️","المكتبة القانونية"),

("🔎","البحث عن دعوى")

]



cols=st.columns(4)



for i,(icon,name) in enumerate(pages):


    with cols[i%4]:


        st.markdown(

        f"""

        <div class="card">


        <div class="icon">

        {icon}

        </div>


        <div class="card-title">

        {name}

        </div>


        </div>

        """,

        unsafe_allow_html=True

        )


        if st.button(

            name,

            key=f"nav{i}",

            use_container_width=True

        ):

            st.session_state.page=name



# ============================================================
# الإحصائيات
# ============================================================


st.divider()


c1,c2,c3,c4=st.columns(4)


for col,title,value in [

(c1,"القضايا المتداولة",0),

(c2,"الأحكام الصادرة",0),

(c3,"للصالح",0),

(c4,"للضد",0)

]:

    col.metric(title,value)



# ============================================================
# التوقيع
# ============================================================


st.markdown(

"""

<div class="footer">


مع تحيات / وليد شعبان حماد

<br>

الإدارة العامة للشئون القانونية ـ ديوان عام منطقة البحيرة ـ الهيئة القومية للتأمين الاجتماعى


</div>


""",

unsafe_allow_html=True

            )
# ============================================================
# إدارة القضايا
# الجزء الثاني : تسجيل القضايا
# ============================================================


def save_case_record(data):

    conn=db()

    cur=conn.cursor()


    cur.execute("""

    INSERT INTO cases(

    case_type,
    court_type,
    court_name,
    department,
    case_number,
    case_year,
    circle,
    category,
    plaintiff,
    defendant,
    subject,
    first_session,
    roll,
    required_action,
    notes,
    whatsapp,
    whatsapp_number,
    created_at

    )

    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)

    """,data)


    case_id=cur.lastrowid



    # إضافة أول جلسة تلقائيا

    cur.execute("""

    INSERT INTO sessions

    (

    case_id,

    session_date,

    roll,

    action

    )

    VALUES(?,?,?,?)

    """,

    (

    case_id,

    data[11],

    data[12],

    data[13]

    ))



    conn.commit()

    conn.close()




def registration_page():


    st.markdown(

    """

    <h1 style="text-align:center">

    📝 تسجيل القضايا

    </h1>

    """,

    unsafe_allow_html=True

    )


    st.divider()



    c1,c2,c3=st.columns(3)



    with c1:

        case_type=st.selectbox(

        "نوع الدعوى",

        [

        "دعوى",

        "استئناف",

        "طعن"

        ]

        )



    with c2:

        court_type=st.selectbox(

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



    with c3:

        court_name=st.text_input(

        "اسم المحكمة"

        )



    department=""



    if case_type=="استئناف":

        department=st.text_input(

        "المأمورية"

        )



    c1,c2,c3=st.columns(3)



    with c1:

        case_number=st.text_input(

        "رقم الدعوى / الاستئناف / الطعن"

        )



    with c2:

        case_year=st.text_input(

        "السنة القضائية"

        )



    with c3:

        circle=st.text_input(

        "الدائرة"

        )



    c1,c2=st.columns(2)



    with c1:

        category=st.text_input(

        "النوع"

        )


    with c2:

        first_session=st.date_input(

        "تاريخ أول جلسة"

        )



    plaintiff=st.text_input(

    "اسم المدعي / المستأنف / الطاعن"

    )



    defendant=st.text_input(

    "اسم المدعى عليه / المستأنف ضده / المطعون ضده"

    )



    subject=st.text_area(

    "موضوع الدعوى"

    )



    roll=st.text_input(

    "الرول"

    )



    required_action=st.text_area(

    "الإجراء المطلوب"

    )



    notes=st.text_area(

    "ملاحظات"

    )



    st.subheader(

    "🔔 تنبيهات الواتس اب"

    )



    whatsapp=st.checkbox(

    "تفعيل التنبيهات"

    )


    whatsapp_number=st.text_input(

    "رقم واتس اب"

    )



    st.subheader(

    "📂 تحميل المستندات"

    )



    document_type=st.selectbox(

    "نوع المستند",

    [

    "صحيفة الدعوى",

    "صحيفة الاستئناف",

    "صحيفة الطعن"

    ]

    )



    uploaded=st.file_uploader(

    "اختيار الملف"

    )




    col1,col2=st.columns(2)



    with col1:


        if st.button(

        "💾 حفظ القضية",

        use_container_width=True

        ):


            save_case_record(

            (

            case_type,

            court_type,

            court_name,

            department,

            case_number,

            case_year,

            circle,

            category,

            plaintiff,

            defendant,

            subject,

            str(first_session),

            roll,

            required_action,

            notes,

            1 if whatsapp else 0,

            whatsapp_number,

            str(datetime.now())

            )

            )


            st.success(

            "تم حفظ القضية ونقلها إلى الحصر العام"

            )



    with col2:


        if st.button(

        "🗑 حذف القضية",

        use_container_width=True

        ):


            st.warning(

            "الحذف يتم من إدارة القضايا"

            )



# ============================================================
# تشغيل صفحة تسجيل القضايا
# ============================================================


if st.session_state.page=="تسجيل القضايا":

    registration_page()
