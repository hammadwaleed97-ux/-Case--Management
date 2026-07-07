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
    # ============================================================
# إدارة القضايا
# الجزء الثالث : الحصر العام + فتح القضية + الجلسات
# ============================================================


def get_all_cases():

    conn=db()


    df=pd.read_sql_query(

    """

    SELECT

    id,

    case_number AS رقم_القضية,

    case_year AS السنة_القضائية,

    circle AS الدائرة,

    category AS النوع,

    court_name AS المحكمة,

    department AS المأمورية,

    plaintiff AS المدعي,

    defendant AS المدعى_عليه,

    subject AS موضوع_الدعوى,

    first_session AS آخر_جلسة,

    required_action AS السبب


    FROM cases


    WHERE status='متداولة'


    ORDER BY first_session ASC


    """,

    conn

    )


    conn.close()


    return df




def add_new_session(case_id):


    st.subheader(

    "📅 متابعة الجلسات والإجراءات"

    )



    session_date=st.date_input(

    "تاريخ الجلسة القادمة"

    )


    roll=st.text_input(

    "الرول"

    )


    action=st.text_area(

    "سبب التأجيل / الإجراء"

    )



    if st.button(

    "حفظ الجلسة",

    key=f"session_{case_id}"

    ):


        conn=db()

        cur=conn.cursor()



        cur.execute(

        """

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

        str(session_date),

        roll,

        action

        )

        )



        cur.execute(

        """

        UPDATE cases

        SET

        first_session=?,

        required_action=?

        WHERE id=?

        """,

        (

        str(session_date),

        action,

        case_id

        )

        )



        conn.commit()

        conn.close()



        st.success(

        "تم تحديث الحصر العام"

        )





def show_case(case_id):


    conn=db()



    case=pd.read_sql_query(

    """

    SELECT *

    FROM cases

    WHERE id=?

    """,

    conn,

    params=(case_id,)

    )



    sessions=pd.read_sql_query(

    """

    SELECT

    roll AS الرول,

    session_date AS تاريخ_الجلسة,

    action AS الإجراءات


    FROM sessions


    WHERE case_id=?


    ORDER BY session_date ASC


    """,

    conn,

    params=(case_id,)

    )


    conn.close()



    st.markdown(

    """

    <h2 style="text-align:center">

    ⚖️ ملف القضية

    </h2>

    """,

    unsafe_allow_html=True

    )



    st.dataframe(

    case,

    use_container_width=True,

    hide_index=True

    )



    st.markdown(

    """

    <h3>

    📅 جدول الجلسات والإجراءات

    </h3>

    """,

    unsafe_allow_html=True

    )


    st.dataframe(

    sessions,

    use_container_width=True,

    hide_index=True

    )



    add_new_session(case_id)



    st.divider()



    # المستندات

    st.subheader(

    "📂 مستندات القضية"

    )


    docs=pd.read_sql_query(

    """

    SELECT

    doc_type AS النوع,

    doc_name AS البيان


    FROM documents


    WHERE case_id=?


    """,

    sqlite3.connect(DB),

    params=(case_id,)

    )



    st.dataframe(

    docs,

    use_container_width=True,

    hide_index=True

    )



    # الحكم

    st.subheader(

    "⚖️ بيانات الحكم"

    )


    judgment_date=st.date_input(

    "تاريخ جلسة الحكم",

    key=f"jd{case_id}"

    )


    judgment_text=st.text_area(

    "منطوق الحكم",

    key=f"jt{case_id}"

    )


    result=st.selectbox(

    "النتيجة",

    [

    "للصالح",

    "للضد"

    ],

    key=f"jr{case_id}"

    )



    if st.button(

    "تسجيل الحكم ونقل للأرشيف",

    key=f"jud{case_id}"

    ):


        conn=db()

        cur=conn.cursor()



        cur.execute(

        """

        UPDATE cases

        SET

        status='منتهي',

        judgment_date=?,

        judgment_text=?,

        judgment_result=?


        WHERE id=?

        """,

        (

        str(judgment_date),

        judgment_text,

        result,

        case_id

        )

        )



        cur.execute(

        """

        INSERT INTO archive

        (

        case_id,

        judgment_date,

        judgment_text,

        result

        )

        VALUES(?,?,?,?)

        """,

        (

        case_id,

        str(judgment_date),

        judgment_text,

        result

        )

        )



        conn.commit()

        conn.close()



        st.success(

        "تم أرشفة الحكم"

        )





def general_archive_page():


    st.markdown(

    """

    <h1 style="text-align:center">

    📚 الحصر العام

    </h1>

    """,

    unsafe_allow_html=True

    )


    df=get_all_cases()



    if df.empty:

        st.info(

        "لا توجد قضايا مسجلة"

        )

        return



    st.dataframe(

    df,

    use_container_width=True,

    hide_index=True

    )



    selected=st.selectbox(

    "فتح القضية",

    df["رقم_القضية"].astype(str)

    )



    case_id=int(

    df[df["رقم_القضية"].astype(str)==selected]["id"].iloc[0]

    )


    show_case(case_id)




if st.session_state.page=="الحصر العام":

    general_archive_page()
