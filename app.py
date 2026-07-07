# ============================================================
# إدارة القضايا
# Streamlit Professional Edition
# الجزء الأول : الأساس + التصميم + قاعدة البيانات
# ==========================================================
import streamlit as st
import sqlite3
import os
import pandas as pd
import io
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
    # ============================================================
# إدارة القضايا
# الجزء الرابع : المستندات + الأرشيف + البحث
# ============================================================


# ============================================================
# رفع وحفظ مستندات القضايا
# ============================================================


def upload_case_document(case_id):


    st.subheader(

    "📂 رفع مستندات القضية"

    )



    doc_type=st.selectbox(

    "نوع المستند",

    [

    "صحيفة دعوى",

    "صحيفة استئناف",

    "صحيفة طعن",

    "مذكرة دفاع",

    "حافظة مستندات",

    "تقرير خبير",

    "تقرير طب شرعي",

    "تقرير لجنة طبية",

    "صحيفة تجديد من الشطب",

    "صحيفة تعجيل من الوقف",

    "صورة حكم تمهيدي",

    "أخرى"

    ],

    key=f"dtype{case_id}"

    )



    doc_name=st.text_input(

    "بيان المستند",

    key=f"dname{case_id}"

    )



    file=st.file_uploader(

    "اختيار الملف",

    key=f"file{case_id}"

    )



    if st.button(

    "حفظ المستند",

    key=f"save_doc{case_id}"

    ):



        if file:


            path=os.path.join(

            DOC_FOLDER,

            f"{case_id}_{file.name}"

            )



            with open(path,"wb") as f:

                f.write(

                file.getbuffer()

                )



            conn=db()

            cur=conn.cursor()



            cur.execute(

            """

            INSERT INTO documents

            (

            case_id,

            doc_type,

            doc_name,

            file_path

            )

            VALUES(?,?,?,?)

            """,

            (

            case_id,

            doc_type,

            doc_name,

            path

            )

            )



            conn.commit()

            conn.close()



            st.success(

            "تم حفظ المستند"

            )



# ============================================================
# عرض المستندات
# ============================================================


def display_documents(case_id):


    conn=db()



    docs=pd.read_sql_query(

    """

    SELECT

    id,

    doc_type AS النوع,

    doc_name AS البيان,

    file_path


    FROM documents


    WHERE case_id=?


    """,

    conn,

    params=(case_id,)

    )


    conn.close()



    if not docs.empty:


        st.markdown(

        """

        <h3>

        📁 مستندات القضية

        </h3>

        """,

        unsafe_allow_html=True

        )



        for _,row in docs.iterrows():


            c1,c2=st.columns([4,1])



            with c1:

                st.info(

                f"{row['النوع']} - {row['البيان']}"

                )



            with c2:


                if os.path.exists(row["file_path"]):


                    with open(

                    row["file_path"],

                    "rb"

                    ) as f:


                        st.download_button(

                        "فتح",

                        f,

                        file_name=os.path.basename(

                        row["file_path"]

                        ),

                        key=f"open{row['id']}"

                        )




# ============================================================
# صفحة الأرشيف
# ============================================================


def archive_page():


    st.markdown(

    """

    <h1 style="text-align:center">

    🗄️ الأرشيف

    </h1>

    """,

    unsafe_allow_html=True

    )



    conn=db()



    df=pd.read_sql_query(

    """

    SELECT


    archive.id,

    cases.case_number AS رقم_القضية,

    cases.case_year AS السنة,

    cases.subject AS الموضوع,

    archive.judgment_date AS تاريخ_الحكم,

    archive.judgment_text AS المنطوق,

    archive.result AS النتيجة


    FROM archive


    JOIN cases


    ON archive.case_id=cases.id


    ORDER BY archive.judgment_date DESC


    """,

    conn

    )


    conn.close()



    if df.empty:


        st.warning(

        "لا توجد أحكام بالأرشيف"

        )


    else:


        st.dataframe(

        df,

        use_container_width=True,

        hide_index=True

        )





# ============================================================
# البحث عن دعوى
# ============================================================


def search_case_page():


    st.markdown(

    """

    <h1 style="text-align:center">

    🔎 البحث عن دعوى

    </h1>

    """,

    unsafe_allow_html=True

    )



    c1,c2,c3=st.columns(3)



    with c1:

        name=st.text_input(

        "اسم المدعي"

        )



    with c2:

        number=st.text_input(

        "رقم القضية"

        )



    with c3:

        year=st.text_input(

        "السنة القضائية"

        )



    if st.button(

    "🔍 بحث",

    use_container_width=True

    ):



        conn=db()



        query="""

        SELECT *

        FROM cases

        WHERE 1=1

        """



        params=[]



        if name:

            query+=" AND plaintiff LIKE ?"

            params.append(

            "%"+name+"%"

            )



        if number:

            query+=" AND case_number=?"

            params.append(number)



        if year:

            query+=" AND case_year=?"

            params.append(year)



        result=pd.read_sql_query(

        query,

        conn,

        params=params

        )



        conn.close()



        if result.empty:


            st.error(

            "لا توجد قضية - تأكد من بيانات البحث"

            )


        else:


            st.success(

            "تم العثور على القضية"

            )


            st.dataframe(

            result,

            use_container_width=True,

            hide_index=True

            )




# ============================================================
# تشغيل الصفحات
# ============================================================


if st.session_state.page=="الأرشيف":

    archive_page()



if st.session_state.page=="البحث عن دعوى":

    search_case_page()
    # ============================================================
# إدارة القضايا
# الجزء الخامس : التنبيهات + المكتبة القانونية
# ============================================================

from datetime import timedelta


# ============================================================
# التنبيهات
# ============================================================


def notifications_page():


    st.markdown(

    """

    <h1 style="text-align:center">

    🔔 التنبيهات

    </h1>

    """,

    unsafe_allow_html=True

    )



    today=datetime.now().date()



    week_date=today + timedelta(days=7)



    conn=db()



    sessions=pd.read_sql_query(

    """

    SELECT


    cases.id,

    cases.case_number AS رقم_القضية,

    cases.case_year AS السنة,

    cases.plaintiff AS المدعي,

    cases.defendant AS المدعى_عليه,

    cases.subject AS الموضوع,

    sessions.session_date AS تاريخ_الجلسة,

    sessions.action AS الإجراء



    FROM sessions


    JOIN cases


    ON sessions.case_id=cases.id


    WHERE sessions.session_date BETWEEN ? AND ?



    ORDER BY sessions.session_date


    """,

    conn,

    params=(

    str(today),

    str(week_date)

    )

    )



    conn.close()



    st.subheader(

    "📅 جلسات خلال أسبوع"

    )



    if sessions.empty:


        st.success(

        "لا توجد جلسات قريبة"

        )


    else:


        st.dataframe(

        sessions,

        use_container_width=True,

        hide_index=True

        )



    st.divider()



    st.subheader(

    "⚠️ تنبيهات الطعون"

    )



    st.info(

    """

    يتم حساب مواعيد الطعن تلقائيا:

    - الاستئناف 40 يوم

    - النقض والإدارية العليا والقضاء الإداري 60 يوم

    """

    )




# ============================================================
# المكتبة القانونية
# ============================================================



library_sections=[


"القوانين",

"القرارات الوزارية",

"قرارات الهيئة",

"المنشورات الوزارية",

"منشورات الهيئة",

"الكتب الدورية",

"تعليمات الهيئة",

"رسائل الهيئة",

"المرصد الفني",

"فتاوى لجنة الشئون القانونية بالوزارة",

"فتاوى الإدارة المركزية للشئون القانونية",

"أحكام المحكمة الدستورية العليا",

"أحكام محكمة النقض",

"أحكام المحكمة الإدارية العليا",

"أحكام المحاكم الاستئنافية",

"أحكام محاكم القضاء الإداري",

"أحكام المحاكم الابتدائية",

"أحكام المحكمة الإدارية",

"منشورات القضاء العادي",

"منشورات مجلس الدولة",

"فتاوى الجمعية العمومية",

"صحف طعون",

"صحف استئنافات",

"صحف دعاوى",

"مذكرات دفاع",

"أخرى"

]



def add_library_file():



    st.subheader(

    "📚 إضافة مادة قانونية"

    )


    section=st.selectbox(

    "القسم",

    library_sections

    )



    title=st.text_input(

    "بيان المستند"

    )



    file=st.file_uploader(

    "رفع الملف"

    )



    if st.button(

    "حفظ بالمكتبة"

    ):



        if file:



            path=os.path.join(

            LIB_FOLDER,

            file.name

            )



            with open(path,"wb") as f:

                f.write(

                file.getbuffer()

                )



            conn=db()

            cur=conn.cursor()



            cur.execute(

            """

            INSERT INTO library

            (

            section,

            title,

            file_path,

            created

            )

            VALUES(?,?,?,?)

            """,

            (

            section,

            title,

            path,

            str(datetime.now())

            )

            )



            conn.commit()

            conn.close()



            st.success(

            "تم حفظ المادة القانونية"

            )





def library_page():


    st.markdown(

    """

    <h1 style="text-align:center">

    ⚖️ المكتبة القانونية

    </h1>

    """,

    unsafe_allow_html=True

    )



    add_library_file()



    st.divider()



    search=st.text_input(

    "🔎 البحث في المكتبة"

    )



    conn=db()



    if search:


        df=pd.read_sql_query(

        """

        SELECT

        section AS القسم,

        title AS البيان


        FROM library


        WHERE title LIKE ?


        """,

        conn,

        params=(

        "%"+search+"%",

        )

        )



    else:


        df=pd.read_sql_query(

        """

        SELECT

        section AS القسم,

        title AS البيان


        FROM library


        ORDER BY id DESC


        """,

        conn

        )



    conn.close()



    st.dataframe(

    df,

    use_container_width=True,

    hide_index=True

    )




# ============================================================
# تشغيل الأقسام
# ============================================================


if st.session_state.page=="التنبيهات":

    notifications_page()



if st.session_state.page=="المكتبة القانونية":

    library_page()
    # ============================================================
# إدارة القضايا
# الجزء السادس : التقارير الاحترافية + الإحصائيات
# ============================================================


# ============================================================
# بيانات الدعاوى المتداولة
# ============================================================


def ongoing_cases_report():


    conn=db()


    df=pd.read_sql_query(

    """

    SELECT


    case_number AS م,

    case_year AS السنة_القضائية,

    circle AS الدائرة,

    category AS النوع,

    court_name AS المحكمة,

    department AS المأمورية,

    plaintiff || ' ضد ' || defendant AS الخصوم,

    subject AS موضوع_الدعوى,

    first_session AS آخر_جلسة,

    required_action AS السبب,

    notes AS ملاحظات



    FROM cases


    WHERE status='متداولة'


    ORDER BY first_session DESC



    """,

    conn

    )



    conn.close()


    return df




# ============================================================
# بيانات الأحكام
# ============================================================


def judgments_report(result=None):


    conn=db()



    query="""

    SELECT


    cases.case_number AS م,

    cases.case_year AS السنة_القضائية,

    cases.circle AS الدائرة,

    cases.category AS النوع,

    cases.court_name AS المحكمة,

    cases.department AS المأمورية,

    cases.plaintiff || ' ضد ' ||

    cases.defendant AS الخصوم,

    cases.subject AS موضوع_الدعوى,

    archive.judgment_date AS تاريخ_الحكم,

    archive.judgment_text AS منطوق_الحكم,

    archive.result AS النتيجة



    FROM archive



    JOIN cases


    ON archive.case_id=cases.id



    """



    params=[]



    if result:


        query+=" WHERE archive.result=? "

        params.append(result)



    query+=" ORDER BY archive.judgment_date DESC "



    df=pd.read_sql_query(

    query,

    conn,

    params=params

    )


    conn.close()


    return df




# ============================================================
# تصدير Word
# ============================================================


def export_word(df,title):


    from docx import Document


    file=io.BytesIO()


    doc=Document()



    doc.add_heading(

    "الهيئة القومية للتأمين الاجتماعى",

    level=1

    )


    doc.add_paragraph(

    "الإدارة المركزية للإدارات القانونية\n"

    "الإدارة العامة للقضايا\n"

    "ديوان عام منطقة البحيرة"

    )



    doc.add_heading(

    title,

    level=2

    )



    table=doc.add_table(

    rows=1,

    cols=len(df.columns)

    )


    table.style="Table Grid"



    for i,col in enumerate(df.columns):

        table.rows[0].cells[i].text=str(col)



    for _,row in df.iterrows():

        cells=table.add_row().cells

        for i,v in enumerate(row):

            cells[i].text=str(v)



    doc.save(file)


    file.seek(0)


    return file




# ============================================================
# تصدير PDF
# ============================================================


def export_pdf(df,title):


    from reportlab.platypus import (

    SimpleDocTemplate,

    Table,

    TableStyle,

    Paragraph

    )

    from reportlab.lib.styles import getSampleStyleSheet



    file=io.BytesIO()



    pdf=SimpleDocTemplate(file)



    styles=getSampleStyleSheet()



    content=[]



    content.append(

    Paragraph(

    "الهيئة القومية للتأمين الاجتماعى",

    styles["Title"]

    )

    )



    content.append(

    Paragraph(

    title,

    styles["Heading2"]

    )

    )



    data=[list(df.columns)]



    for _,row in df.iterrows():

        data.append(

        list(map(str,row))

        )



    table=Table(

    data,

    repeatRows=1

    )


    table.setStyle(

    TableStyle(

    [

    ("GRID",

    (0,0),

    (-1,-1),

    1,

    None)

    ]

    )

    )



    content.append(table)



    pdf.build(content)


    file.seek(0)



    return file




# ============================================================
# صفحة التقارير
# ============================================================


def reports_page():


    st.markdown(

    """

    <h1 style="text-align:center">

    📊 التقارير

    </h1>

    """,

    unsafe_allow_html=True

    )



    report=st.selectbox(

    "نوع التقرير",

    [

    "الدعاوى المتداولة",

    "جميع الأحكام",

    "الأحكام للصالح",

    "الأحكام للضد"

    ]

    )



    if report=="الدعاوى المتداولة":


        df=ongoing_cases_report()

        title="بيان بالدعاوى المتداولة"



    elif report=="جميع الأحكام":


        df=judgments_report()

        title="بيان الأحكام"



    elif report=="الأحكام للصالح":


        df=judgments_report("للصالح")

        title="الأحكام الصادرة للصالح"



    else:


        df=judgments_report("للضد")

        title="الأحكام الصادرة للضد"




    st.dataframe(

    df,

    use_container_width=True,

    hide_index=True

    )



    if not df.empty:


        c1,c2=st.columns(2)



        with c1:


            st.download_button(

            "⬇️ تحميل Word",

            export_word(df,title),

            file_name="report.docx"

            )



        with c2:


            st.download_button(

            "⬇️ تحميل PDF",

            export_pdf(df,title),

            file_name="report.pdf"

            )





# ============================================================
# الإحصائيات
# ============================================================


def statistics_page():


    st.markdown(

    """

    <h1 style="text-align:center">

    📈 الإحصائيات

    </h1>

    """,

    unsafe_allow_html=True

    )



    conn=db()



    cases=pd.read_sql_query(

    """

    SELECT COUNT(*) عدد

    FROM cases

    WHERE status='متداولة'

    """,

    conn

    )



    judgments=pd.read_sql_query(

    """

    SELECT COUNT(*) عدد

    FROM archive

    """,

    conn

    )



    good=pd.read_sql_query(

    """

    SELECT COUNT(*) عدد

    FROM archive

    WHERE result='للصالح'

    """,

    conn

    )



    bad=pd.read_sql_query(

    """

    SELECT COUNT(*) عدد

    FROM archive

    WHERE result='للضد'

    """,

    conn

    )



    conn.close()



    c1,c2,c3,c4=st.columns(4)



    c1.metric(

    "القضايا المتداولة",

    cases.iloc[0,0]

    )


    c2.metric(

    "الأحكام",

    judgments.iloc[0,0]

    )


    c3.metric(

    "للصالح",

    good.iloc[0,0]

    )


    c4.metric(

    "للضد",

    bad.iloc[0,0]

    )




# ============================================================
# تشغيل التقارير والإحصائيات
# ============================================================


if st.session_state.page=="التقارير":

    reports_page()



if st.session_state.page=="الإحصائيات":

    statistics_page()
    # ============================================================
# إدارة القضايا
# الجزء السابع : تحسين الواجهة النهائية + النسخ الاحتياطي + الإعدادات
# ============================================================

import shutil


# ============================================================
# تحسين شكل الجداول والخانات
# ============================================================

st.markdown("""
<style>

[data-testid="stDataFrame"] {

    border-radius:20px;
    overflow:hidden;
    border:2px solid #D4AF37;

}


[data-testid="stDataFrame"] table {

    direction:rtl;

}



div.stTextInput > label,
div.stTextArea > label,
div.stSelectbox > label,
div.stDateInput > label {

    color:#FFD700 !important;
    font-weight:bold;
    font-size:17px;

}



.stMetric {

    background:
    linear-gradient(
    135deg,
    #09203f,
    #163f70
    );

    padding:20px;
    border-radius:20px;
    border:1px solid #D4AF37;

}


</style>
""",
unsafe_allow_html=True)



# ============================================================
# النسخ الاحتياطي
# ============================================================


def auto_backup():

    if os.path.exists(DB):

        backup_name = (
            "backup_"
            +
            datetime.now().strftime("%Y%m%d_%H%M%S")
            +
            ".db"
        )

        shutil.copyfile(
            DB,
            backup_name
        )

        return True


    return False




# ============================================================
# صفحة الإعدادات
# ============================================================


def settings_page():


    st.markdown(

    """

    <h1 style="text-align:center;color:#FFD700">

    ⚙️ إعدادات النظام

    </h1>

    """,

    unsafe_allow_html=True

    )


    st.markdown(

    """

    <div class="card">

    <div class="card-title">

    نظام إدارة القضايا يحفظ البيانات نهائياً داخل قاعدة البيانات

    </div>

    </div>

    """,

    unsafe_allow_html=True

    )


    st.write("")



    if st.button(

        "💾 إنشاء نسخة احتياطية",

        use_container_width=True

    ):


        result=auto_backup()


        if result:

            st.success(

            "تم إنشاء النسخة الاحتياطية بنجاح"

            )

        else:

            st.error(

            "تعذر إنشاء النسخة الاحتياطية"

            )





# ============================================================
# صفحة عن البرنامج
# ============================================================


def about_page():


    st.markdown(

    """

    <div style="text-align:center">


    <div style="font-size:80px;color:#FFD700">

    ⚖️

    </div>


    <h1 style="color:#FFD700">

    إدارة القضايا

    </h1>


    <h3 style="color:white">

    نظام إلكتروني متكامل لإدارة ومتابعة القضايا

    </h3>


    <br>


    <p style="color:#FFD700;font-size:20px">


    مع تحيات /

    <br>

    وليد شعبان حماد

    <br><br>

    الإدارة العامة للشئون القانونية

    <br>

    ديوان عام منطقة البحيرة

    <br>

    الهيئة القومية للتأمين الاجتماعى


    </p>


    </div>


    """,

    unsafe_allow_html=True

    )





# ============================================================
# لوحة الإدارة
# ============================================================


with st.sidebar:


    st.markdown(

    """

    <h2 style="color:#FFD700;text-align:center">

    ⚖️ إدارة القضايا

    </h2>

    """,

    unsafe_allow_html=True

    )


    admin_page=st.selectbox(

    "انتقال",

    [

    "الرئيسية",

    "الإحصائيات",

    "الإعدادات",

    "عن البرنامج"

    ]

    )



    if admin_page=="الإحصائيات":

        st.session_state.page="الإحصائيات"



    elif admin_page=="الإعدادات":

        st.session_state.page="الإعدادات"



    elif admin_page=="عن البرنامج":

        st.session_state.page="عن البرنامج"

    # ============================================================
# تشغيل صفحات البرنامج
# ============================================================

if st.session_state.page=="تسجيل القضايا":
    registration_page()

elif st.session_state.page=="الحصر العام":
    general_archive_page()

elif st.session_state.page=="الأرشيف":
    archive_page()

elif st.session_state.page=="البحث عن دعوى":
    search_case_page()



# ============================================================
# نهاية الجزء السابع
# ============================================================
