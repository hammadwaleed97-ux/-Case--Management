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
