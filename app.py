# ==========================================================
# إدارة القضايا
# Professional Judicial Edition
# الجزء الأول (1/10)
# ==========================================================

import streamlit as st
import sqlite3
import pandas as pd
import os
from datetime import datetime

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
# إنشاء قاعدة البيانات
# ==========================================================

conn = sqlite3.connect("cases.db", check_same_thread=False)
cur = conn.cursor()

cur.execute("""
CREATE TABLE IF NOT EXISTS cases(
id INTEGER PRIMARY KEY AUTOINCREMENT,
case_number TEXT,
judicial_year TEXT,
court TEXT,
subject TEXT
)
""")

conn.commit()

# ==========================================================
# CSS
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
padding-bottom:0rem;
padding-left:0rem;
padding-right:0rem;
max-width:100%;
}

.stApp{

background:linear-gradient(
135deg,
#02162f 0%,
#052b57 40%,
#031c3d 100%
);

}

/*================ HEADER =================*/

.main_title{

text-align:center;

font-size:55px;

font-weight:900;

color:#FFD700;

margin-top:20px;

text-shadow:0px 0px 20px gold;

}

.sub{

text-align:center;

font-size:25px;

font-weight:bold;

color:white;

margin-top:10px;

}

.name{

text-align:center;

font-size:34px;

font-weight:900;

color:#FFD700;

text-shadow:0 0 15px gold;

margin-top:8px;

}

.footertext{

text-align:center;

font-size:23px;

font-weight:bold;

line-height:45px;

color:white;

margin-top:12px;

}

/*================ BOX =================*/

.box{

background:rgba(255,255,255,.05);

border-radius:25px;

padding:25px;

border:1px solid rgba(255,215,0,.30);

backdrop-filter:blur(10px);

box-shadow:0 0 35px rgba(0,0,0,.45);

height:720px;

}

</style>
""", unsafe_allow_html=True)

# ==========================================================
# الهيدر
# ==========================================================

st.markdown("""

<div class="main_title">

⚖️ إدارة القضايا

</div>

<div class="sub">

إعداد وتصميم

</div>

<div class="name">

وليد شعبان حماد

</div>

<div class="sub">

مع تحيات

</div>

<div class="footertext">

الإدارة العامة للشئون القانونية<br>

بديوان عام منطقة البحيرة<br>

بالهيئة القومية للتأمين الاجتماعى

</div>

<br><br>

""", unsafe_allow_html=True)

# ==========================================================
# تقسيم الصفحة
# ==========================================================

left,center,right=st.columns([1.2,3.5,1.2])
