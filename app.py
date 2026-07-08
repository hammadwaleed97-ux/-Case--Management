# ==========================================================
# إدارة القضايا
# Professional Judicial Edition
# الجزء الأول (1/4)
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
# قاعدة البيانات
# ==========================================================

conn = sqlite3.connect("cases.db", check_same_thread=False)
cur = conn.cursor()

# ==========================================================
# تنسيق البرنامج
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

padding-left:1rem;

padding-right:1rem;

max-width:100%;

}

.stApp{

background:
linear-gradient(rgba(2,17,45,.94),rgba(1,10,28,.96)),
url("https://images.unsplash.com/photo-1589829545856-d10d557cf95f?auto=format&fit=crop&w=1600&q=80");

background-size:cover;

background-position:center;

background-attachment:fixed;

}

/*=============================*/

.title{

text-align:center;

font-size:52px;

font-weight:900;

color:#FFD700;

margin-top:15px;

text-shadow:0 0 20px gold;

}

/*=============================*/

.subtitle{

text-align:center;

font-size:25px;

font-weight:bold;

color:white;

margin-top:10px;

}

/*=============================*/

.name{

text-align:center;

font-size:34px;

font-weight:900;

color:#FFD700;

margin-top:8px;

text-shadow:0 0 15px gold;

}

/*=============================*/

.end{

text-align:center;

font-size:23px;

font-weight:bold;

color:white;

margin-top:10px;

line-height:45px;

}

/*=============================*/

.menu{

background:linear-gradient(90deg,#09244d,#0d47a1);

padding:18px;

border-radius:15px;

margin-bottom:15px;

font-size:22px;

font-weight:bold;

color:white;

border-left:6px solid gold;

transition:.4s;

cursor:pointer;

box-shadow:0 0 15px rgba(0,0,0,.3);

}

.menu:hover{

background:linear-gradient(90deg,#1565c0,#1976d2);

transform:translateX(10px);

box-shadow:0 0 20px gold;

}

/*=============================*/

.center{

background:rgba(255,255,255,.05);

border-radius:25px;

padding:35px;

height:700px;

border:1px solid rgba(255,215,0,.25);

backdrop-filter:blur(10px);

box-shadow:0 0 40px rgba(0,0,0,.5);

}

/*=============================*/

.icon{

font-size:190px;

text-align:center;

margin-top:20px;

filter:drop-shadow(0 0 20px gold);

}

/*=============================*/

.txt{

font-size:28px;

color:white;

text-align:center;

line-height:55px;

margin-top:30px;

}

</style>

""", unsafe_allow_html=True)

# ==========================================================
# الهيدر
# ==========================================================

st.markdown("""

<div class="title">

⚖️ إدارة القضايا

</div>

<div class="subtitle">

إعداد وتصميم

</div>

<div class="name">

وليد شعبان حماد

</div>

<div class="subtitle">

مع تحيات

</div>

<div class="end">

الإدارة العامة للشئون القانونية<br>

بديوان عام منطقة البحيرة<br>

بالهيئة القومية للتأمين الاجتماعى

</div>

<br>

""", unsafe_allow_html=True)

# ==========================================================
# تقسيم الصفحة
# ==========================================================

left, center, right = st.columns([1.1,3,1.1])
# ==========================================================
# القائمة الجانبية + منتصف الشاشة
# الجزء الثاني (2/4)
# ==========================================================

with left:

    st.markdown("""
    <div class="menu">📚 تسجيل القضايا</div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="menu">📑 الحصر العام</div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="menu">🔍 البحث عن دعوى</div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="menu">📊 التقارير</div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="menu">🔔 التنبيهات</div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="menu">🗂 أرشيف القضايا</div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="menu">⚖️ المكتبة القانونية</div>
    """, unsafe_allow_html=True)


# ==========================================================
# منتصف الصفحة
# ==========================================================

with center:

    st.markdown("""

    <div class="center">

        <div style="display:flex;
                    justify-content:space-between;
                    align-items:center;">

            <div style="width:25%;
                        text-align:center;">

                <div style="font-size:170px;
                            color:#FFD700;
                            filter:drop-shadow(0 0 18px gold);">

                    ⚖️

                </div>

                <div style="
                color:#FFD700;
                font-size:24px;
                font-weight:bold;">

                ميزان العدالة

                </div>

            </div>

            <div style="width:50%;
                        text-align:center;">

                <div style="
                font-size:70px;
                color:#FFD700;
                font-weight:900;
                text-shadow:0 0 25px gold;">

                إدارة القضايا

                </div>

                <br>

                <div style="
                font-size:28px;
                color:white;
                line-height:60px;">

                نظام إلكترونى متكامل لإدارة الدعاوى القضائية
                <br>
                ومتابعة الجلسات والتقارير
                <br>
                وإدارة المستندات
                <br>
                والمكتبة القانونية

                </div>

            </div>

            <div style="width:25%;
                        text-align:center;">

                <div style="font-size:170px;
                            color:#FFD700;
                            filter:drop-shadow(0 0 18px gold);">

                    🏛️

                </div>

                <div style="
                color:#FFD700;
                font-size:24px;
                font-weight:bold;">

                العدالة

                </div>

            </div>

        </div>

    </div>

    """, unsafe_allow_html=True)
