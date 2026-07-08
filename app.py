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
# ==========================================================
# الجزء الثاني (2/10)
# القائمة الجانبية + الواجهة الرئيسية
# ==========================================================

with left:

    st.markdown("""
    <style>

    .menu_btn{

        background:linear-gradient(90deg,#0b2d5a,#1359b8);

        color:white;

        padding:16px;

        margin-bottom:15px;

        border-radius:15px;

        text-align:center;

        font-size:22px;

        font-weight:bold;

        border-left:6px solid gold;

        box-shadow:0 0 15px rgba(0,0,0,.4);

        transition:.3s;

    }

    .menu_btn:hover{

        background:linear-gradient(90deg,#1565c0,#1e88e5);

        transform:translateX(8px);

        box-shadow:0 0 20px gold;

    }

    </style>
    """,unsafe_allow_html=True)

    st.markdown('<div class="menu_btn">📚 تسجيل القضايا</div>',unsafe_allow_html=True)

    st.markdown('<div class="menu_btn">📑 الحصر العام</div>',unsafe_allow_html=True)

    st.markdown('<div class="menu_btn">🔍 البحث عن دعوى</div>',unsafe_allow_html=True)

    st.markdown('<div class="menu_btn">📊 التقارير</div>',unsafe_allow_html=True)

    st.markdown('<div class="menu_btn">🔔 التنبيهات</div>',unsafe_allow_html=True)

    st.markdown('<div class="menu_btn">🗂️ أرشيف القضايا</div>',unsafe_allow_html=True)

    st.markdown('<div class="menu_btn">⚖️ المكتبة القانونية</div>',unsafe_allow_html=True)


# ==========================================================
# منتصف الشاشة
# ==========================================================

with center:

    st.markdown("""
    <div class="box">

    <div style="display:flex;
                justify-content:space-between;
                align-items:center;">

        <div style="width:22%;text-align:center;">

            <div style="
            font-size:170px;
            color:#FFD700;
            text-shadow:0 0 25px gold;">
            ⚖️
            </div>

            <div style="
            color:#FFD700;
            font-size:24px;
            font-weight:bold;">
            ميزان العدالة
            </div>

        </div>

        <div style="width:56%;
                    text-align:center;">

            <div style="
            font-size:70px;
            font-weight:900;
            color:#FFD700;
            text-shadow:0 0 20px gold;">

            إدارة القضايا

            </div>

            <br>

            <div style="
            color:white;
            font-size:28px;
            line-height:60px;">

            نظام إلكترونى متكامل

            <br>

            لإدارة الدعاوى القضائية

            <br>

            ومتابعة الجلسات

            <br>

            والتقارير والمستندات

            <br>

            والمكتبة القانونية

            </div>

        </div>

        <div style="width:22%;text-align:center;">

            <div style="
            font-size:170px;
            color:#FFD700;
            text-shadow:0 0 25px gold;">
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
    # ==========================================================
# الجزء الثالث (3/10)
# صور الواجهة الرئيسية
# ==========================================================

with right:

    st.markdown("""
    <style>

    .card{

        background:rgba(255,255,255,.05);

        border-radius:20px;

        padding:18px;

        border:1px solid rgba(255,215,0,.35);

        text-align:center;

        margin-bottom:20px;

        box-shadow:0 0 25px rgba(0,0,0,.35);

    }

    .cardtitle{

        color:#FFD700;

        font-size:22px;

        font-weight:bold;

        margin-top:10px;

    }

    .cardtext{

        color:white;

        font-size:17px;

        line-height:35px;

    }

    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="card">

    <img src="https://cdn-icons-png.flaticon.com/512/3132/3132693.png"
         width="130">

    <div class="cardtitle">

    العدالة

    </div>

    <div class="cardtext">

    تطبيق إلكترونى لإدارة
    <br>
    جميع الدعاوى القضائية

    </div>

    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="card">

    <img src="https://cdn-icons-png.flaticon.com/512/3062/3062634.png"
         width="130">

    <div class="cardtitle">

    الجلسات

    </div>

    <div class="cardtext">

    متابعة الجلسات
    <br>
    والتنبيهات

    </div>

    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="card">

    <img src="https://cdn-icons-png.flaticon.com/512/2991/2991112.png"
         width="130">

    <div class="cardtitle">

    المكتبة القانونية

    </div>

    <div class="cardtext">

    القوانين
    <br>
    والأحكام
    <br>
    واللوائح

    </div>

    </div>
    """, unsafe_allow_html=True)


# ==========================================================
# شريط سفلى
# ==========================================================

st.markdown("""

<hr style="
border:1px solid gold;
margin-top:25px;">

<div style="

text-align:center;

color:white;

font-size:22px;

padding:15px;

background:linear-gradient(90deg,#061a3b,#0d47a1,#061a3b);

border-top:2px solid gold;

">

⚖️ جميع الحقوق محفوظة ©

<span style="color:#FFD700;
font-weight:bold;">

وليد شعبان حماد

</span>

| الإدارة العامة للشئون القانونية |

ديوان عام منطقة البحيرة

</div>

""", unsafe_allow_html=True)
