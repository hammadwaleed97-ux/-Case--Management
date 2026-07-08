# ==========================================================
# إدارة القضايا
# Professional Judicial UI
# الجزء الأول (1/4)
# ==========================================================

import streamlit as st
import sqlite3
import os

st.set_page_config(
    page_title="إدارة القضايا",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# -------------------------------
# إخفاء واجهة Streamlit
# -------------------------------

st.markdown("""
<style>

#MainMenu{visibility:hidden;}
header{visibility:hidden;}
footer{visibility:hidden;}

.block-container{
padding-top:0rem;
padding-bottom:0rem;
padding-left:0rem;
padding-right:0rem;
max-width:100%;
}

.stApp{
background:
linear-gradient(rgba(2,18,45,.93),rgba(1,12,32,.95)),
url("https://images.unsplash.com/photo-1589829545856-d10d557cf95f?auto=format&fit=crop&w=1600&q=80");
background-size:cover;
background-attachment:fixed;
}

/*======================================*/

.title1{
text-align:center;
font-size:44px;
font-weight:900;
color:#FFD700;
margin-top:15px;
text-shadow:0 0 18px gold;
}

.title2{
text-align:center;
font-size:23px;
font-weight:bold;
color:white;
margin-top:-10px;
}

.title3{
text-align:center;
font-size:20px;
color:#d6d6d6;
margin-top:-10px;
}

.author{
text-align:center;
font-size:20px;
color:#FFD700;
font-weight:bold;
margin-top:-8px;
margin-bottom:25px;
}

/*======================================*/

.books{

background:rgba(255,255,255,.05);

backdrop-filter:blur(12px);

border-radius:20px;

padding:15px;

border:1px solid rgba(255,215,0,.3);

box-shadow:0 0 25px rgba(255,215,0,.25);

}

/*======================================*/

.menu{

background:linear-gradient(90deg,#0b1f3a,#123f77);

padding:18px;

margin-bottom:15px;

border-radius:16px;

color:white;

font-size:22px;

font-weight:bold;

transition:.4s;

border-left:6px solid gold;

cursor:pointer;

}

.menu:hover{

transform:translateX(10px);

background:linear-gradient(90deg,#1565c0,#0d47a1);

box-shadow:0 0 25px gold;

}

/*======================================*/

.centerbox{

background:rgba(255,255,255,.04);

border-radius:25px;

padding:35px;

height:650px;

border:1px solid rgba(255,215,0,.2);

backdrop-filter:blur(8px);

box-shadow:0 0 40px rgba(0,0,0,.4);

}

/*======================================*/

.bigicon{

font-size:180px;

text-align:center;

color:gold;

filter:drop-shadow(0 0 20px gold);

margin-top:20px;

}

.desc{

text-align:center;

font-size:26px;

color:white;

line-height:55px;

}

</style>
""",unsafe_allow_html=True)

st.markdown("""
<div class='title1'>
⚖ إدارة القضايا
</div>

<div class='title2'>
الهيئة القومية للتأمين الاجتماعى
</div>

<div class='title3'>
الإدارة العامة للقضايا
</div>

<div class='author'>
إعداد / وليد حماد
</div>
""",unsafe_allow_html=True)

left,center,right=st.columns([1.2,3,1.2])
