# =============================================
# ============ الجزء الاول: الاساسيات ============
# ================================================
import streamlit as st
import pandas as pd
import json
import os
import smtplib
import secrets
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

st.set_page_config(
    page_title="إدارة القضايا",
    layout="wide",
    page_icon="⚖️"
)

# ============= التصميم النهائي =============
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');

*{
    font-family:'Cairo',sans-serif!important;
}

html,body{
    direction:rtl;
    color:#FFF!important;
}

.stApp{
    background:linear-gradient(180deg,#0A1428 0%,#1E2A47 100%);
}

.marquee{
    background:linear-gradient(90deg,#D4AF37,#FFD700,#D4AF37);
    color:#0A1428;
    padding:12px;
    font-weight:900;
    font-size:16px;
    white-space:nowrap;
    overflow:hidden;
    border-radius:0 0 15px 15px;
}

.marquee span{
    display:inline-block;
    animation:marquee 15s linear infinite;
}

@keyframes marquee{
0%{transform:translateX(-100%);}
100%{transform:translateX(100%);}
}

.main-title{
    color:#D4AF37;
    text-align:center;
    font-size:36px;
    font-weight:900;
    padding:15px 0;
}

h1,h2,h3{
    color:#D4AF37!important;
    text-align:center!important;
}

div[data-testid="column"]{
    display:flex;
    justify-content:center;
}

[data-testid="stForm"] label,
.stMarkdown{
    color:#FFF!important;
    font-weight:700;
}

.stButton>button{
    width:100%!important;
    max-width:400px!important;
    border:none!important;
    border-radius:15px!important;
    font-size:18px!important;
    font-weight:900!important;
    padding:16px!important;
    color:#000!important;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="marquee">
<span>
مع تحيات وليد حماد - الإدارة العامة للشئون القانونية
بديوان عام منطقة البحيرة بالهيئة القومية للتأمين الاجتماعي
</span>
</div>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="main-title">⚖️ إدارة القضايا ⚖️</div>',
    unsafe_allow_html=True
)

# ==================================================
# المتغيرات العامة
# ==================================================

DATA_FILE = "cases_data.json"
UPLOAD_FOLDER = "uploads"
TOKENS_FILE = "tokens.json"

ANWA3_MOSTANDAT = [
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

SENDER_EMAIL=""
SENDER_PASSWORD=""
APP_URL="https://qpyqpsmkqcvdou4imbfunp.streamlit.app/"

os.makedirs(UPLOAD_FOLDER,exist_ok=True)

if "page" not in st.session_state:
    st.session_state.page="الرئيسية"

if "selected_case_id" not in st.session_state:
    st.session_state.selected_case_id=None


# ==================================================
# دوال التحميل والحفظ (تم إصلاحها)
# ==================================================

def load_data():

    if not os.path.exists(DATA_FILE):
        return {
            "cases":[],
            "library":[]
        }

    try:
        with open(DATA_FILE,"r",encoding="utf-8") as f:
            data=json.load(f)

        if not isinstance(data,dict):
            data={}

        data.setdefault("cases",[])
        data.setdefault("library",[])

        return data

    except Exception:
        return {
            "cases":[],
            "library":[]
        }


def save_data(data):

    data.setdefault("cases",[])
    data.setdefault("library",[])

    with open(DATA_FILE,"w",encoding="utf-8") as f:
        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=4
        )


def load_tokens():

    if os.path.exists(TOKENS_FILE):
        try:
            with open(TOKENS_FILE,"r",encoding="utf-8") as f:
                return json.load(f)
        except:
            pass

    return {"tokens":[]}


def save_tokens(tokens_data):

    with open(TOKENS_FILE,"w",encoding="utf-8") as f:
        json.dump(
            tokens_data,
            f,
            ensure_ascii=False,
            indent=4
)
