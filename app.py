# ==========================================================
# إدارة القضايا
# الإصدار الاحترافي
# الجزء الأول (1/12)
# ==========================================================

import streamlit as st
from PIL import Image
import sqlite3
import os

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
# إخفاء واجهة Streamlit
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
padding:0rem;
margin:0rem;
max-width:100%;
}

.stApp{
background:#061a3b;
}

</style>
""", unsafe_allow_html=True)

# ==========================================================
# قاعدة البيانات
# ==========================================================

conn = sqlite3.connect("cases.db", check_same_thread=False)

# ==========================================================
# تحميل صورة الواجهة
# ==========================================================

if os.path.exists("home.png"):

    img = Image.open("home.png")

    st.image(
        img,
        use_container_width=True
    )

else:

    st.error("لم يتم العثور على home.png")

# ==========================================================
# سيتم إضافة الأزرار فوق الصورة فى الجزء القادم
# ==========================================================
