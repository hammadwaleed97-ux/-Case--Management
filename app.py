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
        # ==================================================
# دوال التنبيهات (تم إصلاحها)
# ==================================================

def get_alert_cases():

    data = load_data()

    if not isinstance(data, dict):
        data = {
            "cases": [],
            "library": []
        }

    today = datetime.now().date()

    all_cases = data.get("cases", [])

    alerts = {
        "sessions": [],
        "appeals": []
    }

    for case in all_cases:

        # ==================================
        # تنبيهات الجلسات خلال 7 أيام
        # ==================================
        if (
            case.get("حالة") == "متداولة"
            and case.get("تاريخ_جلسة")
        ):

            try:

                session_date = datetime.strptime(
                    case["تاريخ_جلسة"],
                    "%Y-%m-%d"
                ).date()

                days_left = (
                    session_date - today
                ).days

                if 0 <= days_left <= 7:

                    case_copy = case.copy()
                    case_copy["days_left"] = days_left

                    alerts["sessions"].append(
                        case_copy
                    )

            except Exception:
                pass

        # ==================================
        # تنبيهات الطعون
        # ==================================
        if (
            case.get("حالة") == "منتهية"
            and case.get("مسندة_ل_الحكم") == "الضد"
            and case.get("تاريخ_الحكم")
        ):

            try:

                judgment_date = datetime.strptime(
                    case["تاريخ_الحكم"],
                    "%Y-%m-%d"
                ).date()

                appeal_days = (
                    40
                    if case.get("نوع") == "دعوى"
                    else 60
                )

                last_appeal_day = (
                    judgment_date +
                    timedelta(days=appeal_days)
                )

                notify_start = (
                    last_appeal_day -
                    timedelta(days=15)
                )

                days_left_appeal = (
                    last_appeal_day - today
                ).days

                if (
                    notify_start <= today <= last_appeal_day
                    and days_left_appeal >= 0
                ):

                    case_copy = case.copy()

                    case_copy[
                        "days_left_appeal"
                    ] = days_left_appeal

                    alerts["appeals"].append(
                        case_copy
                    )

            except Exception:
                pass

    return alerts


# ==================================================
# الصفحة الرئيسية
# ==================================================

if st.session_state.page == "الرئيسية":

    st.markdown(
        "<h2>الأقسام</h2>",
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            '<div class="btn-add">',
            unsafe_allow_html=True
        )

        if st.button(
            "➕ تسجيل القضايا",
            use_container_width=True
        ):
            st.session_state.page = "تسجيل"
            st.rerun()

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            '<div class="btn-list">',
            unsafe_allow_html=True
        )

        if st.button(
            "📋 الحصر العام",
            use_container_width=True
        ):
            st.session_state.page = "الحصر"
            st.rerun()

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:

        st.markdown(
            '<div class="btn-alert">',
            unsafe_allow_html=True
        )

        if st.button(
            "🔴 مركز التنبيهات",
            use_container_width=True
        ):
            st.session_state.page = "تنبيهات"
            st.rerun()

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            '<div class="btn-report">',
            unsafe_allow_html=True
        )

        if st.button(
            "📊 التقارير",
            use_container_width=True
        ):
            st.session_state.page = "تقارير"
            st.rerun()

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            '<div class="btn-lib">',
            unsafe_allow_html=True
        )

        if st.button(
            "📚 المكتبة القانونية",
            use_container_width=True
        ):
            st.session_state.page = "مكتبة"
            st.rerun()

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:

        st.markdown(
            '<div class="btn-arch">',
            unsafe_allow_html=True
        )

        if st.button(
            "🗄️ الأرشيف",
            use_container_width=True
        ):
            st.session_state.page = "الأرشيف"
            st.rerun()

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:

        st.markdown(
            '<div class="btn-search">',
            unsafe_allow_html=True
        )

        if st.button(
            "🔍 البحث عن دعوى",
            use_container_width=True
        ):
            st.session_state.page = "بحث"
            st.rerun()

        st.markdown(
            "</div>",
            unsafe_allow_html=True
                )
        # ==========================================
========= الجزء الثاني: تسجيل القضايا ============
# ================================================
elif st.session_state.page == "تسجيل":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#D4AF37; text-align:center'>➕ تسجيل القضايا</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", key="back_add", use_container_width=True):
        st.session_state.page = "الرئيسية"
        st.rerun()

    نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"], key="case_type_add")
    with st.form("form_case_add"):
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        محكمة_اسم = st.text_input("اسم المحكمة", key="court_name_add"); مأمورية = st.text_input("المأمورية", key="mamoria_add") if نوع == "استئناف" else ""
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>2- بيانات الدعوى</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: رقم = st.text_input("رقم الدعوى / الاستئناف / الطعن", key="case_num_add")
        with col2: سنة = st.text_input("السنة القضائية", key="case_year_add")
        دائرة = st.text_input("الدائرة", key="circle_add"); st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>3- بيانات الخصوم</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: مدعي = st.text_input("اسم المدعى / المستأنف / الطاعن", key="plaintiff_add")
        with col2: مدعي_عليه = st.text_input("اسم المدعى عليه / المستأنف ضده / المطعون ضده", key="defendant_add")
        موضوع = st.text_area("موضوع الدعوى", height=100, key="subject_add"); st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>4- بيانات الجلسة</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: تاريخ_جلسة = st.date_input("تاريخ أول جلسة", value=datetime.now(), key="session_date_add")
        with col2: الرول = st.text_input("الرول", key="roll_add")
        سبب = st.text_input("سبب الجلسة", key="reason_add"); ملاحظات = st.text_area("ملاحظات", height=100, key="notes_add"); st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>5- صحيفة الدعوى</div>", unsafe_allow_html=True)
        مسندة_ل = st.selectbox("نوع الصحيفة", ["صحيفة الدعوى"], key="paper_type_add"); st.markdown("</div>", unsafe_allow_html=True)

        if st.form_submit_button("💾 حفظ القضية", use_container_width=True, type="primary"):
            if not رقم or not سنة: st.error("❌ من فضلك ادخل رقم الدعوى والسنة")
            else:
                new_case = {"id": len(data["cases"])+1, "نوع": نوع, "محكمة_اسم": محكمة_اسم, "مأمورية": مأمورية, "رقم": رقم, "سنة": سنة, "دائرة": دائرة, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة), "الرول": الرول, "سبب": سبب, "ملاحظات": ملاحظات, "جلسات": [], "مستندات": [], "حالة": "متداولة", "مسندة_ل": مسندة_ل}
                if الرول or سبب: new_case["جلسات"].append({"تاريخ":str(تاريخ_جلسة),"الرول":الرول,"سبب":سبب,"ملاحظات":ملاحظات})
                data["cases"].append(new_case); save_data(data)
                st.success(f"✅ تم حفظ القضية رقم {رقم} لسنة {سنة}")
                st.session_state.page = "الحصر"; st.rerun()
# ===============================
