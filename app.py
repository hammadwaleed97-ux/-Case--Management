# ============================================================
# ================== إدارة القضايا v5.9 =====================
# ========== الإدارة العامة للشئون القانونية البحيرة ==========
# ============================================================

import streamlit as st
import json
import os
from datetime import datetime

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

DATA_FILE = "cases_data.json"
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return {"cases": []}
def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=4)

data = load_data()
today = datetime.now().strftime("%A, %d %B %Y")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;800&display=swap');
    html, body {font-family: 'Cairo', sans-serif; direction: rtl;}
.stApp {background: linear-gradient(180deg, #0F1C2E 0%, #1A2F4F 100%);}
.marquee-container {background: linear-gradient(90deg, #C9A961, #D4B96A); padding: 10px; border-radius: 8px; margin-bottom: 15px;}
.marquee-text {color: #0F1C2E; font-weight: 800; font-size: 14px; white-space: nowrap; display: inline-block; animation: scroll-rtl 18s linear infinite;}
    @keyframes scroll-rtl {0% {transform: translateX(-100%);} 100% {transform: translateX(100%);}}
.header-calm {background: linear-gradient(135deg, #1A2F4F, #2C4A73); padding: 18px; border-radius: 12px; text-align: center; border: 2px solid #C9A961; margin-bottom: 20px;}
.header-calm h1 {color: #D4B96A; font-size: 30px; font-weight: 800; margin: 0;}
.header-calm p {color: #E8E8E8; font-size: 13px; font-weight: 600; margin: 6px 0 0 0;}
.section-divider {height: 4px; background: linear-gradient(90deg, transparent, #C9A961, transparent); margin: 15px 0;}
    label,.stTextInput label,.stSelectbox label,.stTextArea label,.stDateInput label {
        color: #D4B96A!important; font-weight: 700!important; font-size: 15px!important; text-align: right;
    }
    input, textarea, select {color: #0F1C2E!important; font-weight: 600; background-color: #FFFFFF!important;}
    div[data-testid="stButton"] > button {
        background: linear-gradient(135deg, #2C4A73, #3A5F8A)!important; color: #E8E8E8!important; width: 100%; 
        border-radius: 12px; border: 2px solid #C9A961; font-weight: 700; font-size: 15px; height: 90px;
    }
.btn-back button {height: 55px!important;}
.btn-save button {background: linear-gradient(135deg, #C9A961, #D4B96A)!important; color: #0F1C2E!important; height: 50px!important; font-weight:800!important}
.card {background: rgba(26,47,79,0.85); padding: 18px; border-radius: 12px; border: 2px solid #C9A961; margin-bottom: 18px;}
.card-title {color: #D4B96A; font-weight: 800; font-size: 18px; margin-bottom: 15px; border-bottom: 2px solid #C9A961; padding-bottom: 10px; text-align: right;}
.table-container {overflow-x: auto; border: 3px solid #C9A961; border-radius: 10px; background: white;}
.case-table {width: 100%; border-collapse: collapse; background: #FFFFFF; color: #0F1C2E; font-size: 14px;}
.case-table th {background: linear-gradient(90deg, #D4B96A, #C9A961); color: #0F1C2E; padding: 12px 8px; text-align: center; font-weight: 800;}
.case-table td {padding: 10px 8px; text-align: center; border: 1px solid #DDD; font-weight: 600;}
    h2, h3, h4, p {color: #FFFFFF!important;}
</style>
""", unsafe_allow_html=True)

if 'page' not in st.session_state: st.session_state.page = "الرئيسية"

st.markdown("<div class='marquee-container'><div class='marquee-text'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div></div>", unsafe_allow_html=True)
st.markdown(f"<div class='header-calm'><div style='font-size:40px'>⚖️</div><h1>إدارة القضايا</h1><p>📅 {today}</p></div>", unsafe_allow_html=True)
st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)

# ==================== بداية قسم 1: الرئيسية ====================
if st.session_state.page == "الرئيسية":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>الأقسام</h2>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("تسجيل القضايا", use_container_width=True): st.session_state.page = "تسجيل"; st.rerun()
    with col2:
        if st.button("الحصر العام", use_container_width=True): st.session_state.page = "حصر"; st.rerun()
# ==================== نهاية قسم 1: الرئيسية ====================

# ==================== بداية قسم 2: التسجيل كامل ====================
elif st.session_state.page == "تسجيل":
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#C9A961; text-align:center'>تسجيل القضايا</h2>", unsafe_allow_html=True)
    st.markdown("<div class='btn-back'>", unsafe_allow_html=True)
    if st.button("العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"]) # شيلنا نوع المحكمة

    with st.form("form_case"):
        # كارت 1
        st.markdown("<div class='card'><div class='card-title'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        محكمة_اسم = st.text_input("اسم المحكمة")
        if نوع == "استئناف":
            مأمورية = st.text_input("المأمورية")
        else:
            مأمورية = ""
        st.markdown("</div>", unsafe_allow_html=True)
        
        # كارت 2
        st.markdown("<div class='card'><div class='card-title'>2- بيانات الدعوى</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: رقم = st.text_input("رقم الدعوى / الاستئناف / الطعن")
        with col2: سنة = st.text_input("السنة القضائية")
        دائرة = st.text_input("الدائرة")
        st.markdown("</div>", unsafe_allow_html=True)

        # كارت 3
        st.markdown("<div class='card'><div class='card-title'>3- بيانات الخصوم</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: مدعي = st.text_input("اسم المدعى / المستأنف / الطاعن")
        with col2: مدعي_عليه = st.text_input("اسم المدعى عليه / المستأنف ضده / المطعون ضده")
        موضوع = st.text_area("موضوع الدعوى")
        st.markdown("</div>", unsafe_allow_html=True)

        # كارت 4
        st.markdown("<div class='card'><div class='card-title'>4- بيانات الجلسة</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: تاريخ_جلسة = st.date_input("تاريخ أول جلسة", value=datetime.now())
        with col2: الرول = st.text_input("الرول")
        سبب = st.text_input("سبب الجلسة")
        ملاحظات = st.text_area("ملاحظات")
        st.markdown("</div>", unsafe_allow_html=True)

        # كارت 5
        st.markdown("<div class='card'><div class='card-title'>5- التنبيهات والمستندات</div>", unsafe_allow_html=True)
        تنبيه = st.checkbox("تفعيل التنبيهات عبر الواتس اب")
        واتس = st.text_input("رقم هاتف واتس اب") if تنبيه else ""
        col1, col2 = st.columns(2)
        with col1: مستند_نوع = st.selectbox("نوع المستند", ["صحيفة الدعوى", "صحيفة الاستئناف", "صحيفة الطعن"])
        with col2: مستند_ملف = st.file_uploader("اختر الملف")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='btn-save'>", unsafe_allow_html=True)
        if st.form_submit_button("حفظ القضية", use_container_width=True):
            if not رقم or not سنة:
                st.error("❌ من فضلك ادخل رقم الدعوى والسنة")
            else:
                new_case = {
                    "id": len(data["cases"])+1, "نوع": نوع, "محكمة_اسم": محكمة_اسم, "مأمورية": مأمورية,
                    "رقم": رقم, "سنة": سنة, "دائرة": دائرة, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, 
                    "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة), "الرول": الرول, "سبب": سبب, 
                    "ملاحظات": ملاحظات, "تنبيه": تنبيه, "واتس": واتس, "جلسات": []
                }
                data["cases"].append(new_case)
                save_data(data)
                st.success(f"✅ تم حفظ القضية رقم {رقم} لسنة {سنة}")
                st.session_state.page = "الرئيسية"
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
# ==================== نهاية قسم 2: التسجيل ====================
# ==================== بداية قسم 3: الحصر العام ====================
elif st.session_state.page == "حصر":
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📊 الحصر العام 📊</h2>", unsafe_allow_html=True)
    st.markdown("<div class='btn-back'>", unsafe_allow_html=True)
    if st.button("العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    if not data["cases"]:
        st.info("لا توجد قضايا مسجلة")
    else:
        sorted_cases = sorted(data["cases"], key=lambda x: x.get("تاريخ_جلسة","9999"))

        st.markdown("<div class='table-container'>", unsafe_allow_html=True)
        table_html = "<table class='case-table'>"
        table_html += "<tr><th>م</th><th>الرقم</th><th>المحكمة</th><th>الدائرة</th><th>المدعي</th><th>المدعى عليه</th><th>الموضوع</th><th>اخر جلسة</th><th>سببها</th><th>فتح</th></tr>"

        for idx, case in enumerate(sorted_cases, 1):
            رقم_كامل = f"{case.get('رقم','')}<br>لسنة {case.get('سنة','')}"

            # شكل المحكمة: نوع + اسم + مأمورية
            نوع_المحكمة = case.get('نوع','')
            اسم_المحكمة = case.get('محكمة_اسم','')
            مأمورية = case.get('مأمورية','')
            محكمة_كاملة = f"{نوع_المحكمة}<br>{اسم_المحكمة}"
            if مأمورية: محكمة_كاملة += f"<br>مأمورية {مأمورية}"

            دائرة_كاملة = f"{case.get('دائرة','')}"

            اخر_جلسة = case.get('تاريخ_جلسة','')
            سببها = case.get('سبب','')

            # ===== الالوان حسب المدعي عليه =====
            مدعى_عليه = case.get('مدعي_عليه','')
            if "الهيئة" in مدعى_عليه:
                bg = "#FFCDD2" # احمر للهيئة
            elif "مدعي" in مدعى_عليه or "مستأنف" in مدعى_عليه or "طاعن" in مدعى_عليه:
                bg = "#C8E6C9" # اخضر لو الهيئة مدعية
            else:
                bg = "#FFF8E1" if idx % 2 == 1 else "#F0F4F8" # ابيض و رمادي

            table_html += f"<tr style='background:{bg}'>"
            table_html += f"<td>{idx}</td>"
            table_html += f"<td>{رقم_كامل}</td>"
            table_html += f"<td>{محكمة_كاملة}</td>"
            table_html += f"<td>{دائرة_كاملة}</td>"
            table_html += f"<td>{case.get('مدعي','')}</td>"
            table_html += f"<td>{مدعى_عليه}</td>"
            table_html += f"<td>{case.get('موضوع','')}</td>"
            table_html += f"<td>{اخر_جلسة}</td>"
            table_html += f"<td>{سببها}</td>"
            # زر الفتح شغال ب session_state
            table_html += f"<td><button onclick=\"window.parent.postMessage({{type: 'streamlit:setComponentValue', value: {case.get('id')}}}, '*')\" style='background:#C9A961;color:#0F1C2E;border:none;border-radius:5px;padding:8px 20px;font-weight:800;cursor:pointer'>فتح</button></td>"
            table_html += "</tr>"

        table_html += "</table></div>"
        st.markdown(table_html, unsafe_allow_html=True)

        # طريقة فتح جديدة تشتغل مع streamlit
        cols = st.columns([1,10,1])
        with cols[0]:
            for case in sorted_cases:
                if st.button("فتح", key=f"open_{case['id']}"):
                    st.session_state.selected_case_id = case['id']
                    st.session_state.page = "تفاصيل"
                    st.rerun()
# ==================== نهاية قسم 3: الحصر العام ====================

# ==================== بداية قسم 4: التفاصيل ====================
elif st.session_state.page == "تفاصيل":
    case = next((c for c in data["cases"] if c["id"] == st.session_state.selected_case_id), None)
    if case:
        st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='color:#C9A961; text-align:center'>📄 تفاصيل القضية رقم {case['رقم']} لسنة {case['سنة']}</h2>", unsafe_allow_html=True)
        if st.button("العودة للحصر"): st.session_state.page = "حصر"; st.rerun()

        st.markdown("<div class='card'><div class='card-title'>1- بيانات القضية</div>", unsafe_allow_html=True)
        col1,col2 = st.columns(2)
        with col1:
            st.write(f"**نوع الدعوى:** {case['نوع']}")
            st.write(f"**المحكمة:** {case['نوع']} {case['محكمة_اسم']}")
            st.write(f"**المأمورية:** {case['مأمورية']}")
            st.write(f"**الدائرة:** {case['دائرة']}")
        with col2:
            st.write(f"**المدعي:** {case['مدعي']}")
            st.write(f"**المدعى عليه:** {case['مدعي_عليه']}")
            st.write(f"**الموضوع:** {case['موضوع']}")
            st.write(f"**ملاحظات:** {case['ملاحظات']}")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card'><div class='card-title'>2- اخر جلسة</div>", unsafe_allow_html=True)
        st.write(f"**التاريخ:** {case['تاريخ_جلسة']}")
        st.write(f"**الرول:** {case['الرول']}")
        st.write(f"**السبب:** {case['سبب']}")
        st.markdown("</div>", unsafe_allow_html=True)

    else:
        st.error("القضية غير موجودة")
        st.session_state.page = "حصر"
        st.rerun()
# ==================== نهاية قسم 4: التفاصيل ====================
