# ===========================================================
# ================== إدارة القضايا v5.22 =====================
# ========== الإدارة العامة للشئون القانونية البحيرة ==========
# ============================================================
import streamlit as st
import pandas as pd
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
.table-container {overflow-x: auto; border: 3px solid #C9A961; border-radius: 10px; background: white; margin-bottom:15px}
.case-table {width: 100%; border-collapse: collapse; background: #FFFFFF; color: #0F1C2E; font-size: 14px;}
.case-table th {background: linear-gradient(90deg, #D4B96A, #C9A961); color: #0F1C2E; padding: 12px 8px; text-align: center; font-weight: 800;}
.case-table td {padding: 10px 8px; text-align: center; border: 1px solid #DDD; font-weight: 600;}
.row1 {background:#F9F9F9}
.row2 {background:#FFFFFF}
.row-hey2a {background:#FFE5E5; font-weight:800}
.btn-open-small {background: linear-gradient(135deg, #8B0000, #A52A2A)!important; color: #FFFFFF!important; border:none!important; border-radius:6px!important; padding:4px 15px!important; font-size:12px!important; font-weight:700!important; height:28px!important; width:75px!important;}
.info-box {background:#1E2A47; border:1px solid #C9A961; border-radius:8px; padding:15px; margin-bottom:10px; color:#FFFFFF}
    h2, h3, h4, p {color: #FFFFFF!important;}
</style>
""", unsafe_allow_html=True)

if 'page' not in st.session_state: st.session_state.page = "الرئيسية"
if 'selected_case_id' not in st.session_state: st.session_state.selected_case_id = None

st.markdown("<div class='marquee-container'><div class='marquee-text'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div></div>", unsafe_allow_html=True)
st.markdown(f"<div class='header-calm'><div style='font-size:40px'>⚖️</div><h1>إدارة القضايا</h1><p>📅 {today}</p></div>", unsafe_allow_html=True)
st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)

if st.session_state.page == "الرئيسية":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>الأقسام</h2>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("تسجيل القضايا", use_container_width=True): st.session_state.page = "تسجيل"; st.rerun()
    with col2:
        if st.button("الحصر العام", use_container_width=True): st.session_state.page = "حصر"; st.rerun()

elif st.session_state.page == "تسجيل":
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#C9A961; text-align:center'>تسجيل القضايا</h2>", unsafe_allow_html=True)
    st.markdown("<div class='btn-back'>", unsafe_allow_html=True)
    if st.button("العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)
    نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
    with st.form("form_case"):
        st.markdown("<div class='card'><div class='card-title'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        محكمة_اسم = st.text_input("اسم المحكمة")
        مأمورية = st.text_input("المأمورية") if نوع == "استئناف" else ""
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<div class='card'><div class='card-title'>2- بيانات الدعوى</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: رقم = st.text_input("رقم الدعوى / الاستئناف / الطعن")
        with col2: سنة = st.text_input("السنة القضائية")
        دائرة = st.text_input("الدائرة")
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<div class='card'><div class='card-title'>3- بيانات الخصوم</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: مدعي = st.text_input("اسم المدعى / المستأنف / الطاعن")
        with col2: مدعي_عليه = st.text_input("اسم المدعى عليه / المستأنف ضده / المطعون ضده")
        موضوع = st.text_area("موضوع الدعوى")
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<div class='card'><div class='card-title'>4- بيانات الجلسة</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: تاريخ_جلسة = st.date_input("تاريخ أول جلسة", value=datetime.now())
        with col2: الرول = st.text_input("الرول")
        سبب = st.text_input("سبب الجلسة")
        ملاحظات = st.text_area("ملاحظات")
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<div class='card'><div class='card-title'>5- التنبيهات والمستندات</div>", unsafe_allow_html=True)
        تنبيه = st.checkbox("تفعيل التنبيهات عبر الواتس اب")
        واتس = st.text_input("رقم هاتف واتس اب") if تنبيه else ""
        col1, col2 = st.columns(2)
        with col1: مستند_نوع = st.selectbox("نوع المستند", ["صحيفة الدعوى", "صحيفة الاستئناف", "صحيفة الطعن"])
        with col2: مستند_ملف = st.file_uploader("اختر الملف")
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<div class='btn-save'>", unsafe_allow_html=True)
        if st.form_submit_button("حفظ القضية", use_container_width=True):
            if not رقم or not سنة: st.error("❌ من فضلك ادخل رقم الدعوى والسنة")
            else:
                new_case = {"id": len(data["cases"])+1, "نوع": نوع, "محكمة_اسم": محكمة_اسم, "مأمورية": مأمورية, "رقم": رقم, "سنة": سنة, "دائرة": دائرة, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة), "الرول": الرول, "سبب": سبب, "ملاحظات": ملاحظات, "تنبيه": تنبيه, "واتس": واتس, "جلسات": []}
                data["cases"].append(new_case); save_data(data)
                st.success(f"✅ تم حفظ القضية رقم {رقم} لسنة {سنة}")
                st.session_state.page = "الرئيسية"; st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

elif st.session_state.page == "حصر":
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📊 الحصر العام الخارجي</h2>", unsafe_allow_html=True)
    if st.button("العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()
    if not data["cases"]: st.info("لا توجد قضايا مسجلة")
    else:
        for i, case in enumerate(data["cases"]):
            if "id" not in case: case["id"] = i + 1
        save_data(data)
        sorted_cases = sorted(data["cases"], key=lambda x: x.get("تاريخ_جلسة","9999"))
        for idx, case in enumerate(sorted_cases, 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
            if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = f"{case.get('دائرة','')} عمال" if case.get('دائرة','') else ""
            محكمة_كاملة += f"<br>{دائرة_كاملة}"
            خصوم = f"{case.get('مدعي','')}<br>ضد<br>{case.get('مدعي_عليه','')}"
            row_class = "row-hey2a" if "الهيئة" in str(case.get('مدعي','')) else "row1" if idx % 2 == 1 else "row2"
            st.markdown("<div class='table-container'>", unsafe_allow_html=True)
            table_html = f"<table class='case-table'><tr><th>م</th><th>الرقم والسنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>اخر جلسة</th><th>السبب</th></tr><tr class='{row_class}'><td>{idx}</td><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td>{case.get('تاريخ_جلسة','')}</td><td>{case.get('سبب','')}</td></tr></table></div>"
            st.markdown(table_html, unsafe_allow_html=True)
            c1, c2, c3 = st.columns([4,1,4])
            with c2:
                if st.button("فتح", key=f"open_{case['id']}"): st.session_state.selected_case_id = case['id']; st.session_state.page = "تفاصيل"; st.rerun()

elif st.session_state.page == "تفاصيل":
    case = next((c for c in data["cases"] if c['id'] == st.session_state.selected_case_id), None)
    if not case: st.error("القضية غير موجودة"); st.session_state.page = "حصر"; st.rerun()
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown(f"<h2 style='color:#FFFFFF; text-align:center'>📄 تفاصيل القضية رقم {case.get('رقم')} لسنة {case.get('سنة')}</h2>", unsafe_allow_html=True)
    if st.button("العودة للحصر", use_container_width=True): st.session_state.page = "حصر"; st.rerun()
    
    # ================= 1. بيانات القضية كاملة =================
    st.markdown("<h3 style='color:#C9A961'>📌 بيانات القضية</h3>", unsafe_allow_html=True)
    table_html = f"""
    <div style='background:linear-gradient(180deg, #0A1428 0%, #1E2A47 100%); padding:15px; border-radius:18px; border:2px solid #C9A961'>
    <table style='width:100%; border-spacing:8px 8px; border-collapse:separate'>
    <tr>
        <td style='background:linear-gradient(145deg, #1E3A6B 0%, #2C5282 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:right'>
            <div style='font-size:12px; color:#FFD700; margin-bottom:6px; font-weight:bold'>رقم القضية</div>
            <div style='font-size:20px; color:#FFFFFF; font-weight:bold'>{case.get('رقم')}</div>
        </td>
        <td style='background:linear-gradient(145deg, #2C5282 0%, #1E3A6B 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:right'>
            <div style='font-size:12px; color:#FFD700; margin-bottom:6px; font-weight:bold'>السنة</div>
            <div style='font-size:20px; color:#FFFFFF; font-weight:bold'>{case.get('سنة')}</div>
        </td>
        <td style='background:linear-gradient(145deg, #1E3A6B 0%, #2C5282 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:right'>
            <div style='font-size:12px; color:#FFD700; margin-bottom:6px; font-weight:bold'>الدائرة</div>
            <div style='font-size:18px; color:#FFFFFF; font-weight:bold'>{case.get('دائرة')} عمال</div>
        </td>
        <td style='background:linear-gradient(145deg, #2C5282 0%, #1E3A6B 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:right'>
            <div style='font-size:12px; color:#FFD700; margin-bottom:6px; font-weight:bold'>النوع</div>
            <div style='font-size:16px; color:#FFFFFF; font-weight:bold'>{case.get('نوع')}</div>
        </td>
    </tr>
    <tr>
        <td colspan='2' style='background:linear-gradient(145deg, #1E3A6B 0%, #2C5282 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:right'>
            <div style='font-size:12px; color:#FFD700; margin-bottom:6px; font-weight:bold'>المحكمة</div>
            <div style='font-size:15px; color:#FFFFFF; font-weight:bold'>{case.get('محكمة_اسم')}</div>
        </td>
        <td colspan='2' style='background:linear-gradient(145deg, #2C5282 0%, #1E3A6B 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:right'>
            <div style='font-size:12px; color:#FFD700; margin-bottom:6px; font-weight:bold'>المأمورية</div>
            <div style='font-size:15px; color:#FFFFFF; font-weight:bold'>{case.get('مأمورية') or '-'}</div>
        </td>
    </tr>
    <tr>
        <td colspan='2' style='background:linear-gradient(145deg, #FFF3CD 0%, #FFE69C 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:right'>
            <div style='font-size:12px; color:#8B6914; margin-bottom:6px; font-weight:bold'>المدعي</div>
            <div style='font-size:15px; color:#1E3A6B; font-weight:bold'>{case.get('مدعي')}</div>
        </td>
        <td colspan='2' style='background:linear-gradient(145deg, #CFF4FC 0%, #9EEAF9 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:right'>
            <div style='font-size:12px; color:#055160; margin-bottom:6px; font-weight:bold'>المدعى عليه</div>
            <div style='font-size:15px; color:#1E3A6B; font-weight:bold'>{case.get('مدعي_عليه')}</div>
        </td>
    </tr>
    <tr>
        <td colspan='4' style='background:linear-gradient(145deg, #1E3A6B 0%, #2C5282 100%); border:2px solid #C9A961; border-radius:12px; padding:12px; text-align:center'>
            <div style='font-size:12px; color:#FFD700; margin-bottom:6px; font-weight:bold'>الموضوع</div>
            <div style='font-size:15px; color:#FFFFFF; font-weight:bold'>{case.get('موضوع')}</div>
        </td>
    </tr>
    </table></div>
    """
    st.markdown(table_html, unsafe_allow_html=True)
    
    # ================= 2. متابعة الجلسات =================
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h3 style='color:#C9A961'>📅 متابعة الجلسات</h3>", unsafe_allow_html=True)
    if 'جلسات' not in case: case['جلسات'] = []
    if len(case['جلسات']) > 0:
        جلسات_مرتبة = sorted(case["جلسات"], key=lambda x: x['تاريخ'])
        html = "<table style='width:100%; border:3px solid #C9A961; background:#0A1428; border-radius:12px; border-collapse: collapse'><tr style='background:#C9A961; color:#0F1C2E'><th style='padding:12px; border:2px solid #C9A961'>م</th><th style='padding:12px; border:2px solid #C9A961'>الرول</th><th style='padding:12px; border:2px solid #C9A961'>الجلسات</th><th style='padding:12px; border:2px solid #C9A961'>الإجراءات</th><th style='padding:12px; border:2px solid #C9A961'>ملاحظات</th></tr>"
        for i, ج in enumerate(جلسات_مرتبة, 1):
            لون = "#1E2A47" if i % 2 == 0 else "#142038"
            html += f"<tr style='background:{لون}; color:#FFF'><td style='text-align:center; padding:12px; border:2px solid #C9A961'>{i}</td><td style='text-align:center; padding:12px; border:2px solid #C9A961; color:#FFD700; font-weight:800'>{ج.get('الرول','-')}</td><td style='text-align:center; padding:12px; border:2px solid #C9A961'>{ج['تاريخ']}</td><td style='text-align:center; padding:12px; border:2px solid #C9A961'>{ج.get('سبب','-')}</td><td style='text-align:center; padding:12px; border:2px solid #C9A961'>{ج.get('ملاحظات','-')}</td></tr>"
        html += "</table>"
        st.markdown(html, unsafe_allow_html=True)
    else: st.info("لا توجد جلسات مسجلة")
    with st.expander("➕ اضافة جلسة"):
        with st.form("add_session"):
            c1,c2 = st.columns(2)
            t = c1.date_input("تاريخ الجلسة")
            r = c2.text_input("الرول")
            s = st.text_input("سبب التأجيل")
            m = st.text_area("ملاحظات")
            if st.form_submit_button("حفظ الجلسة"):
                case['جلسات'].append({'تاريخ':str(t),'الرول':r,'سبب':s,'ملاحظات':m})
                case['تاريخ_جلسة']=str(t); case['سبب']=s
                save_data(data); st.success("✅ تم حفظ الجلسة"); st.rerun()

    # ================= 3. المستندات =================
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h3 style='color:#C9A961'>📎 المذكرات والمستندات</h3>", unsafe_allow_html=True)
    st.info("سيتم اضافة رفع الملفات هنا في التحديث القادم")

    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    if st.button("🗑️ حذف القضية نهائيا", type="primary"):
        data["cases"] = [c for c in data["cases"] if c['id'] != case['id']]
        save_data(data); st.success("تم حذف القضية"); st.session_state.page = "حصر"; st.rerun()
