# ===========================================================
# ================== إدارة القضايا v5.13 =====================
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

    نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])

    with st.form("form_case"):
        st.markdown("<div class='card'><div class='card-title'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        محكمة_اسم = st.text_input("اسم المحكمة")
        if نوع == "استئناف":
            مأمورية = st.text_input("المأمورية")
        else:
            مأمورية = ""
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

# ==================== القسم 3: الحصر العام الخارجي ====================
elif st.session_state.page == "حصر":
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📊 الحصر العام الخارجي</h2>", unsafe_allow_html=True)

    if st.button("العودة للرئيسية", use_container_width=True):
        st.session_state.page = "الرئيسية"
        st.rerun()

    if not data["cases"]:
        st.info("لا توجد قضايا مسجلة")
    else:
        updated = False
        for i, case in enumerate(data["cases"]):
            if "id" not in case:
                case["id"] = i + 1
                updated = True
        if updated: save_data(data)

        sorted_cases = sorted(data["cases"], key=lambda x: x.get("تاريخ_جلسة","9999"))

        for idx, case in enumerate(sorted_cases, 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
            if case.get('مأمورية',''):
                محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = f"{case.get('دائرة','')} عمال" if case.get('دائرة','') else ""
            محكمة_كاملة += f"<br>{دائرة_كاملة}"

            خصوم = f"{case.get('مدعي','')}<br>ضد<br>{case.get('مدعي_عليه','')}"

            if "الهيئة" in str(case.get('مدعي','')):
                row_class = "row-hey2a"
            else:
                row_class = "row1" if idx % 2 == 1 else "row2"

            # جدول القضية
            st.markdown("<div class='table-container'>", unsafe_allow_html=True)
            table_html = "<table class='case-table'>"
            table_html += "<tr><th>م</th><th>الرقم والسنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>اخر جلسة</th><th>السبب</th></tr>"
            table_html += f"<tr class='{row_class}'>"
            table_html += f"<td>{idx}</td>"
            table_html += f"<td>{رقم_كامل}</td>"
            table_html += f"<td>{محكمة_كاملة}</td>"
            table_html += f"<td>{خصوم}</td>"
            table_html += f"<td>{case.get('موضوع','')}</td>"
            table_html += f"<td>{case.get('تاريخ_جلسة','')}</td>"
            table_html += f"<td>{case.get('سبب','')}</td>"
            table_html += "</tr></table></div>"
            st.markdown(table_html, unsafe_allow_html=True)
            
            # الزرار الصغير الاحمر الداكن في النص تحت كل قضية
            c1, c2, c3 = st.columns([4,1,4])
            with c2:
                if st.button("فتح", key=f"open_{case['id']}"):
                    st.session_state.selected_case_id = case['id']
                    st.session_state.page = "تفاصيل"
                    st.rerun()

# ==================== نهاية القسم 3: الحصر العام الخارجي ====================

# ==================== بداية قسم 4: تفاصيل القضية ====================
elif st.session_state.page == "تفاصيل":
    case = next((c for c in data["cases"] if c['id'] == st.session_state.selected_case_id), None)
    if not case:
        st.error("القضية غير موجودة")
        st.session_state.page = "حصر"; st.rerun()

    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown(f"<h2 style='color:#FFFFFF; text-align:center'>📄 تفاصيل القضية رقم {case.get('رقم')} لسنة {case.get('سنة')}</h2>", unsafe_allow_html=True)
    
    if st.button("العودة للحصر", use_container_width=True): st.session_state.page = "حصر"; st.rerun()

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
    
    st.markdown("<h3 style='color:#C9A961'>📅 إدارة الجلسات 📆</h3>", unsafe_allow_html=True)
    with st.expander("➕ اضافة جلسة جديدة"):
        with st.form("new_session"):
            c1, c2 = st.columns(2)
            تاريخ = c1.date_input("تاريخ الجلسة")
            سبب = c2.text_input("سبب الجلسة")
            قرار = st.text_area("قرار الجلسة")
            if st.form_submit_button("حفظ الجلسة"):
                if "جلسات" not in case: case["جلسات"] = []
                case["جلسات"].append({"تاريخ": str(تاريخ), "سبب": سبب, "قرار": قرار})
                case["تاريخ_جلسة"] = str(تاريخ); case["سبب"] = سبب
                save_data(data); st.success("تم حفظ الجلسة"); st.rerun()

    if "جلسات" in case and case["جلسات"]:
        st.markdown("<h4 style='color:#C9A961'>سجل الجلسات</h4>", unsafe_allow_html=True)
        جلسات_مرتبة = sorted(case["جلسات"], key=lambda x: x['تاريخ'])
        table_html = "<table style='width:100%; border:2px solid #C9A961; border-radius:10px'><tr style='background:linear-gradient(90deg, #1E3A6B, #2C5282)'><th>م</th><th>التاريخ</th><th>السبب</th><th>القرار</th></tr>"
        for i, ج in enumerate(جلسات_مرتبة, 1):
            table_html += f"<tr><td>{i}</td><td>{ج['تاريخ']}</td><td>{ج['سبب']}</td><td>{ج['قرار']}</td></tr>"
        table_html += "</table>"
        st.markdown(table_html, unsafe_allow_html=True)
    else:
        st.info("لا توجد جلسات مسجلة لهذه القضية")

    st.markdown("<h3 style='color:#C9A961'>📎 المذكرات والمستندات</h3>", unsafe_allow_html=True)
    st.info("سيتم اضافة رفع الملفات هنا في التحديث القادم")

    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    if st.button("🗑️ حذف القضية نهائيا", type="primary"):
        data["cases"] = [c for c in data["cases"] if c['id'] != case['id']]
        save_data(data); st.success("تم حذف القضية"); st.session_state.page = "حصر"; st.rerun()
# ==================== نهاية قسم 4: تفاصيل القضية ====================
# ==================== بداية قسم 4: تفاصيل القضية v5.18 ====================
elif st.session_state.page == "تفاصيل":
    case = next((c for c in data["cases"] if c['id'] == st.session_state.selected_case_id), None)
    if not case:
        st.error("القضية غير موجودة")
        st.session_state.page = "حصر"; st.rerun()

    col_main, col_docs = st.columns([19,1]) # 95% بيانات و 5% مستندات

    with col_docs:
        st.markdown("<div style='background:#0F1C2E; border-right:3px solid #C9A961; padding:10px; height:100vh; overflow-y:auto'><div style='color:#C9A961; font-weight:800; text-align:center; margin-bottom:10px; border-bottom:2px solid #C9A961; padding-bottom:5px'>مستندات<br>القضية</div>", unsafe_allow_html=True)
        if case.get('مستندات'):
            for م in case['مستندات']:
                st.markdown(f"<div style='font-size:11px; color:#FFF; margin-bottom:8px; text-align:center'>{م['النوع']}<br><small>{م['البيان']}</small></div>", unsafe_allow_html=True)
        else:
            st.markdown("<div style='font-size:10px; color:#AAA; text-align:center'>لا توجد</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with col_main:
        st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='color:#FFFFFF; text-align:center'>📄 تفاصيل القضية رقم {case.get('رقم')} لسنة {case.get('سنة')}</h2>", unsafe_allow_html=True)
        
        if st.button("العودة للحصر", use_container_width=True): st.session_state.page = "حصر"; st.rerun()

        st.markdown("<h3 style='color:#C9A961'>📌 بيانات القضية</h3>", unsafe_allow_html=True)
        # ... سيب بيانات القضية بتاعتك زي ما هي فوق هنا ...

        # ================= 1. متابعة الجلسات =================
        st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
        st.markdown("<h3 style='color:#C9A961'>📅 متابعة الجلسات</h3>", unsafe_allow_html=True)
        if case.get("جلسات"):
            جلسات_مرتبة = sorted(case["جلسات"], key=lambda x: x['تاريخ'])
            table_html = """
            <table style='width:100%; border:3px solid #C9A961; border-radius:12px; background:#0A1428; border-collapse: collapse; margin-bottom:20px; box-shadow: 0 4px 12px rgba(201,169,97,0.3)'>
            <tr style='background:linear-gradient(90deg, #C9A961, #D4B96A); color:#0F1C2E; font-weight:900; font-size:16px'>
                <th style='padding:12px; border:2px solid #C9A961; text-align:center'>م</th>
                <th style='padding:12px; border:2px solid #C9A961; text-align:center'>الرول</th>
                <th style='padding:12px; border:2px solid #C9A961; text-align:center'>الجلسات</th>
                <th style='padding:12px; border:2px solid #C9A961; text-align:center'>الإجراءات</th>
                <th style='padding:12px; border:2px solid #C9A961; text-align:center'>ملاحظات</th>
            </tr>"""
            for i, ج in enumerate(جلسات_مرتبة, 1):
                لون_الصف = "#1E2A47" if i % 2 == 0 else "#0F1C2E" # تلطيش
                table_html += f"""
                <tr style='background:{لون_الصف}'>
                    <td style='text-align:center; padding:12px; border:2px solid #C9A961; color:#FFFFFF; font-weight:700; font-size:15px'>{i}</td>
                    <td style='text-align:center; padding:12px; border:2px solid #C9A961; color:#FFD700; font-weight:800; font-size:15px'>{ج.get('الرول','')}</td>
                    <td style='text-align:center; padding:12px; border:2px solid #C9A961; color:#FFFFFF; font-weight:700; font-size:15px'>{ج['تاريخ']}</td>
                    <td style='text-align:center; padding:12px; border:2px solid #C9A961; color:#FFFFFF; font-weight:700; font-size:15px'>{ج.get('سبب','')}</td>
                    <td style='text-align:center; padding:12px; border:2px solid #C9A961; color:#FFFFFF; font-weight:700; font-size:15px'>{ج.get('ملاحظات','')}</td>
                </tr>"""
            table_html += "</table>"
            st.markdown(table_html, unsafe_allow_html=True)
        else:
            st.info("لا توجد جلسات مسجلة")

        # ================= 2. اضافة جلسة =================
        with st.expander("➕ اضافة جلسة جديدة"):
            with st.form("add_session"):
                c1, c2 = st.columns(2)
                تاريخ_جديد = c1.date_input("تاريخ الجلسة")
                رول_جديد = c2.text_input("الرول")
                سبب_جديد = st.text_input("سبب التأجيل")
                ملاحظات_جديد = st.text_area("ملاحظات")
                if st.form_submit_button("حفظ الجلسة"):
                    جلسه_جديده = {
                        "تاريخ": str(تاريخ_جديد), 
                        "الرول": رول_جديد, 
                        "سبب": سبب_جديد,
                        "ملاحظات": ملاحظات_جديد
                    }
                    if "جلسات" not in case: case["جلسات"] = []
                    case["جلسات"].append(جلسه_جديده)
                    # تحديث اخر جلسة في الحصر العام
                    case["تاريخ_جلسة"] = str(تاريخ_جديد); case["سبب"] = سبب_جديد
                    save_data(data); st.success("✅ تم حفظ الجلسة"); st.rerun()

        # ================= 3. تعديل جلسة =================
        with st.expander("✏️ تعديل جلسة سابقة"):
            if case.get("جلسات"):
                جلسات_مرتبة = sorted(case["جلسات"], key=lambda x: x['تاريخ'])
                تواريخ = [ج['تاريخ'] for ج in جلسات_مرتبة]
                تاريخ_للتعديل = st.selectbox("اختر تاريخ الجلسة للتعديل", تواريخ)
                جلسه_مختارة = next((ج for ج in جلسات_مرتبة if ج['تاريخ'] == تاريخ_للتعديل), None)
                
                if جلسه_مختارة:
                    with st.form("edit_session"):
                        c1, c2 = st.columns(2)
                        تاريخ_معدل = c1.date_input("التاريخ الجديد", value=datetime.strptime(جلسه_مختارة['تاريخ'], "%Y-%m-%d"))
                        رول_معدل = c2.text_input("الرول الجديد", value=جلسه_مختارة.get('الرول',''))
                        سبب_معدل = st.text_input("سبب التأجيل الجديد", value=جلسه_مختارة.get('سبب',''))
                        ملاحظات_معدل = st.text_area("ملاحظات جديدة", value=جلسه_مختارة.get('ملاحظات',''))
                        if st.form_submit_button("حفظ التعديل"):
                            جلسه_مختارة['تاريخ'] = str(تاريخ_معدل)
                            جلسه_مختارة['الرول'] = رول_معدل
                            جلسه_مختارة['سبب'] = سبب_معدل
                            جلسه_مختارة['ملاحظات'] = ملاحظات_معدل
                            # لو دي اخر جلسة حدثها فوق
                            if str(تاريخ_معدل) >= case["تاريخ_جلسة"]:
                                case["تاريخ_جلسة"] = str(تاريخ_معدل); case["سبب"] = سبب_معدل
                            save_data(data); st.success("✅ تم تعديل الجلسة"); st.rerun()
            else:
                st.warning("لا توجد جلسات للتعديل")

        # ================= 4. مستندات القضية =================
        st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
        st.markdown("<h3 style='color:#C9A961'>📎 مستندات القضية</h3>", unsafe_allow_html=True)
        with st.form("upload_doc"):
            نوع_مستند = st.selectbox("نوع المستند", [
                "صحيفة دعوى", "صحيفة استئناف", "صحيفة طعن", "مذكرة دفاع", "حافظة مستندات", 
                "تقرير خبير", "تقرير طب شرعي", "تقرير لجنة طبية", "صحيفة تجديد من الشطب", 
                "صحيفة تعجيل من الوقف", "صورة حكم تمهيدي", "اخرى"])
            بيان_مستند = st.text_input("بيان المستند")
            ملف_مستند = st.file_uploader("اختر الملف PDF او Word")
            c1, c2 = st.columns(2)
            with c1:
                if st.form_submit_button("حفظ المستند"):
                    if ملف_مستند:
                        if "مستندات" not in case: case["مستندات"] = []
                        case["مستندات"].append({"النوع": نوع_مستند, "البيان": بيان_مستند, "الاسم": ملف_مستند.name})
                        save_data(data); st.success("✅ تم حفظ المستند"); st.rerun()
            with c2:
                if st.form_submit_button("الغاء حفظ المستند"): st.rerun()

        # ================= 5. جلسة الحكم =================
        st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
        st.markdown("<h3 style='color:#C9A961'>⚖️ بيانات جلسة الحكم</h3>", unsafe_allow_html=True)
        with st.form("judgement_form"):
            c1, c2 = st.columns(2)
            تاريخ_الحكم = c1.date_input("تاريخ جلسة الحكم")
            مسند = c2.selectbox("مسند الحكم", ["", "لصالح الهيئة", "ضد الهيئة"])
            منطوق_الحكم = st.text_area("منطوق الحكم")
            ملاحظات_الحكم = st.text_area("ملاحظات على الحكم")
            if st.form_submit_button("حفظ الحكم"):
                case["تاريخ_الحكم"] = str(تاريخ_الحكم)
                case["منطوق_الحكم"] = منطوق_الحكم
                case["مسند_الحكم"] = مسند
                # اضافة الحكم لجدول الاجراءات
                if "جلسات" not in case: case["جلسات"] = []
                case["جلسات"].append({
                    "تاريخ": str(تاريخ_الحكم), 
                    "الرول": "الحكم", 
                    "سبب": منطوق_الحكم,
                    "ملاحظات": f"{مسند}"
                })
                case["تاريخ_جلسة"] = str(تاريخ_الحكم); case["سبب"] = "صدر حكم"
                save_data(data); st.success("✅ تم حفظ الحكم"); st.rerun()

        st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
        if st.button("🗑️ حذف القضية نهائيا", type="primary"):
            data["cases"] = [c for c in data["cases"] if c['id'] != case['id']]
            save_data(data); st.success("تم حذف القضية"); st.session_state.page = "حصر"; st.rerun()
# ==================== نهاية قسم 4: تفاصيل القضية ====================
