# ===========================================================
# ================== إدارة القضايا v5.9 =====================
# ========== الإدارة العامة للشئون القانونية البحيرة ==========
# ============================================================
import streamlit as st
import pandas as pd # <--- ده اللي كان ناقص
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
.info-box {background:#1E2A47; border:1px solid #C9A961; border-radius:8px; padding:15px; margin-bottom:10px; color:#FFFFFF}
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

    # لو دوست فتح
    if "open_case" in st.session_state:
        st.session_state.selected_case_id = st.session_state.open_case
        st.session_state.page = "تفاصيل"
        del st.session_state.open_case
        st.rerun()

    if not data["cases"]:
        st.info("لا توجد قضايا مسجلة")
    else:
        # ترتيب من الاقدم للاحدث
        sorted_cases = sorted(data["cases"], key=lambda x: x.get("تاريخ_جلسة","9999"))

        st.markdown("<div class='table-container'>", unsafe_allow_html=True)
        # الجدول يبدأ بـ م وشلنا الحالة
        table_html = "<table class='case-table'>"
        table_html += "<tr><th>م</th><th>الرقم</th><th>المحكمة</th><th>الدائرة</th><th>المدعي</th><th>المدعى عليه</th><th>الموضوع</th><th>اخر جلسة</th><th>سببها</th><th>فتح</th></tr>"

        for idx, case in enumerate(sorted_cases, 1):
            رقم_كامل = f"{case.get('رقم','')}<br>لسنة {case.get('سنة','')}"

            محكمة_كاملة = f"{case.get('نوع','')}<br>{case.get('محكمة_اسم','')}"
            if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"

            دائرة_كاملة = f"{case.get('دائرة','')} مدنى"

            مدعى = case.get('مدعي','')
            مدعى_عليه = case.get('مدعي_عليه','')

            # التلوين
            if "الهيئة" in مدعى: # الهيئة مدعية = احمر
                bg = "#FFCDD2"
            else: # الهيئة مدعى عليها = اصفر ورمادي
                bg = "#FFF8E1" if idx % 2 == 1 else "#F0F4F8"

            # زر الفتح بلينك يغير session_state
            table_html += f"<tr style='background:{bg}'>"
            table_html += f"<td>{idx}</td>"
            table_html += f"<td>{رقم_كامل}</td>"
            table_html += f"<td>{محكمة_كاملة}</td>"
            table_html += f"<td>{دائرة_كاملة}</td>"
            table_html += f"<td>{مدعى}</td>"
            table_html += f"<td>{مدعى_عليه}</td>"
            table_html += f"<td>{case.get('موضوع','')}</td>"
            table_html += f"<td>{case.get('تاريخ_جلسة','')}</td>"
            table_html += f"<td>{case.get('سبب','')}</td>"
            table_html += f"<td><a href='?open={case.get('id')}' target='_self' style='background:#C9A961;color:#0F1C2E;text-decoration:none;border-radius:5px;padding:6px 15px;font-weight:800'>فتح</a></td>"
            table_html += "</tr>"

        table_html += "</table></div>"
        st.markdown(table_html, unsafe_allow_html=True)

        # نمسك اللينك
        query_params = st.query_params
        if "open" in query_params:
            st.session_state.open_case = int(query_params["open"])
            st.query_params.clear()
            st.rerun()
# ==================== نهاية قسم 3: الحصر العام ====================
# ==================== بداية قسم 4: تفاصيل القضية ====================

        st.markdown(f"<div class='info-box' style='background:linear-gradient(135deg, #F0F4F8 0%, #E1E8F0 100%)'><b>المدعى عليه:</b><br>{case.get('مدعي_عليه')}</div>", unsafe_allow_html=True)

    # كارت 3: ادارة الجلسات
    st.markdown("<h3 style='color:#C9A961'>📅 إدارة الجلسات</h3>", unsafe_allow_html=True)
    
    with st.expander("➕ اضافة جلسة جديدة"):
        with st.form("new_session"):
            c1, c2 = st.columns(2)
            تاريخ = c1.date_input("تاريخ الجلسة")
            سبب = c2.text_input("سبب الجلسة")
            قرار = st.text_area("قرار الجلسة")
            if st.form_submit_button("حفظ الجلسة"):
                if "جلسات" not in case: case["جلسات"] = []
                case["جلسات"].append({"تاريخ": str(تاريخ), "سبب": سبب, "قرار": قرار})
                case["تاريخ_جلسة"] = str(تاريخ)
                case["سبب"] = سبب
                save_data(data)
                st.success("تم حفظ الجلسة"); st.rerun()

    if "جلسات" in case and case["جلسات"]:
        st.markdown("<h4>سجل الجلسات</h4>", unsafe_allow_html=True)
        جلسات_مرتبة = sorted(case["جلسات"], key=lambda x: x['تاريخ'])
        table_html = "<table class='case-table'>"
        table_html += "<tr><th>م</th><th>التاريخ</th><th>السبب</th><th>القرار</th></tr>"
        for i, ج in enumerate(جلسات_مرتبة, 1):
            table_html += f"<tr><td>{i}</td><td>{ج['تاريخ']}</td><td>{ج['سبب']}</td><td>{ج['قرار']}</td></tr>"
        table_html += "</table>"
        st.markdown(table_html, unsafe_allow_html=True)
    else:
        st.info("لا توجد جلسات مسجلة لهذه القضية")

    # كارت 4: المذكرات والمستندات
    st.markdown("<h3 style='color:#C9A961'>📎 المذكرات والمستندات</h3>", unsafe_allow_html=True)
    st.info("سيتم اضافة رفع الملفات هنا في التحديث القادم")

    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    if st.button("🗑️ حذف القضية نهائيا", type="primary"):
        data["cases"] = [c for c in data["cases"] if c['id'] != case['id']]
        save_data(data)
        st.success("تم حذف القضية"); st.session_state.page = "حصر"; st.rerun()
        # ==================== بداية قسم 4: تفاصيل القضية ====================
elif st.session_state.page == "تفاصيل":
    case = next((c for c in data["cases"] if c['id'] == st.session_state.selected_case_id), None)
    if not case:
        st.error("القضية غير موجودة")
        st.session_state.page = "حصر"; st.rerun()

    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown(f"<h2 style='color:#FFFFFF; text-align:center'>📄 تفاصيل القضية رقم {case.get('رقم')} لسنة {case.get('سنة')}</h2>", unsafe_allow_html=True)
    st.markdown("<div class='btn-back'>", unsafe_allow_html=True)
    if st.button("العودة للحصر", use_container_width=True): st.session_state.page = "حصر"; st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    # جدول بيانات القضية - مززهه
    st.markdown("<h3 style='color:#C9A961'>📌 بيانات القضية</h3>", unsafe_allow_html=True)
    
    table_html = """
    <div class='table-container' style='background:linear-gradient(180deg, #0F1A2E 0%, #1E2A47 100%); padding:15px; border-radius:15px; border:2px solid #C9A961; box-shadow:0 0 20px rgba(201,169,97,0.3)'>
    <table style='width:100%; table-layout:fixed; border-spacing:8px; border-collapse:separate'>
    <!-- الصف الاول 4 مربعات -->
    <tr>
        <td style='background:linear-gradient(135deg, #1A365D 0%, #2C5282 100%); border:2px solid #C9A961; border-radius:12px; padding:15px; text-align:center; height:90px; box-shadow:0 4px 8px rgba(0,0,0,0.3)'>
            <div style='font-size:11px; color:#C9A961; margin-bottom:5px; font-weight:bold; letter-spacing:1px'>رقم القضية</div>
            <div style='font-size:20px; color:#FFFFFF; font-weight:bold'>""" + str(case.get('رقم')) + """</div>
        </td>
        <td style='background:linear-gradient(135deg, #2C5282 0%, #1A365D 100%); border:2px solid #C9A961; border-radius:12px; padding:15px; text-align:center; height:90px; box-shadow:0 4px 8px rgba(0,0,0,0.3)'>
            <div style='font-size:11px; color:#C9A961; margin-bottom:5px; font-weight:bold; letter-spacing:1px'>السنة</div>
            <div style='font-size:20px; color:#FFFFFF; font-weight:bold'>""" + str(case.get('سنة')) + """</div>
        </td>
        <td style='background:linear-gradient(135deg, #1A365D 0%, #2C5282 100%); border:2px solid #C9A961; border-radius:12px; padding:15px; text-align:center; height:90px; box-shadow:0 4px 8px rgba(0,0,0,0.3)'>
            <div style='font-size:11px; color:#C9A961; margin-bottom:5px; font-weight:bold; letter-spacing:1px'>الدائرة</div>
            <div style='font-size:18px; color:#FFFFFF; font-weight:bold'>""" + str(case.get('دائرة')) + """ مدنى</div>
        </td>
        <td style='background:linear-gradient(135deg, #2C5282 0%, #1A365D 100%); border:2px solid #C9A961; border-radius:12px; padding:15px; text-align:center; height:90px; box-shadow:0 4px 8px rgba(0,0,0,0.3)'>
            <div style='font-size:11px; color:#C9A961; margin-bottom:5px; font-weight:bold; letter-spacing:1px'>النوع</div>
            <div style='font-size:16px; color:#FFFFFF; font-weight:bold'>""" + str(case.get('نوع')) + """</div>
        </td>
    </tr>
    <!-- الصف الثاني -->
    <tr>
        <td colspan='2' style='background:linear-gradient(135deg, #1A365D 0%, #2C5282 100%); border:2px solid #C9A961; border-radius:12px; padding:15px; text-align:center; height:90px; box-shadow:0 4px 8px rgba(0,0,0,0.3)'>
            <div style='font-size:11px; color:#C9A961; margin-bottom:5px; font-weight:bold; letter-spacing:1px'>المحكمة</div>
            <div style='font-size:16px; color:#FFFFFF; font-weight:bold'>""" + str(case.get('محكمة_اسم')) + """</div>
        </td>
        <td colspan='2' style='background:linear-gradient(135deg, #2C5282 0%, #1A365D 100%); border:2px solid #C9A961; border-radius:12px; padding:15px; text-align:center; height:90px; box-shadow:0 4px 8px rgba(0,0,0,0.3)'>
            <div style='font-size:11px; color:#C9A961; margin-bottom:5px; font-weight:bold; letter-spacing:1px'>المأمورية</div>
            <div style='font-size:16px; color:#FFFFFF; font-weight:bold'>""" + str(case.get('مأمورية') or '-') + """</div>
        </td>
    </tr>
    <!-- الصف الثالث -->
    <tr>
        <td colspan='4' style='background:linear-gradient(135deg, #1A365D 0%, #2C5282 100%); border:2px solid #C9A961; border-radius:12px; padding:15px; text-align:center; height:90px; box-shadow:0 4px 8px rgba(0,0,0,0.3)'>
            <div style='font-size:11px; color:#C9A961; margin-bottom:5px; font-weight:bold; letter-spacing:1px'>الموضوع</div>
            <div style='font-size:16px; color:#FFFFFF; font-weight:bold'>""" + str(case.get('موضوع')) + """</div>
        </td>
    </tr>
    </table>
    </div>
    """
    st.markdown(table_html, unsafe_allow_html=True)
    
    # كارت 2: الخصوم - واضحين ومززهين
    st.markdown("<h3 style='color:#C9A961'>👥 الخصوم</h3>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div style='background:linear-gradient(135deg, #FFF8E1 0%, #FFECB3 100%); border:2px solid #C9A961; border-radius:12px; padding:15px; box-shadow:0 4px 8px rgba(0,0,0,0.2)'>
            <div style='font-size:12px; color:#8B6914; font-weight:bold; margin-bottom:5px'>المدعي:</div>
            <div style='font-size:15px; color:#1A365D; font-weight:bold'>{case.get('مدعي')}</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div style='background:linear-gradient(135deg, #E3F2FD 0%, #BBDEFB 100%); border:2px solid #C9A961; border-radius:12px; padding:15px; box-shadow:0 4px 8px rgba(0,0,0,0.2)'>
            <div style='font-size:12px; color:#1565C0; font-weight:bold; margin-bottom:5px'>المدعى عليه:</div>
            <div style='font-size:15px; color:#1A365D; font-weight:bold'>{case.get('مدعي_عليه')}</div>
        </div>
        """, unsafe_allow_html=True)

    # كارت 3: ادارة الجلسات
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
                case["تاريخ_جلسة"] = str(تاريخ)
                case["سبب"] = سبب
                save_data(data)
                st.success("تم حفظ الجلسة"); st.rerun()

    if "جلسات" in case and case["جلسات"]:
        st.markdown("<h4 style='color:#C9A961'>سجل الجلسات</h4>", unsafe_allow_html=True)
        جلسات_مرتبة = sorted(case["جلسات"], key=lambda x: x['تاريخ'])
        table_html = "<table class='case-table' style='border:2px solid #C9A961'>"
        table_html += "<tr style='background:linear-gradient(90deg, #1A365D, #2C5282)'><th>م</th><th>التاريخ</th><th>السبب</th><th>القرار</th></tr>"
        for i, ج in enumerate(جلسات_مرتبة, 1):
            table_html += f"<tr><td>{i}</td><td>{ج['تاريخ']}</td><td>{ج['سبب']}</td><td>{ج['قرار']}</td></tr>"
        table_html += "</table>"
        st.markdown(table_html, unsafe_allow_html=True)
    else:
        st.info("لا توجد جلسات مسجلة لهذه القضية")

    # كارت 4: المذكرات والمستندات
    st.markdown("<h3 style='color:#C9A961'>📎 المذكرات والمستندات</h3>", unsafe_allow_html=True)
    st.info("سيتم اضافة رفع الملفات هنا في التحديث القادم")

    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    if st.button("🗑️ حذف القضية نهائيا", type="primary"):
        data["cases"] = [c for c in data["cases"] if c['id'] != case['id']]
        save_data(data)
        st.success("تم حذف القضية"); st.session_state.page = "حصر"; st.rerun()
# ==================== نهاية قسم 4: تفاصيل القضية ====================
