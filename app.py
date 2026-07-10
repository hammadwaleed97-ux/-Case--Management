# ============================================================
# ================== إدارة القضايا v3.9 =====================
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

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;800&display=swap');
    html, body {{font-family: 'Cairo', sans-serif; direction: rtl;}}
 .stApp {{background: linear-gradient(180deg, #0F1C2E 0%, #1A2F4F 100%);}}
 .stApp::before {{
        content: "⚖️"; position: fixed; top: 50%; left: 50%;
        transform: translate(-50%, -50%); font-size: 400px;
        opacity: 0.04; z-index: 0; color: #C9A961;
    }}
 .marquee-container {{
        background: linear-gradient(90deg, #C9A961, #D4B96A);
        padding: 10px; overflow: hidden; border-radius: 8px;
        margin-bottom: 15px; box-shadow: 0 0 10px rgba(201,169,97,0.4);
    }}
 .marquee-text {{
        color: #0F1C2E; font-weight: 800; font-size: 14px;
        white-space: nowrap; display: inline-block;
        animation: scroll-rtl 18s linear infinite;
    }}
    @keyframes scroll-rtl {{
        0% {{transform: translateX(-100%);}}
        100% {{transform: translateX(100%);}}
    }}
 .header-calm {{
        background: linear-gradient(135deg, #1A2F4F, #2C4A73);
        padding: 18px; border-radius: 12px; text-align: center;
        border: 2px solid #C9A961; margin-bottom: 20px;
    }}
 .header-calm h1 {{color: #D4B96A; font-size: 30px; font-weight: 800; margin: 0;}}
 .header-calm p {{color: #E8E8E8; font-size: 13px; font-weight: 600; margin: 6px 0 0 0;}}
 .section-title {{color: #C9A961; text-align: center; font-size: 22px; font-weight: 800; margin: 15px 0;}}
 .section-divider {{height: 4px; background: linear-gradient(90deg, transparent, #C9A961, transparent); margin: 15px 0;}}

    label,.stTextInput label,.stSelectbox label,.stTextArea label,.stDateInput label {{
        color: #D4B96A!important;
        font-weight: 700!important;
        font-size: 15px!important;
    }}
    input, textarea, select {{color: #0F1C2E!important; font-weight: 600; background-color: #FFFFFF!important;}}

    div[data-testid="stButton"] > button {{
        background: linear-gradient(135deg, #2C4A73, #3A5F8A)!important;
        color: #E8E8E8!important; width: 100%; padding: 20px 10px; border-radius: 12px;
        border: 2px solid #C9A961; font-weight: 700; font-size: 15px;
        height: 90px; margin: 5px 0;
    }}
 .btn-save button {{background: linear-gradient(135deg, #C9A961, #D4B96A)!important; color: #0F1C2E!important; height: 50px!important; font-weight:800!important}}
 .btn-delete button {{background: linear-gradient(135deg, #6E4A4A, #8A5A5A)!important; color: #E8E8E8!important; height: 50px!important; font-weight:800!important}}

 .card {{
        background: rgba(26,47,79,0.85); padding: 18px; border-radius: 12px;
        border: 2px solid #C9A961; margin-bottom: 18px; box-shadow: 0 0 10px rgba(201,169,97,0.2);
    }}
 .card-title {{color: #D4B96A; font-weight: 800; font-size: 17px; margin-bottom: 12px; border-bottom: 1px solid #C9A961; padding-bottom: 8px;}}

   /* جدول الحصر */
 .table-container {{overflow-x: auto; border: 2px solid #C9A961; border-radius: 10px;}}
 .case-table {{
        width: 100%; border-collapse: collapse; background: #FFFFFF; color: #0F1C2E; font-size: 13px;
    }}
 .case-table th {{
        background: linear-gradient(90deg, #C9A961, #D4B96A); color: #0F1C2E; padding: 10px;
        text-align: center; font-weight: 800; border: 1px solid #C9A961; position: sticky; top: 0;
    }}
 .case-table td {{
        padding: 8px; text-align: center; border: 1px solid #DDD; font-weight: 600;
    }}
 .case-table tr:nth-child(even) {{background: #F5F5F5;}}
 .case-table tr:hover {{background: #FFF8E1;}}
 .btn-open-td button {{
        background: linear-gradient(135deg, #C9A961, #D4B96A)!important;
        color: #0F1C2E!important; height: 28px!important; font-weight:800!important;
        font-size:12px!important; padding:2px 10px!important; border:none!important; border-radius:6px!important;
    }}

 .small-stat {{
        padding: 14px; border-radius: 10px; text-align: center;
        border: 2px solid #C9A961; margin-bottom: 10px;
    }}
 .small-stat p {{color: #E8E8E8; font-size: 12px; margin: 0; font-weight: 600;}}
 .small-stat h2 {{color: #D4B96A; font-size: 28px; font-weight: 800; margin: 4px 0 0 0;}}
 .s1 {{background: linear-gradient(135deg, #3A4F63, #4A6578);}}
 .s2 {{background: linear-gradient(135deg, #2C4A73, #3A5F8A);}}
 .s3 {{background: linear-gradient(135deg, #4A5A4A, #5A6E5A);}}
 .s4 {{background: linear-gradient(135deg, #5A3A3A, #6E4A4A);}}
 .case-detail {{
        background: #1A2F4F; padding: 20px; border-radius: 10px;
        border: 2px solid #C9A961; margin: 10px 0; color: white;
    }}
 .case-detail h3 {{color: #D4B96A; text-align: center;}}
 .stats-title {{color: #C9A961; font-size: 18px; font-weight: 800; margin: 20px 0 10px 0;}}
    h2, h3, h4, p {{color: #FFFFFF!important;}}
</style>
""", unsafe_allow_html=True)

if 'page' not in st.session_state: st.session_state.page = "الرئيسية"
if 'selected_case_id' not in st.session_state: st.session_state.selected_case_id = None

# الشريط المتحرك
st.markdown("""
<div class='marquee-container'>
    <div class='marquee-text'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>
</div>
""", unsafe_allow_html=True)

# الهيدر
st.markdown(f"""
<div class='header-calm'>
    <div style='font-size:40px'>⚖️</div>
    <h1>إدارة القضايا</h1>
    <p>📅 {today}</p>
</div>
""", unsafe_allow_html=True)

st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)

# ============================================================
# =================== الصفحة الرئيسية ======================
# ============================================================
if st.session_state.page == "الرئيسية":
    st.markdown("<div class='section-title'>الأقسام</div>", unsafe_allow_html=True)
    r1c1, r1c2 = st.columns(2)
    with r1c1:
        if st.button("📝\nتسجيل القضايا", use_container_width=True, key="btn_main_1"): st.session_state.page = "تسجيل"; st.rerun()
    with r1c2:
        if st.button("📊\nالحصر العام", use_container_width=True, key="btn_main_2"): st.session_state.page = "حصر"; st.rerun()

    r2c1, r2c2, r2c3, r2c4 = st.columns(4)
    with r2c1:
        if st.button("🔍\nالبحث", use_container_width=True, key="btn_main_3"): st.session_state.page = "بحث"; st.rerun()
    with r2c2:
        if st.button("🔔\nالتنبيهات", use_container_width=True, key="btn_main_4"): st.session_state.page = "تنبيهات"; st.rerun()
    with r2c3:
        if st.button("📑\nالتقارير", use_container_width=True, key="btn_main_5"): st.session_state.page = "تقارير"; st.rerun()
    with r2c4:
        if st.button("🗃️\nالأرشيف", use_container_width=True, key="btn_main_6"): st.session_state.page = "ارشيف"; st.rerun()

    r3c1, r3c2 = st.columns(2)
    with r3c1:
        if st.button("📚\nالمكتبة", use_container_width=True, key="btn_main_7"): st.session_state.page = "مكتبة"; st.rerun()
    with r3c2:
        if st.button("📈\nالإحصائيات", use_container_width=True, key="btn_main_8"): st.session_state.page = "احصائيات"; st.rerun()

    st.markdown("<hr style='border:1px solid #C9A961; margin:20px 0'>", unsafe_allow_html=True)
    st.markdown("<div class='stats-title'>📊 ملخص الإحصائيات</div>", unsafe_allow_html=True)
    total_cases = len([c for c in data["cases"] if c.get("status")!= "منتهية"])
    أحكام_لصالح = len([c for c in data["cases"] if c.get("result") == "لصالح"])
    أحكام_ضد = len([c for c in data["cases"] if c.get("result") == "ضد"])
    st.markdown(f"<div class='small-stat s1'><p>📁 القضايا المتداولة</p><h2>{total_cases}</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='small-stat s2'><p>📅 الجلسات القادمة</p><h2>18</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='small-stat s3'><p>✅ أحكام لصالح</p><h2>{أحكام_لصالح}</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='small-stat s4'><p>❌ أحكام ضد</p><h2>{أحكام_ضد}</h2></div>", unsafe_allow_html=True)

# ============================================================
# =================== 1. تسجيل القضايا ======================
# ============================================================
elif st.session_state.page == "تسجيل":
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#C9A961; text-align:center'>تسجيل القضايا 📝</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", key="btn_back_1"):
        st.session_state.page = "الرئيسية"
        st.rerun()

    نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"], key="sel_type")

    with st.form("form_case"):
        st.markdown("<div class='card'><div class='card-title'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            محكمة_نوع = st.selectbox("المحكمة", ["الابتدائية", "الاستئناف", "النقض", "الإدارية", "القضاء الإدارى", "الإدارية العليا"], key="sel_court_type")
        with col2:
            محكمة_اسم = st.text_input("اسم المحكمة", key="txt_court_name")
        if نوع == "استئناف":
            مأمورية = st.text_input("المأمورية", key="txt_mission")
        else:
            مأمورية = ""
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card'><div class='card-title'>2- بيانات الدعوى</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: رقم = st.text_input("رقم الدعوى / الاستئناف / الطعن", key="txt_num")
        with col2: سنة = st.text_input("السنة القضائية", key="txt_year")
        col1, col2 = st.columns(2)
        with col1: دائرة = st.text_input("الدائرة", key="txt_circle")
        with col2: النوع_تفصيلي = st.text_input("النوع", key="txt_type_detail")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card'><div class='card-title'>3- بيانات الخصوم</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: مدعي = st.text_input("اسم المدعى / المستأنف / الطاعن", key="txt_plaintiff")
        with col2: مدعي_عليه = st.text_input("اسم المدعى عليه / المستأنف ضده / المطعون ضده", key="txt_defendant")
        موضوع = st.text_area("موضوع الدعوى", key="txt_subject")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card'><div class='card-title'>4- بيانات الجلسة</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: تاريخ_جلسة = st.date_input("تاريخ أول جلسة", value=datetime.now(), key="date_session")
        with col2: الرول = st.text_input("الرول", key="txt_role")
        سبب = st.text_input("سبب الجلسة", key="txt_reason")
        ملاحظات = st.text_area("ملاحظات", key="txt_notes")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card'><div class='card-title'>5- التنبيهات والمستندات</div>", unsafe_allow_html=True)
        تنبيه = st.checkbox("تفعيل التنبيهات عبر الواتس اب", key="chk_whats")
        if تنبيه:
            واتس = st.text_input("رقم هاتف واتس اب", key="txt_whats")
        else:
            واتس = ""
        col1, col2 = st.columns(2)
        with col1: مستند_نوع = st.selectbox("تحميل المستندات", ["صحيفة الدعوى", "صحيفة الاستئناف", "صحيفة الطعن"], key="sel_doc")
        with col2: مستند_ملف = st.file_uploader("اختر الملف", key="file_upload")
        st.markdown("</div>", unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("<div class='btn-save'>", unsafe_allow_html=True)
            submitted = st.form_submit_button("💾 حفظ القضية", use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)
        with col2:
            st.markdown("<div class='btn-delete'>", unsafe_allow_html=True)
            deleted = st.form_submit_button("🗑️ حذف القضية", use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        if submitted:
            if not رقم or not سنة:
                st.error("❌ من فضلك ادخل رقم الدعوى والسنة")
            else:
                new_case = {
                    "id": len(data["cases"])+1, "نوع": نوع, "محكمة_نوع": محكمة_نوع, "محكمة_اسم": محكمة_اسم,
                    "مأمورية": مأمورية, "رقم": رقم, "سنة": سنة, "دائرة": دائرة, "النوع": النوع_تفصيلي,
                    "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة),
                    "الرول": الرول, "سبب": سبب, "ملاحظات": ملاحظات, "تنبيه": تنبيه, "واتس": واتس,
                    "جلسات": [{"الرول": الرول, "التاريخ": str(تاريخ_جلسة), "الاجراء": سبب}],
                    "مستندات": [], "status": "متداولة", "result": ""
                }
                data["cases"].append(new_case)
                save_data(data)
                st.success(f"✅ تم حفظ القضية رقم {رقم}/{سنة} بنجاح")
                st.balloons()
                st.session_state.page = "الرئيسية"
                st.rerun()
        if deleted: st.warning("ميزة الحذف سيتم تفعيلها لاحقا")

# ============================================================
# =================== 2. الحصر العام جدول ====================
# ============================================================
elif st.session_state.page == "حصر":
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#C9A961; text-align:center'>📊 الحصر العام</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", key="btn_back_2"):
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

        sorted_cases = sorted(data["cases"], key=lambda x: x.get("تاريخ_جلسة","0000"), reverse=True)

        # نعرض كل قضية في سطر بجدول + زرار
        st.markdown("<div class='table-container'>", unsafe_allow_html=True)

        # الهيدر
        header_cols = st.columns([0.5,1,1,1.2,0.8,1.2,1.2,1.8,1,1,0.7])
        headers = ["م", "الرقم", "النوع", "المحكمة", "الدائرة", "المدعي", "المدعى عليه", "الموضوع", "الجلسة", "الاجراء", "فتح"]
        for col, h in zip(header_cols, headers):
            col.markdown(f"<div style='background:linear-gradient(90deg,#C9A961,#D4B96A); color:#0F1C2E; padding:10px; font-weight:800; text-align:center; border:1px solid #C9A961'>{h}</div>", unsafe_allow_html=True)

        # الصفوف
        for case in sorted_cases:
            row_cols = st.columns([0.5,1,1,1.2,0.8,1.2,1.2,1.8,1,1,0.7])
            with row_cols[0]: st.markdown(f"<div style='background:#FFF; color:#0F1C2E; padding:8px; text-align:center; border:1px solid #DDD; font-weight:600'>{case.get('id')}</div>", unsafe_allow_html=True)
            with row_cols[1]: st.markdown(f"<div style='background:#FFF; color:#0F1C2E; padding:8px; text-align:center; border:1px solid #DDD; font-weight:600'>{case.get('رقم')}/{case.get('سنة')}</div>", unsafe_allow_html=True)
            with row_cols[2]: st.markdown(f"<div style='background:#FFF; color:#0F1C2E; padding:8px; text-align:center; border:1px solid #DDD; font-weight:600'>{case.get('نوع')}</div>", unsafe_allow_html=True)
            with row_cols[3]: st.markdown(f"<div style='background:#FFF; color:#0F1C2E; padding:8px; text-align:center; border:1px solid #DDD; font-weight:600'>{case.get('محكمة_اسم')}</div>", unsafe_allow_html=True)
            with row_cols[4]: st.markdown(f"<div style='background:#FFF; color:#0F1C2E; padding:8px; text-align:center; border:1px solid #DDD; font-weight:600'>{case.get('دائرة')}</div>", unsafe_allow_html=True)
            with row_cols[5]: st.markdown(f"<div style='background:#FFF; color:#0F1C2E; padding:8px; text-align:center; border:1px solid #DDD; font-weight:600'>{case.get('مدعي')}</div>", unsafe_allow_html=True)
            with row_cols[6]: st.markdown(f"<div style='background:#FFF; color:#0F1C2E; padding:8px; text-align:center; border:1px solid #DDD; font-weight:600'>{case.get('مدعي_عليه')}</div>", unsafe_allow_html=True)
            with row_cols[7]: st.markdown(f"<div style='background:#FFF; color:#0F1C2E; padding:8px; text-align:center; border:1px solid #DDD; font-weight:600'>{case.get('موضوع')}</div>", unsafe_allow_html=True)
            with row_cols[8]: st.markdown(f"<div style='background:#FFF; color:#0F1C2E; padding:8px; text-align:center; border:1px solid #DDD; font-weight:600'>{case.get('تاريخ_جلسة')}</div>", unsafe_allow_html=True)
            with row_cols[9]: st.markdown(f"<div style='background:#FFF; color:#0F1C2E; padding:8px; text-align:center; border:1px solid #DDD; font-weight:600'>{case.get('سبب')}</div>", unsafe_allow_html=True)
            with row_cols[10]:
                if st.button("فتح", key=f"open_{case['id']}"):
                    st.session_state.selected_case_id = case['id']
                    st.session_state.page = "تفاصيل"
                    st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# =================== 3. تفاصيل القضية ======================
# ============================================================
elif st.session_state.page == "تفاصيل":
    case = next((c for c in data["cases"] if c["id"] == st.session_state.selected_case_id), None)
    if case:
        st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='color:#C9A961; text-align:center'>📄 تفاصيل القضية رقم {case['رقم']}/{case['سنة']}</h2>", unsafe_allow_html=True)
        if st.button("⬅️ العودة للحصر", key="btn_back_detail"):
            st.session_state.page = "حصر"
            st.rerun()
        st.markdown(f"""
        <div class='case-detail'>
            <h3>البيانات الأساسية</h3>
            <p><b>نوع الدعوى:</b> {case['نوع']}</p>
            <p><b>المحكمة:</b> {case['محكمة_نوع']} - {case['محكمة_اسم']} {case['مأمورية']}</p>
            <p><b>الدائرة:</b> {case['دائرة']} | <b>النوع:</b> {case['النوع']}</p>
            <p><b>الخصوم:</b> {case['مدعي']} ضد {case['مدعي_عليه']}</p>
            <p><b>الموضوع:</b> {case['موضوع']}</p>
            <p><b>آخر جلسة:</b> {case['تاريخ_جلسة']} - الرول: {case['الرول']} - {case['سبب']}</p>
            <p><b>ملاحظات:</b> {case['ملاحظات']}</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.error("القضية غير موجودة")
        st.session_state.page = "حصر"
        st.rerun()

elif st.session_state.page == "بحث":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>🔍 البحث عن دعوى</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", key="btn_back_3"): st.session_state.page = "الرئيسية"; st.rerun()
elif st.session_state.page == "تنبيهات":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>🔔 التنبيهات</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", key="btn_back_4"): st.session_state.page = "الرئيسية"; st.rerun()
elif st.session_state.page == "تقارير":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>📑 التقارير</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية",
