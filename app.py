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
    
    /* الشريط المتحرك صح RTL */
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
    
    /* هيدر هادي فخم */
    .header-calm {{
        background: linear-gradient(135deg, #1A2F4F, #2C4A73);
        padding: 18px; border-radius: 12px; text-align: center;
        border: 2px solid #C9A961; margin-bottom: 20px;
    }}
    .header-calm h1 {{color: #D4B96A; font-size: 30px; font-weight: 800; margin: 0;}}
    .header-calm p {{color: #E8E8E8; font-size: 13px; font-weight: 600; margin: 6px 0 0 0;}}
    
    /* عنوان الاقسام */
    .section-title {{color: #C9A961; text-align: center; font-size: 22px; font-weight: 800; margin: 15px 0;}}
    
    /* الاقسام فوق - 2 و 4 و 2 */
    .section-btn {{
        background: linear-gradient(135deg, #2C4A73, #3A5F8A) !important;
        color: #E8E8E8 !important;
        width: 100%; padding: 20px 10px; border-radius: 12px; 
        border: 2px solid #C9A961; font-weight: 700; font-size: 15px;
        height: 90px; margin: 5px 0;
        display: flex; flex-direction: column; justify-content: center; align-items: center;
        transition: 0.3s;
    }}
    .section-btn:hover {{background: linear-gradient(135deg, #3A5F8A, #4A76A8) !important; border-color: #D4B96A;}}
    .icon {{font-size: 26px; margin-bottom: 6px;}}
    
    /* كروت الاحصائيات الصغيرة تحت بعض بالوان هادئة */
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
    
    .stats-title {{color: #C9A961; font-size: 18px; font-weight: 800; margin: 20px 0 10px 0;}}
</style>
""", unsafe_allow_html=True)

if 'page' not in st.session_state: st.session_state.page = "الرئيسية"

# الشريط المتحرك
st.markdown("""
<div class='marquee-container'>
    <div class='marquee-text'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div>
</div>
""", unsafe_allow_html=True)

# الهيدر
st.markdown(f"""
<div class='header-calm'>
    <h1>⚖️ إدارة القضايا</h1>
    <p>📅 {today}</p>
</div>
""", unsafe_allow_html=True)

# الصفحة الرئيسية
if st.session_state.page == "الرئيسية":
    
    st.markdown("<div class='section-title'>الأقسام</div>", unsafe_allow_html=True)
    
    # الاقسام فوق: 2
    r1c1, r1c2 = st.columns(2)
    with r1c1:
        if st.button("📝\nتسجيل القضايا", use_container_width=True, key="b1"): st.session_state.page = "تسجيل"
    with r1c2:
        if st.button("📊\nالحصر العام", use_container_width=True, key="b2"): st.session_state.page = "حصر"
    
    # 4 في النص
    r2c1, r2c2, r2c3, r2c4 = st.columns(4)
    with r2c1:
        if st.button("🔍\nالبحث", use_container_width=True, key="b3"): st.session_state.page = "بحث"
    with r2c2:
        if st.button("🔔\nالتنبيهات", use_container_width=True, key="b4"): st.session_state.page = "تنبيهات"
    with r2c3:
        if st.button("📑\nالتقارير", use_container_width=True, key="b5"): st.session_state.page = "تقارير"
    with r2c4:
        if st.button("🗃️\nالأرشيف", use_container_width=True, key="b6"): st.session_state.page = "ارشيف"
    
    # 2 تحت
    r3c1, r3c2 = st.columns(2)
    with r3c1:
        if st.button("📚\nالمكتبة", use_container_width=True, key="b7"): st.session_state.page = "مكتبة"
    with r3c2:
        if st.button("📈\nالإحصائيات", use_container_width=True, key="b8"): st.session_state.page = "احصائيات"
    
    st.markdown("<hr style='border:1px solid #C9A961; margin:20px 0'>", unsafe_allow_html=True)
    
    # الاحصائيات تحت
    st.markdown("<div class='stats-title'>📊 ملخص الإحصائيات</div>", unsafe_allow_html=True)
    
    total_cases = len(data["cases"])
    st.markdown(f"<div class='small-stat s1'><p>📁 القضايا المتداولة</p><h2>{total_cases}</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='small-stat s2'><p>📅 الجلسات القادمة</p><h2>18</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='small-stat s3'><p>✅ أحكام لصالح</p><h2>42</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='small-stat s4'><p>❌ أحكام ضد</p><h2>7</h2></div>", unsafe_allow_html=True)

# صفحة التسجيل
elif st.session_state.page == "تسجيل":
    st.markdown("<h2 style='color:#C9A961'>📝 تسجيل القضايا</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    with st.form("form"):
        نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
        محكمة = st.text_input("اسم المحكمة")
        رقم = st.text_input("رقم القضية")
        if st.form_submit_button("💾 حفظ"):
            data["cases"].append({"نوع": نوع, "محكمة": محكمة, "رقم": رقم})
            save_data(data)
            st.success("✅ تم الحفظ")

# صفحة الحصر
elif st.session_state.page == "حصر":
    st.markdown("<h2 style='color:#C9A961'>📊 الحصر العام</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    if data["cases"]:
        st.dataframe(pd.DataFrame(data["cases"]), use_container_width=True, hide_index=True)
    else:
        st.info("لا توجد قضايا مسجلة")

# باقي الصفحات...
elif st.session_state.page == "تنبيهات":
    st.markdown("<h2 style='color:#C9A961'>🔔 التنبيهات</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
        import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

DATA_FILE = "cases_data.json"
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f: 
            return json.load(f)
    return {"cases": []}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f: 
        json.dump(data, f, ensure_ascii=False, indent=4)

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
    .marquee-container {{background: linear-gradient(90deg, #C9A961, #D4B96A); padding: 10px; overflow: hidden; border-radius: 8px; margin-bottom: 15px;}}
    .marquee-text {{color: #0F1C2E; font-weight: 800; font-size: 14px; white-space: nowrap; display: inline-block; animation: scroll-rtl 18s linear infinite;}}
    @keyframes scroll-rtl {{0% {{transform: translateX(-100%);}} 100% {{transform: translateX(100%);}}}}
    .header-calm {{background: linear-gradient(135deg, #1A2F4F, #2C4A73); padding: 18px; border-radius: 12px; text-align: center; border: 2px solid #C9A961; margin-bottom: 10px;}}
    .header-calm h1 {{color: #D4B96A; font-size: 30px; font-weight: 800; margin: 0;}}
    .section-divider {{height: 4px; background: linear-gradient(90deg, transparent, #C9A961, transparent); margin: 0 0 20px 0;}}
    div[data-testid="stButton"] > button {{background: linear-gradient(135deg, #2C4A73, #3A5F8A) !important; color: #E8E8E8 !important; width: 100%; padding: 20px 10px; border-radius: 12px; border: 2px solid #C9A961; font-weight: 700; font-size: 15px; height: 90px; margin: 5px 0;}}
    .small-stat {{padding: 14px; border-radius: 10px; text-align: center; border: 2px solid #C9A961; margin-bottom: 10px;}}
    .small-stat p {{color: #E8E8E8; font-size: 12px; margin: 0; font-weight: 600;}}
    .small-stat h2 {{color: #D4B96A; font-size: 28px; font-weight: 800; margin: 4px 0 0 0;}}
    .s1 {{background: linear-gradient(135deg, #3A4F63, #4A6578);}}
    .s2 {{background: linear-gradient(135deg, #2C4A73, #3A5F8A);}}
    .s3 {{background: linear-gradient(135deg, #4A5A4A, #5A6E5A);}}
    .s4 {{background: linear-gradient(135deg, #5A3A3A, #6E4A4A);}}
    .case-card {{background: #1A2F4F; padding: 15px; border-radius: 10px; border: 2px solid #C9A961; margin: 10px 0; color: white;}}
    .case-card h4 {{color: #D4B96A;}}
    h2, h3, h4, p, label {{color: #FFFFFF !important;}}
</style>
""", unsafe_allow_html=True)

if 'page' not in st.session_state: st.session_state.page = "الرئيسية"
if 'selected_case_id' not in st.session_state: st.session_state.selected_case_id = None

st.markdown("<div class='marquee-container'><div class='marquee-text'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعى</div></div>", unsafe_allow_html=True)
st.markdown(f"<div class='header-calm'><div style='font-size:40px'>⚖️</div><h1>إدارة القضايا</h1><p>📅 {today}</p></div>", unsafe_allow_html=True)
st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)

# ============================================================
# ===================  الصفحة الرئيسية  ======================
# ============================================================
if st.session_state.page == "الرئيسية":
    
    st.markdown("<h3 style='color:#C9A961; text-align:center'>الأقسام</h3>", unsafe_allow_html=True)
    
    r1c1, r1c2 = st.columns(2)
    with r1c1:
        if st.button("📝\nتسجيل القضايا", use_container_width=True, key="b1"): st.session_state.page = "تسجيل"
    with r1c2:
        if st.button("📊\nالحصر العام", use_container_width=True, key="b2"): st.session_state.page = "حصر"
    
    r2c1, r2c2, r2c3, r2c4 = st.columns(4)
    with r2c1:
        if st.button("🔍\nالبحث", use_container_width=True, key="b3"): st.session_state.page = "بحث"
    with r2c2:
        if st.button("🔔\nالتنبيهات", use_container_width=True, key="b4"): st.session_state.page = "تنبيهات"
    with r2c3:
        if st.button("📑\nالتقارير", use_container_width=True, key="b5"): st.session_state.page = "تقارير"
    with r2c4:
        if st.button("🗃️\nالأرشيف", use_container_width=True, key="b6"): st.session_state.page = "ارشيف"
    
    r3c1, r3c2 = st.columns(2)
    with r3c1:
        if st.button("📚\nالمكتبة", use_container_width=True, key="b7"): st.session_state.page = "مكتبة"
    with r3c2:
        if st.button("📈\nالإحصائيات", use_container_width=True, key="b8"): st.session_state.page = "احصائيات"
    
    st.markdown("<hr style='border:1px solid #C9A961; margin:20px 0'>", unsafe_allow_html=True)
    
    st.markdown("<h4 style='color:#C9A961'>📊 ملخص الإحصائيات</h4>", unsafe_allow_html=True)
    total_cases = len([c for c in data["cases"] if c.get("status") != "منتهية"])
    st.markdown(f"<div class='small-stat s1'><p>📁 القضايا المتداولة</p><h2>{total_cases}</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='small-stat s2'><p>📅 الجلسات القادمة</p><h2>18</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='small-stat s3'><p>✅ أحكام لصالح</p><h2>42</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='small-stat s4'><p>❌ أحكام ضد</p><h2>7</h2></div>", unsafe_allow_html=True)

# ============================================================
# =================== 1. تسجيل القضايا  =====================
# ============================================================
elif st.session_state.page == "تسجيل":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>📝 تسجيل القضايا</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    
    with st.form("form_case"):
        نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
        محكمة_نوع = st.selectbox("المحكمة", ["الابتدائية", "الاستئناف", "النقض", "الإدارية", "القضاء الإدارى", "الإدارية العليا"])
        محكمة_اسم = st.text_input("اسم المحكمة")
        
        مأمورية = ""
        if نوع == "استئناف":
            مأمورية = st.text_input("المأمورية")
        
        رقم = st.text_input("رقم الدعوى / الاستئناف / الطعن")
        سنة = st.text_input("السنة القضائية")
        دائرة = st.text_input("الدائرة")
        النوع_تفصيلي = st.text_input("النوع")
        مدعي = st.text_input("اسم المدعى / المستأنف / الطاعن")
        مدعي_عليه = st.text_input("اسم المدعى عليه / المستأنف ضده / المطعون ضده")
        موضوع = st.text_area("موضوع الدعوى")
        تاريخ_جلسة = st.date_input("تاريخ أول جلسة", value=datetime.now())
        الرول = st.text_input("الرول")
        سبب = st.text_input("سبب الجلسة")
        ملاحظات = st.text_area("ملاحظات")
        
        تنبيه = st.checkbox("تفعيل التنبيهات عبر الواتس اب")
        واتس = st.text_input("رقم هاتف واتس اب") if تنبيه else ""
        
        مستند_نوع = st.selectbox("تحميل المستندات", ["صحيفة الدعوى", "صحيفة الاستئناف", "صحيفة الطعن"])
        مستند_ملف = st.file_uploader("اختر الملف")
        
        if st.form_submit_button("💾 حفظ القضية"):
            new_case = {
                "id": len(data["cases"])+1, "نوع": نوع, "محكمة_نوع": محكمة_نوع, "محكمة_اسم": محكمة_اسم,
                "مأمورية": مأمورية, "رقم": رقم, "سنة": سنة, "دائرة": دائرة, "النوع": النوع_تفصيلي,
                "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة),
                "الرول": الرول, "سبب": سبب, "ملاحظات": ملاحظات, "تنبيه": تنبيه, "واتس": واتس,
                "جلسات": [{"الرول": الرول, "التاريخ": str(تاريخ_جلسة), "الاجراء": سبب}],
                "مستندات": [], "status": "متداولة"
            }
            data["cases"].append(new_case)
            save_data(data)
            st.success("✅ تم الحفظ")
            st.session_state.page = "حصر"

# ============================================================
# ===================  2. الحصر العام  =======================
# ============================================================
elif st.session_state.page == "حصر":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>📊 الحصر العام</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    
    if not data["cases"]:
        st.info("لا توجد قضايا مسجلة")
    else:
        for case in sorted(data["cases"], key=lambda x: x["id"]):
            st.markdown(f"<div class='case-card'><h4>قضية رقم: {case['رقم']} / {case['سنة']}</h4><p><b>الدائرة:</b> {case['دائرة']} | <b>النوع:</b> {case['نوع']}</p><p><b>المحكمة:</b> {case['محكمة_اسم']} {case['مأمورية']}</p><p><b>الخصوم:</b> {case['مدعي']} ضد {case['مدعي_عليه']}</p><p><b>الموضوع:</b> {case['موضوع']}</p><p><b>آخر جلسة:</b> {case['تاريخ_جلسة']} - {case['سبب']}</p></div>", unsafe_allow_html=True)
            if st.button("فتح القضية", key=f"open{case['id']}"):
                st.session_state.selected_case_id = case['id']
                st.session_state.page = "تفاصيل_قضية"

# ============================================================
# ============== 3. تفاصيل القضية والجلسات  ================
# ============================================================
elif st.session_state.page == "تفاصيل_قضية":
    case = next((c for c in data["cases"] if c["id"] == st.session_state.selected_case_id), None)
    if case:
        st.markdown(f"<h2 style='color:#C9A961; text-align:center'>📄 القضية رقم {case['رقم']} / {case['سنة']}</h2>", unsafe_allow_html=True)
        if st.button("⬅️ العودة للحصر"): st.session_state.page = "حصر"
        
        st.markdown("<h4 style='color:#C9A961'>جدول الجلسات والإجراءات</h4>", unsafe_allow_html=True)
        st.dataframe(pd.DataFrame(case["جلسات"]), use_container_width=True, hide_index=True)
        
        with st.form("add_session"):
            st.markdown("<h4>متابعة الجلسات - إضافة قرار</h4>", unsafe_allow_html=True)
            تاريخ_قادم = st.date_input("تاريخ الجلسة القادمة")
            رول_قادم = st.text_input("الرول")
            سبب_تأجيل = st.text_input("سبب التأجيل")
            if st.form_submit_button("حفظ الجلسة"):
                case["جلسات"].append({"الرول": رول_قادم, "التاريخ": str(تاريخ_قادم), "الاجراء": سبب_تأجيل})
                case["تاريخ_جلسة"] = str(تاريخ_قادم)
                case["سبب"] = سبب_تأجيل
                save_data(data)
                st.success("تمت الإضافة")
                st.rerun()

# ============================================================
# ===================  4. البحث عن دعوى  =====================
# ============================================================
elif st.session_state.page == "بحث":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>🔍 البحث عن دعوى</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"

# ============================================================
# ===================  5. التنبيهات  =========================
# ============================================================
elif st.session_state.page == "تنبيهات":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>🔔 التنبيهات</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"

# ============================================================
# ===================  6. التقارير  ==========================
# ============================================================
elif st.session_state.page == "تقارير":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>📑 التقارير</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"

# ============================================================
# ===================  7. الأرشيف  ===========================
# ============================================================
elif st.session_state.page == "ارشيف":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>🗃️ الأرشيف</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"

# ============================================================
# ===================  8. المكتبة القانونية  ==================
# ============================================================
elif st.session_state.page == "مكتبة":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>📚 المكتبة القانونية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"

# ============================================================
# =================== 9. الإحصائيات  ========================
# ============================================================
elif st.session_state.page == "احصائيات":
    st.markdown("<h2 style='color:#C9A961; text-align:center'>📈 الإحصائيات</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
