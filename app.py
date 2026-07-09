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
        # ============================================================
# =================== 4. البحث عن دعوى  =====================
# ============================================================
elif st.session_state.page == "بحث":
    st.markdown("<h2 style='color:#C9A961'>🔍 البحث عن دعوى</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    
    بحث = st.text_input("ابحث باسم المدعي أو رقم القضية أو السنة")
    if st.button("اضغط للبحث"):
        if بحث:
            results = [c for c in data["cases"] if بحث.lower() in str(c).lower()]
            if results: 
                st.dataframe(pd.DataFrame(results), use_container_width=True, hide_index=True)
            else: 
                st.error("لا توجد نتائج. تأكد من بيانات البحث")
        else:
            st.warning("من فضلك ادخل كلمة للبحث")

# ============================================================
# ===================  5. التنبيهات  =========================
# ============================================================
elif st.session_state.page == "تنبيهات":
    st.markdown("<h2 style='color:#C9A961'>🔔 التنبيهات</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    
    st.info("📅 التنبيهات: جلسات خلال 7 ايام + مواعيد الطعن 40 يوم و 60 يوم")
    st.warning("سيتم ربطها بالواتس اب لاحقاً")

# ============================================================
# =================== 6. التقارير  ==========================
# ============================================================
elif st.session_state.page == "تقارير":
    st.markdown("<h2 style='color:#C9A961'>📑 التقارير</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    
    نوع_تقرير = st.selectbox("اختر نوع التقرير", ["الدعاوى المتداولة", "الاحكام الصادرة"])
    ديوان = st.text_input("ديوان عام منطقة")
    اسم_المحامي = st.text_input("اسم المحامي")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📄 تصدير PDF"): st.success("سيتم التصدير PDF")
    with col2:
        if st.button("📝 تصدير Word"): st.success("سيتم التصدير Word")

# ============================================================
# =================== 7. الأرشيف  ===========================
# ============================================================
elif st.session_state.page == "ارشيف":
    st.markdown("<h2 style='color:#C9A961'>🗃️ الأرشيف</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    
    ended = [c for c in data["cases"] if c.get("status") == "منتهية"]
    if ended:
        st.dataframe(pd.DataFrame(ended), use_container_width=True, hide_index=True)
    else:
        st.info("لا توجد قضايا منتهية في الأرشيف")

# ============================================================
# =================== 8. المكتبة القانونية  ==================
# ============================================================
elif st.session_state.page == "مكتبة":
    st.markdown("<h2 style='color:#C9A961'>📚 المكتبة القانونية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    
    قسم = st.selectbox("اختر القسم", ["القوانين", "القرارات الوزارية", "احكام المحكمة الدستورية", "احكام محكمة النقض", "أخرى"])
    st.file_uploader("ارفع المستند")
    st.info("سيتم حفظ الملفات في مجلد المكتبة")

# ============================================================
# =================== 9. الإحصائيات  ========================
# ============================================================
elif st.session_state.page == "احصائيات":
    st.markdown("<h2 style='color:#C9A961'>📈 الإحصائيات التفصيلية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية"): st.session_state.page = "الرئيسية"
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("اجمالي القضايا", len(data["cases"]))
    with col2:
        st.metric("القضايا المتداولة", len([c for c in data["cases"] if c.get("status") != "منتهية"]))
    
    st.bar_chart({"المتداولة": len([c for c in data["cases"] if c.get("status") != "منتهية"]), "المنتهية": len([c for c in data["cases"] if c.get("status") == "منتهية"])})

# ============================================================
# ====================== نهاية الكود =========================
# ========== مع تحيات وليد حماد - الشئون القانونية ==========
# ============================================================
