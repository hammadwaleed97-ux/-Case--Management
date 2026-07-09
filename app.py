import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import json

st.set_page_config(page_title="إدارة القضايا", layout="wide", initial_sidebar_state="expanded")

# الالوان والتصميم الفخم
st.markdown("""
<style>
   .stApp {
        background: linear-gradient(135deg, #0a1a3a 0%, #1e3a5f 100%);
        background-image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" viewBox="0 0 100 100"><text y="50" font-size="40" opacity="0.03">⚖️</text></svg>');
    }
    h1, h2, h3 { color: #DAA520!important; text-align: center; }
   .logo-text {
        animation: blink 3s infinite;
        color: red; font-weight: bold; text-align: center; font-size: 18px;
    }
    @keyframes blink { 0%,100%{opacity:1} 50%{opacity:0} }
   .card {
        background: #1e3a5f; padding: 20px; border-radius: 15px;
        border: 2px solid #DAA520; margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# اللوجو
st.markdown("<h1>⚖️</h1>", unsafe_allow_html=True)
st.markdown("<h1>إدارة القضايا</h1>", unsafe_allow_html=True)

st.markdown("<p class='logo-text'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة - الهيئة القومية للتأمين الاجتماعي</p>", unsafe_allow_html=True)
st.markdown("---")

# القائمة الجانبية
menu = st.sidebar.selectbox("📋 القائمة الرئيسية",
["تسجيل القضايا", "الحصر العام", "التنبيهات", "التقارير", "الأرشيف", "المكتبة القانونية", "الإحصائيات"],
index=0)

# قاعدة البيانات المؤقتة
if 'cases' not in st.session_state:
    st.session_state.cases = []

# 1. تسجيل القضايا
if menu == "تسجيل القضايا":
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("📝 تسجيل قضية جديدة")

    col1, col2 = st.columns(2)
    with col1:
        نوع_الدعوى = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
        المحكمة_النوع = st.selectbox("نوع المحكمة", ["الابتدائية", "الاستئناف", "النقض", "الإدارية", "القضاء الإداري", "الإدارية العليا"])
        اسم_المحكمة = st.text_input("اسم المحكمة")
        if نوع_الدعوى == "استئناف":
            المأمورية = st.text_input("المأمورية")
        رقم_الدعوى = st.text_input("رقم الدعوى/الاستئناف/الطعن")
        السنة = st.text_input("السنة القضائية")
        الدائرة = st.text_input("الدائرة")

    with col2:
        النوع = st.text_input("نوع القضية")
        المدعي = st.text_input("اسم المدعي/المستأنف/الطاعن")
        المدعي_عليه = st.text_input("اسم المدعي عليه/المستأنف ضده/المطعون ضده")
        الموضوع = st.text_area("موضوع الدعوى")
        تاريخ_الجلسة = st.date_input("تاريخ أول جلسة")
        الرول = st.text_input("الرول")
        سبب_الجلسة = st.text_input("سبب الجلسة")
        ملاحظات = st.text_area("ملاحظات")

    تنبيه_واتس = st.checkbox("تفعيل التنبيهات عبر الواتس اب")
    رقم_الواتس = st.text_input("رقم الواتس اب") if تنبيه_واتس else ""
    المستند = st.file_uploader(f"تحميل {نوع_الدعوى}", type=['pdf', 'docx'])

    if st.button("💾 حفظ القضية", use_container_width=True):
        قضية = {
            "رقم": f"{رقم_الدعوى} لسنة {السنة}", "نوع_الدعوى": نوع_الدعوى,
            "المحكمة": اسم_المحكمة, "الدائرة": الدائرة, "المدعي": المدعي,
            "المدعي_عليه": المدعي_عليه, "الموضوع": الموضوع, "تاريخ_الجلسة": str(تاريخ_الجلسة),
            "الرول": الرول, "السبب": سبب_الجلسة, "الحالة": "متداولة", "الجلسات": []
        }
        st.session_state.cases.append(قضية)
        st.success("تم حفظ القضية بنجاح! انتقلت للحصر العام")
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# 2. الحصر العام
elif menu == "الحصر العام":
    st.subheader("📊 الحصر العام للقضايا")
    if len(st.session_state.cases) == 0:
        st.warning("لا توجد قضايا مسجلة")
    else:
        for i, ق in enumerate(st.session_state.cases):
            with st.expander(f"{ق['رقم']} - {ق['المدعي']} ضد {ق['المدعي_عليه']}"):
                st.write(f"**المحكمة:** {ق['المحكمة']} | **الدائرة:** {ق['الدائرة']}")
                st.write(f"**الموضوع:** {ق['الموضوع']}")
                st.write(f"**آخر جلسة:** {ق['تاريخ_الجلسة']} | **السبب:** {ق['السبب']}")

                st.markdown("### جدول الجلسات")
                df_جلسات = pd.DataFrame([{"الرول": ق['الرول'], "التاريخ": ق['تاريخ_الجلسة'], "الإجراءات": ق['السبب']}])
                st.dataframe(df_جلسات, use_container_width=True)

                تاريخ_جديد = st.date_input("تاريخ الجلسة القادمة", key=f"d{i}")
                رول_جديد = st.text_input("الرول الجديد", key=f"r{i}")
                سبب_جديد = st.text_input("سبب التأجيل", key=f"s{i}")
                if st.button("حفظ الجلسة", key=f"b{i}"):
                    st.session_state.cases[i]['الجلسات'].append({"الرول": رول_جديد, "التاريخ": str(تاريخ_جديد), "السبب": سبب_جديد})
                    st.success("تمت الإضافة")

st.sidebar.info(": التنبيهات والتقارير")
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="إدارة القضايا", layout="wide")

# التصميم الفخم
st.markdown("""
<style>
    .stApp {background: linear-gradient(135deg, #0a1a3a 0%, #1e3a5f 100%);}
    h1, h2, h3 { color: #DAA520!important; text-align: center; }
    .logo-text {animation: blink 3s infinite; color: #FFD700; font-weight: bold; text-align: center;}
    @keyframes blink { 0%,100%{opacity:1} 50%{opacity:0} }
    .card {background: #1e3a5f; padding: 20px; border-radius: 15px; border: 2px solid #DAA520; margin: 10px 0;}
    .alert {background: #8B0000; color: white; padding: 15px; border-radius: 10px;}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1>⚖️</h1>", unsafe_allow_html=True)
st.markdown("<h1>إدارة القضايا</h1>", unsafe_allow_html=True)
st.markdown("<p class='logo-text'>مع تحيات / وليد حماد - الإدارة العامة للشئون القانونية - ديوان عام منطقة البحيرة</p>", unsafe_allow_html=True)
st.markdown("---")

if 'cases' not in st.session_state: st.session_state.cases = []
if 'archive' not in st.session_state: st.session_state.archive = []

menu = st.sidebar.selectbox("📋 القائمة", 
["تسجيل القضايا", "الحصر العام", "البحث عن دعوى", "التنبيهات", "التقارير", "الأرشيف", "المكتبة القانونية", "الإحصائيات"])

# 1. تسجيل القضايا - نفس اللي فات
if menu == "تسجيل القضايا":
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("📝 تسجيل قضية جديدة")
    with st.form("form"):
        c1, c2 = st.columns(2)
        with c1:
            نوع = st.selectbox("نوع الدعوى", ["دعوى", "استئناف", "طعن"])
            محكمة = st.selectbox("نوع المحكمة", ["الابتدائية","الاستئناف","النقض","الإدارية","القضاء الإداري","الإدارية العليا"])
            اسم_المحكمة = st.text_input("اسم المحكمة")
            رقم = st.text_input("رقم الدعوى")
            سنة = st.text_input("السنة")
            دائرة = st.text_input("الدائرة")
        with c2:
            مدعي = st.text_input("المدعي")
            مدعي_عليه = st.text_input("المدعي عليه")
            موضوع = st.text_area("الموضوع")
            تاريخ = st.date_input("تاريخ أول جلسة")
            رول = st.text_input("الرول")
            سبب = st.text_input("سبب الجلسة")
        واتس = st.checkbox("تفعيل تنبيهات الواتس")
        رقم_واتس = st.text_input("رقم الواتس") if واتس else ""
        if st.form_submit_button("💾 حفظ القضية"):
            st.session_state.cases.append({
                "رقم": f"{رقم} لسنة {سنة}", "نوع": نوع, "المحكمة": اسم_المحكمة, "الدائرة": دائرة,
                "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, "موضوع": موضوع, "تاريخ": str(تاريخ),
                "رول": رول, "سبب": سبب, "الحالة": "متداولة", "واتس": رقم_واتس
            })
            st.success("تم الحفظ")
    st.markdown("</div>", unsafe_allow_html=True)

# 2. الحصر العام
elif menu == "الحصر العام":
    st.subheader("📊 الحصر العام")
    for i, ق in enumerate(st.session_state.cases):
        with st.expander(f"{ق['رقم']} - {ق['مدعي']}"):
            st.write(f"المحكمة: {ق['المحكمة']} | الموضوع: {ق['موضوع']}")
            st.write(f"آخر جلسة: {ق['تاريخ']} | السبب: {ق['سبب']}")
            if st.button("أرشفة القضية", key=i):
                st.session_state.archive.append(ق)
                st.session_state.cases.pop(i)
                st.rerun()

# 3. البحث
elif menu == "البحث عن دعوى":
    st.subheader("🔍 البحث")
    بحث = st.text_input("ابحث برقم القضية أو اسم المدعي")
    if st.button("اضغط للبحث"):
        نتائج = [ق for ق in st.session_state.cases + st.session_state.archive if بحث in ق['رقم'] or بحث in ق['مدعي']]
        if نتائج:
            for ق in نتائج: st.json(ق)
        else: st.error("لا توجد قضية. تأكد من بيانات البحث")

# 4. التنبيهات
elif menu == "التنبيهات":
    st.subheader("🔔 التنبيهات")
    st.markdown("### 1. تنبيهات الجلسات - خلال 7 أيام")
    اليوم = datetime.now().date()
    for ق in st.session_state.cases:
        ت_الجلسة = datetime.strptime(ق['تاريخ'], '%Y-%m-%d').date()
        if 0 <= (ت_الجلسة - اليوم).days <= 7:
            st.markdown(f"<div class='alert'>⚠️ جلسة بتاريخ {ق['تاريخ']} - {ق['رقم']}</div>", unsafe_allow_html=True)

    st.markdown("### 2. تنبيهات مواعيد الطعن - خلال 15 يوم")
    st.info("40 يوم استئناف | 60 يوم نقض")
    تفعيل = st.checkbox("تفعيل تنبيهات الواتس الساعة 9 صباحاً")
    if تفعيل: st.text_input("رقم الواتس للتنبيهات")

# 5. التقارير
elif menu == "التقارير":
    st.subheader("📑 التقارير")
    نوع_التقرير = st.selectbox("نوع التقرير", ["المتداولة", "الأحكام"])
    من = st.date_input("من")
    إلى = st.date_input("إلى")
    المحامي = st.text_input("طرف الاستاذ المحامي")
    منطقة = st.text_input("ديوان عام منطقة")

    if st.button("إنشاء التقرير"):
        df = pd.DataFrame(st.session_state.cases)
        st.dataframe(df, use_container_width=True)
        
        # تصدير
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("📥 تحميل Excel", csv, "تقرير.csv")
        st.button("📄 فتح PDF") # هنضيف مكتبة reportlab بعدين
        st.button("📝 فتح Word")

# 6. الأرشيف
elif menu == "الأرشيف":
    st.subheader("🗃️ الأرشيف")
    for ق in st.session_state.archive:
        st.write(f"{ق['رقم']} - {ق['مدعي']} - منتهية")

# 7. المكتبة
elif menu == "المكتبة القانونية":
    st.subheader("📚 المكتبة القانونية")
    اقسام = ["القوانين","القرارات الوزارية","احكام النقض","احكام الدستورية","أخرى"]
    for قسم in اقسام:
        with st.expander(f"📁 {قسم}"):
            st.file_uploader(f"رفع مستند {قسم}", key=قسم)
            st.text_input("اسم المستند", key=f"name{قسم}")

# 8. الإحصائيات
elif menu == "الإحصائيات":
    st.subheader("📈 الإحصائيات")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("المتداولة", len(st.session_state.cases))
    c2.metric("المؤرشفة", len(st.session_state.archive))
    c3.metric("للمصلحة", 0)
    c4.metric("ضد المصلحة", 0)
