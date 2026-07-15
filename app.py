
# ================================================
# ========== الصفحة الرئيسية =======
# ================================================
# ========== الصفحة الرئيسية =======
if st.session_state.page == "الرئيسية":
    alerts = get_alert_cases()
    total_alerts = len(alerts["sessions"]) + len(alerts["appeals"])
    
    if total_alerts > 0:
        st.markdown(f'<div class="btn-alert"><button>🔔 لديك {total_alerts} تنبيهات</button></div>', unsafe_allow_html=True)
        if st.button("عرض التنبيهات", key="show_alerts_btn"):
            st.session_state.page = "التنبيهات"

    st.markdown('<div class="btn-add">', unsafe_allow_html=True)
    if st.button("➕ اضافة قضية جديدة", key="add_case_main_btn"):
        st.session_state.page = "اضافة قضية"
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="btn-list">', unsafe_allow_html=True)
    if st.button("📋 عرض جميع القضايا", key="list_case_main_btn"):
        st.session_state.page = "عرض القضايا"
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="btn-report">', unsafe_allow_html=True)
    if st.button("📊 التقارير", key="report_main_btn"):
        st.session_state.page = "التقارير"
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="btn-lib">', unsafe_allow_html=True)
    if st.button("📚 المكتبة القانونية", key="lib_main_btn"):
        st.session_state.page = "المكتبة"
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="btn-arch">', unsafe_allow_html=True)
    if st.button("📦 الارشيف", key="arch_main_btn"):
        st.session_state.page = "الارشيف"
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="btn-search">', unsafe_allow_html=True)
    if st.button("🔍 بحث متقدم", key="search_main_btn"):
        st.session_state.page = "بحث"
    st.markdown('</div>', unsafe_allow_html=True)
# ================================================
# ========== الصفحة الرئيسية =======
# ================================
# ================================================
# ========== الصفحة الرئيسية =========

data = load_data()
alerts = get_alert_cases()
total_alerts = len(alerts['sessions']) + len(alerts['appeals'])

col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="btn-add">', unsafe_allow_html=True)
    if st.button("➕ اضافة قضية جديدة"): st.session_state.page = "اضافة"
    st.markdown('</div>', unsafe_allow_html=True)
with col2:
    st.markdown('<div class="btn-list">', unsafe_allow_html=True)
    if st.button("📋 عرض كل القضايا"): st.session_state.page = "عرض"
    st.markdown('</div>', unsafe_allow_html=True)

col3, col4 = st.columns(2)
with col3:
    st.markdown('<div class="btn-alert">', unsafe_allow_html=True)
    if st.button(f"🚨 التنبيهات ({total_alerts})"): st.session_state.page = "تنبيهات"
    st.markdown('</div>', unsafe_allow_html=True)
with col4:
    st.markdown('<div class="btn-report">', unsafe_allow_html=True)
    if st.button("📄 التقارير"): st.session_state.page = "تقرير" # <-- ضفت زرار التقارير
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="btn-search">', unsafe_allow_html=True)
if st.button("🔍 بحث"): st.session_state.page = "بحث"
st.markdown('</div>', unsafe_allow_html=True)

# ========== صفحات التطبيق ==========
if st.session_state.page == "تقرير":
    st.subheader("📄 قسم التقارير")
    st.markdown("---")
    
    tab1, tab2, tab3 = st.tabs(["تقرير اجمالي", "تقرير حسب الحالة", "تقرير حسب النوع"])
    
    with tab1:
        st.write(f"**اجمالي القضايا:** {len(data['cases'])}")
        st.write(f"**القضايا المتداولة:** {len([c for c in data['cases'] if c.get('حالة')=='متداولة'])}")
        st.write(f"**القضايا المنتهية:** {len([c for c in data['cases'] if c.get('حالة')=='منتهية'])}")
        
    with tab2:
        if data['cases']:
            df = pd.DataFrame(data['cases'])
            st.dataframe(df.groupby('حالة').size().reset_index(name='العدد'))
        else:
            st.info("لا توجد قضايا")
            
    with tab3:
        if data['cases']:
            df = pd.DataFrame(data['cases'])
            st.dataframe(df.groupby('نوع').size().reset_index(name='العدد'))
        else:
            st.info("لا توجد قضايا")
    
    if st.button("⬅️ الرجوع للرئيسية"):
        st.session_state.page = "الرئيسية"
        st.rerun()
# ========== الصفحة الرئيسية =========
alerts = get_alert_cases()
total_alerts = len(alerts['sessions']) + len(alerts['appeals'])

col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="btn-add">', unsafe_allow_html=True)
    if st.button("اضافة قضية جديدة"): st.session_state.page = "اضافة" # شلت +
    st.markdown('</div>', unsafe_allow_html=True)
with col2:
    st.markdown('<div class="btn-list">', unsafe_allow_html=True)
    if st.button("عرض كل القضايا"): st.session_state.page = "عرض" # شلت 📋
    st.markdown('</div>', unsafe_allow_html=True)

col3, col4 = st.columns(2)
with col3:
    st.markdown('<div class="btn-alert">', unsafe_allow_html=True)
    if st.button(f"التنبيهات ({total_alerts})"): st.session_state.page = "تنبيهات" # شلت 🚨
    st.markdown('</div>', unsafe_allow_html=True)
with col4:
    st.markdown('<div class="btn-report">', unsafe_allow_html=True)
    if st.button("طباعة تقرير"): st.session_state.page = "تقرير"
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="btn-search">', unsafe_allow_html=True)
if st.button("بحث"): st.session_state.page = "بحث" # شلت 🔍
st.markdown('</div>', unsafe_allow_html=True)
# ========== الصفحة الرئيسية =========
alerts = get_alert_cases()
total_alerts = len(alerts['sessions']) + len(alerts['appeals'])

col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="btn-add">', unsafe_allow_html=True)
    if st.button("➕ اضافة قضية جديدة"): st.session_state.page = "اضافة"
    st.markdown('</div>', unsafe_allow_html=True)
with col2:
    st.markdown('<div class="btn-list">', unsafe_allow_html=True)
    if st.button("📋 عرض كل القضايا"): st.session_state.page = "عرض"
    st.markdown('</div>', unsafe_allow_html=True)

col3, col4 = st.columns(2)
with col3:
    st.markdown('<div class="btn-alert">', unsafe_allow_html=True)
    if st.button(f"🚨 التنبيهات ({total_alerts})"): st.session_state.page = "تنبيهات"
    st.markdown('</div>', unsafe_allow_html=True)
with col4:
    st.markdown('<div class="btn-report">', unsafe_allow_html=True)
    if st.button("📄 طباعة تقرير"): st.session_state.page = "تقرير"
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="btn-search">', unsafe_allow_html=True)
if st.button("🔍 بحث"): st.session_state.page = "بحث"
st.markdown('</div>', unsafe_allow_html=True)
# ================================================
# ========== الصفحة الرئيسية ==========
if st.session_state.page == "الرئيسية":
    st.markdown('<h2>الأقسام</h2>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="btn-add">', unsafe_allow_html=True)
        if st.button("➕ تسجيل القضايا", use_container_width=True, key="btn_add"): st.session_state.page = "تسجيل"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="btn-list">', unsafe_allow_html=True)
        if st.button("📋 الحصر العام", use_container_width=True, key="btn_list"): st.session_state.page = "الحصر"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown('<div class="btn-alert">', unsafe_allow_html=True)
        if st.button("🔴 مركز التنبيهات", use_container_width=True, key="btn_alert"): st.session_state.page = "تنبيهات"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="btn-report">', unsafe_allow_html=True)
        if st.button("📊 التقارير", use_container_width=True, key="btn_reports"): # زودت key هنا
            st.session_state.page = "تقارير"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="btn-lib">', unsafe_allow_html=True)
        if st.button("📚 المكتبة القانونية", use_container_width=True, key="btn_lib"):
            st.session_state.page = "المكتبة"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown('<div class="btn-arch">', unsafe_allow_html=True)
        if st.button("🗄️ الأرشيف", use_container_width=True, key="btn_arch"): st.session_state.page = "ارشيف"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown('<div class="btn-search">', unsafe_allow_html=True)
        if st.button("🔍 البحث عن دعوى", use_container_width=True, key="btn_search"): st.session_state.page = "بحث"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
# ================================================
# ========== نهاية الجزء الاول ==========
# =======================================
# ================================================
# =========== الجزء الثاني: تسجيل القضايا ============
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
                if data is None: data = {"cases": [], "archive": [], "library": [], "tasks": [], "users": []}
                new_case = {"id": len(data.get("cases", []))+1, "نوع": نوع, "محكمة_اسم": محكمة_اسم, "مأمورية": مأمورية, "رقم": رقم, "سنة": سنة, "دائرة": دائرة, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة), "الرول": الرول, "سبب": سبب, "ملاحظات": ملاحظات, "جلسات": [], "مستندات": [], "حالة": "متداولة", "مسندة_ل": مسندة_ل}
                if الرول or سبب: new_case["جلسات"].append({"تاريخ":str(تاريخ_جلسة),"الرول":الرول,"سبب":سبب,"ملاحظات":ملاحظات})
                data["cases"].append(new_case); save_data(data)
                st.success(f"✅ تم حفظ القضية رقم {رقم} لسنة {سنة}")
                st.session_state.page = "الحصر"; st.rerun()
# ================================
# ===============================================
# ======= الجزء الثالث: الحصر العام ============
# ================================================
elif st.session_state.page == "الحصر":
    data = load_data() # <-- اتغير
    st.session_state.data = data # <-- فضل زي ما هو
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📊 الحصر العام الخارجي</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()
    # ======= السطرين الجداد دول بس =======
    if st.session_state.get('open_from_search', False):
        st.session_state.open_from_search = False
        st.info("جاري فتح القضية من البحث...")
    # =======================
    # ======================================
    if not data or not data.get("cases", []):
        st.info("لا توجد قضايا مسجلة") # <-- زودت 4 مسافات بس هنا
    else:
        for i, case in enumerate(data["cases"]):
            if "id" not in case: case["id"] = i + 1
            if "مستندات" not in case: case["مستندات"] = []

        save_data(data)
        sorted_cases = sorted(data["cases"], key=lambda x: x.get("تاريخ_جلسة","9999-12-31"))
        total = len(sorted_cases)
        today = datetime.now().date()
        this_week = len([c for c in sorted_cases if c.get('تاريخ_جلسة') and datetime.strptime(c['تاريخ_جلسة'],'%Y-%m-%d').date() <= today + timedelta(days=7)])
        ended = len([c for c in sorted_cases if c.get('حالة') == 'منتهية'])
        
        st.markdown(f"<div style='background:#1E2A47; padding:20px; border-radius:15px; border:2px solid #D4AF37; text-align:center; margin-bottom:20px'>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1: st.markdown(f"<div style='font-size:28px; font-weight:900; color:#D4AF37'>📊 {total}</div><div style='font-size:18px; color:#FFF; font-weight:700'>اجمالي القضايا</div>", unsafe_allow_html=True)
        with col2: st.markdown(f"<div style='font-size:28px; font-weight:900; color:#4DA8DA'>📅 {this_week}</div><div style='font-size:18px; color:#FFF; font-weight:700'>جلسات هذا الاسبوع</div>", unsafe_allow_html=True)
        with col3: st.markdown(f"<div style='font-size:28px; font-weight:900; color:#FF5252'>🚫 {ended}</div><div style='font-size:18px; color:#FFF; font-weight:700'>عدد المنتهية</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        for idx, case in enumerate(sorted_cases, 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
            if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = f"{case.get('دائرة','')} عمال" if case.get('دائرة','') else "" # <-- مرة واحدة بس
            if دائرة_كاملة: محكمة_كاملة += f"<br>{دائرة_كاملة}"
            خصوم = f"<div style='background:#FFF3CD; padding:8px; border-radius:8px; color:#000; margin-bottom:5px; text-align:center'><b>المدعى:</b><br>{case.get('مدعي','')}</div><div style='background:#CFF4FC; padding:8px; border-radius:8px; color:#000; text-align:center'><b>المدعى عليه:</b><br>{case.get('مدعي_عليه','')}</div>"

            if case.get('حالة') == 'منتهية': row_class = "row-judgment"
            elif "الهيئة" in str(case.get('مدعي','')): row_class = "row-hey2a"
            else: row_class = "row1" if idx % 2 == 1 else "row2"

            st.markdown("<div class='table-container'>", unsafe_allow_html=True)
            table_html = f"<table class='case-table'><tr><th>م</th><th>الرقم والسنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>اخر جلسة</th><th>السبب</th><th>الحالة</th></tr><tr class='{row_class}'><td>{idx}</td><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td style='color:#FFD700; font-weight:900'>{case.get('تاريخ_جلسة','')}</td><td>{case.get('سبب','')}</td><td style='color:#4CAF50; font-weight:900'>{case.get('حالة','متداولة')}</td></tr></table></div>"
            st.markdown(table_html, unsafe_allow_html=True)

            c1, c2, c3 = st.columns([4,1,4])
            with c2:
                if st.button("فتح", key=f"open_{case['id']}", use_container_width=True): 
                    st.session_state.selected_case_id = case['id']; st.session_state.page = "تفاصيل"; st.rerun()

# =================================
# ================================================
# ============ الجزء الرابع: تفاصيل القضية ============
# ================================================
elif st.session_state.page == "تفاصيل":
    data = load_data()
    case = next((c for c in data["cases"] if c["id"] == st.session_state.selected_case_id), None)
    if not case: st.error("القضية غير موجودة"); st.session_state.page = "الحصر"; st.rerun()
    if 'جلسات' not in case: case['جلسات'] = []
    if 'مستندات' not in case: case['مستندات'] = []

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown(f"<h2 style='color:#D4AF37; text-align:center'>📄 تفاصيل القضية رقم {case.get('رقم')} لسنة {case.get('سنة')}</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للحصر العام", use_container_width=True): st.session_state.page = "الحصر"; st.rerun()

    # 1- بيانات القضية كروت 3 في سطر
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:15px'>1- بيانات القضية</div>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1: 
        st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>رقم القضية</div><div style='color:#FFF; font-weight:900; font-size:22px'>{case.get('رقم')}</div></div>", unsafe_allow_html=True)
    with col2: 
        st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>السنة</div><div style='color:#FFF; font-weight:900; font-size:22px'>{case.get('سنة')}</div></div>", unsafe_allow_html=True)
    with col3: 
        دائرة_نص = f"{case.get('دائرة')} عمال" if case.get('دائرة') else "" # <-- مرة واحدة بس
        st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>الدائرة</div><div style='color:#FFF; font-weight:900; font-size:18px'>{دائرة_نص}</div></div>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1: 
        st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>النوع</div><div style='color:#FFF; font-weight:900; font-size:18px'>{case.get('نوع')}</div></div>", unsafe_allow_html=True)
    with col2: 
        محكمة_كاملة = f"{case.get('محكمة_اسم')}"
        if case.get('مأمورية'): محكمة_كاملة += f" - مأمورية {case.get('مأمورية')}"
        st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>المحكمة</div><div style='color:#FFF; font-weight:700; font-size:14px'>{محكمة_كاملة}</div></div>", unsafe_allow_html=True)
    with col3: 
        st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>الحالة</div><div style='color:#4CAF50; font-weight:900; font-size:18px'>{case.get('حالة')}</div></div>", unsafe_allow_html=True)
    
    st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>الموضوع</div><div style='color:#FFF; font-weight:700; font-size:16px'>{case.get('موضوع')}</div></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # 2- بيانات الخصوم كارت
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>2- بيانات الخصوم</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1: st.markdown(f"<div style='background:#FFF3CD; padding:10px; border-radius:10px; color:#000; text-align:center'><b>المدعى:</b><br>{case.get('مدعي')}</div>", unsafe_allow_html=True)
    with col2: st.markdown(f"<div style='background:#CFF4FC; padding:10px; border-radius:10px; color:#000; text-align:center'><b>المدعى عليه:</b><br>{case.get('مدعي_عليه')}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # 3- متابعة الجلسات
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>3- متابعة الجلسات</div>", unsafe_allow_html=True)
    if case.get("جلسات"):
        html = "<table style='width:100%; border:2px solid #D4AF37; background:#0A1428; border-radius:12px'><tr style='background:#D4AF37; color:#000'><th>م</th><th>التاريخ</th><th>الرول</th><th>السبب</th><th>ملاحظات</th></tr>"
        for i, ج in enumerate(case["جلسات"], 1):
            لون = "#1E2A47" if i % 2 == 0 else "#142038"
            html += f"<tr style='background:{لون}; color:#FFF'><td>{i}</td><td>{ج.get('تاريخ')}</td><td>{ج.get('الرول')}</td><td>{ج.get('سبب')}</td><td>{ج.get('ملاحظات')}</td></tr>"
        html += "</table>"; st.markdown(html, unsafe_allow_html=True)
    else: st.info("لا توجد جلسات مسجلة")
    
    with st.expander("➕ اضافة جلسة جديدة"): # <-- ظهرتها
        with st.form("add_session"):
            تاريخ_جديد = st.date_input("تاريخ الجلسة", value=datetime.now())
            رول_جديد = st.text_input("الرول"); سبب_جديد = st.text_input("سبب الجلسة"); ملاحظات_جديدة = st.text_area("ملاحظات")
            if st.form_submit_button("حفظ الجلسة"):
                case["جلسات"].append({"تاريخ":str(تاريخ_جديد),"الرول":رول_جديد,"سبب":سبب_جديد,"ملاحظات":ملاحظات_جديدة})
                case["تاريخ_جلسة"] = str(تاريخ_جديد); save_data(data); st.success("تم اضافة الجلسة"); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    # 4- المستندات
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>4- المستندات</div>", unsafe_allow_html=True)
    with st.form("upload_form"):
        نوع_المستند = st.selectbox("نوع المستند", ANWA3_MOSTANDAT)
        uploaded_file = st.file_uploader("اختر الملف")
        if st.form_submit_button("رفع المستند"):
            if uploaded_file:
                file_path = os.path.join(UPLOAD_FOLDER, f"{case['id']}_{uploaded_file.name}")
    st.markdown("</div>", unsafe_allow_html=True)

    # 5- جلسة الحكم
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #FF5252; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#FF5252; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>5- جلسة الحكم</div>", unsafe_allow_html=True)
    
    if case.get('حالة') != 'منتهية':
        with st.form("judgment_form"):
            st.markdown("<div style='background:#142038; padding:10px; border-radius:10px; margin-bottom:10px'>", unsafe_allow_html=True)
            st.markdown("<label style='color:#FFD700; font-weight:900; font-size:16px'>1- تاريخ الجلسة</label>", unsafe_allow_html=True)
            تاريخ_حكم = st.date_input("تاريخ الجلسة", value=datetime.now().date(), label_visibility="collapsed")
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<div style='background:#142038; padding:10px; border-radius:10px; margin-bottom:10px'>", unsafe_allow_html=True)
            st.markdown("<label style='color:#FFD700; font-weight:900; font-size:16px'>2- منطوق الحكم</label>", unsafe_allow_html=True)
            منطوق_الحكم = st.text_area("منطوق الحكم", height=150, placeholder="اكتب منطوق الحكم هنا...", label_visibility="collapsed")
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<div style='background:#142038; padding:10px; border-radius:10px; margin-bottom:10px'>", unsafe_allow_html=True)
            st.markdown("<label style='color:#FFD700; font-weight:900; font-size:16px'>3- مسندة لـ</label>", unsafe_allow_html=True)
            مسندة_ل = st.selectbox("مسندة لـ", ["الصالح", "الضد"], label_visibility="collapsed")
            st.markdown("</div>", unsafe_allow_html=True)
            
            if st.form_submit_button("💾 حفظ الحكم", use_container_width=True, type="primary"):
                if not منطوق_الحكم:
                    st.error("❌ لازم تكتب منطوق الحكم")
                else:
                    case['حالة'] = 'منتهية'
                    case['تاريخ_الحكم'] = str(تاريخ_حكم)
                    case['منطوق_الحكم'] = منطوق_الحكم
                    case['مسندة_ل_الحكم'] = مسندة_ل
                    case['جلسات'].append({'تاريخ':str(تاريخ_حكم),'الرول':'-','سبب':f'الحكم - مسندة لـ {مسندة_ل}','ملاحظات':منطوق_الحكم})
                    case['تاريخ_جلسة'] = str(تاريخ_حكم)
                    case['سبب'] = f'الحكم - مسندة لـ {مسندة_ل}'
                    save_data(data)
                    st.success(f"✅ تم حفظ الحكم واغلاق القضية. تم نقلها للارشيف")
                    st.session_state.page = "الأرشيف"
                    st.rerun()
    else:
        st.success(f"✅ تم الحكم بتاريخ: {case.get('تاريخ_الحكم')}")
        st.info(f"**مسندة لـ:** {case.get('مسندة_ل_الحكم')}")
        st.warning(f"**المنطوق:** {case.get('منطوق_الحكم')}")
        
        if st.button("↩️ ارجاع القضية للتداول", use_container_width=True):
            case['حالة'] = 'متداولة'
            case['تاريخ_الحكم'] = ""
            case['منطوق_الحكم'] = ""
            case['مسندة_ل_الحكم'] = ""
            save_data(data)
            st.session_state.page = "الحصر"
            st.rerun()
            
    st.markdown("</div>", unsafe_allow_html=True)
    # ================================================
# ================================================
# ============ الجزء الخامس: الأرشيف ============
# ================================================
elif st.session_state.page == "الأرشيف":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FF5252; text-align:center'>📁 أرشيف الأحكام النهائية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()

    # ======= ده الجديد عشان يفتح القضية من البحث =======
    if st.session_state.get('selected_case_id'):
        case_id = st.session_state.selected_case_id
        st.session_state.selected_case_id = None # امسحها
        st.session_state.open_from_search = False
        
        case = next((c for c in data["cases"] if c['id'] == case_id), None)
        if case:
            st.success(f"تم فتح القضية رقم {case.get('رقم','')} لسنة {case.get('سنة','')}")
            st.markdown("---")
            st.markdown(f"### 📂 تفاصيل القضية")
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**المدعي:** {case.get('مدعي','')}")
                st.markdown(f"**نوع الدعوى:** {case.get('نوع_الدعوى','')}")
                st.markdown(f"**المحكمة:** {case.get('محكمة_اسم','')}")
                if case.get('مأمورية'): st.markdown(f"**المأمورية:** {case.get('مأمورية','')}")
                if case.get('دائرة'): st.markdown(f"**الدائرة:** {case.get('دائرة','')}")
            with col2:
                st.markdown(f"**المدعى عليه:** {case.get('مدعي_عليه','')}")
                st.markdown(f"**الموضوع:** {case.get('موضوع','')}")
                st.markdown(f"**اخر جلسة:** {case.get('تاريخ_جلسة','')}")
                st.markdown(f"**الحالة:** {case.get('حالة','')}")
            
            st.markdown("---")
            if st.button("⬅️ الرجوع لجدول الارشيف", use_container_width=True):
                st.rerun()
            st.stop() # مهم عشان ميعرضش الجدول تحت
    # =====================================
    # بحث
    search_query = st.text_input("🔍 ابحث برقم القضية او الخصوم او رقم الطعن", key="search_archive")
    
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()

    # نجيب القضايا اللي اتحكم فيها بس للصالح او للضد
    ended_cases = [c for c in data["cases"] if c.get('حالة') == 'منتهية' and c.get('مسندة_ل_الحكم') in ['الصالح', 'الضد']]
    
    # فلتر البحث
    if search_query:
        ended_cases = [c for c in ended_cases if 
                       search_query in str(c.get('رقم','')) or 
                       search_query in str(c.get('مدعي','')) or
                       search_query in str(c.get('مدعي_عليه','')) or
                       search_query in str(c.get('رقم_الطعن',''))]

    if not ended_cases:
        st.info("لا توجد أحكام نهائية في الارشيف")
    else:
        # ترتيب من الاحدث للاقدم
        sorted_ended = sorted(ended_cases, key=lambda x: x.get("تاريخ_الحكم","0000-00-00"), reverse=True)
        total_ended = len(sorted_ended)
        
        st.markdown(f"<div style='background:#1E2A47; padding:20px; border-radius:15px; border:2px solid #FF5252; text-align:center; margin-bottom:20px'>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:28px; font-weight:900; color:#FF5252'>📁 {total_ended}</div><div style='font-size:18px; color:#FFF; font-weight:700'>عدد الأحكام النهائية</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        for idx, case in enumerate(sorted_ended, 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            if case.get('رقم_الطعن'): رقم_كامل += f"<br><b style='color:#FFD700'>طعن رقم:</b> {case.get('رقم_الطعن')}"
            
            محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
            if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = f"{case.get('دائرة','')} عمال" if case.get('دائرة','') else ""
            if دائرة_كاملة: محكمة_كاملة += f"<br>{دائرة_كاملة}"
            
            خصوم = f"<div style='background:#FFF3CD; padding:8px; border-radius:8px; color:#000; margin-bottom:5px; text-align:center'><b>المدعى:</b><br>{case.get('مدعي','')}</div><div style='background:#CFF4FC; padding:8px; border-radius:8px; color:#000; text-align:center'><b>المدعى عليه:</b><br>{case.get('مدعي_عليه','')}</div>"
            
            # ملخص بيانات الحكم
            مسندة = case.get('مسندة_ل_الحكم','')
            لون_مسندة = "#4CAF50" if مسندة == "الصالح" else "#FF5252"
            بيانات_الحكم = f"<div style='background:#142038; padding:8px; border-radius:8px; border:2px solid {لون_مسندة}; text-align:center'><b style='color:{لون_مسندة}; font-size:16px'>تاريخ الحكم:</b><br>{case.get('تاريخ_الحكم','')}<br><b style='color:{لون_مسندة}'>مسندة لـ:</b> {مسندة}<br><b style='color:{لون_مسندة}'>المنطوق:</b><br>{case.get('منطوق_الحكم','')}</div>"
            
            # بيانات الحفظ لو موجودة
            بيانات_الحفظ = f"<div style='background:#2A3A5F; padding:8px; border-radius:8px; color:#FFF; text-align:center'><b>سبب الحفظ:</b><br>{case.get('سبب_الحفظ','-')}</div>" if case.get('سبب_الحفظ') else "-"
            
            # بيانات اعادة التداول لو موجودة
            بيانات_العودة = f"<div style='background:#5F2A2A; padding:8px; border-radius:8px; color:#FFF; text-align:center'><b>سبب العودة:</b><br>{case.get('سبب_العودة','-')}</div>" if case.get('سبب_العودة') else "-"

            st.markdown("<div class='table-container'>", unsafe_allow_html=True)
            st.markdown(f"<table class='case-table'><tr><th>م</th><th>الرقم والسنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>بيانات الحكم</th><th>الحفظ</th><th>عودة للتداول</th></tr><tr class='row-judgment'><td>{idx}</td><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td>{بيانات_الحكم}</td>", unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            with col1:
                سبب_الحفظ = st.text_input("سبب الحفظ", key=f"save_reason_{case['id']}", placeholder="اكتب سبب الحفظ...")
                if st.button("💾 حفظ", key=f"save_btn_{case['id']}", use_container_width=True):
                    if سبب_الحفظ:
                        case['سبب_الحفظ'] = سبب_الحفظ
                        save_data(data)
                        st.success("تم حفظ سبب الحفظ")
                        st.rerun()
                    else:
                        st.error("اكتب سبب الحفظ الاول")
                st.markdown(بيانات_الحفظ, unsafe_allow_html=True)
                
            with col2:
                سبب_العودة = st.text_input("سبب العودة للتداول", key=f"return_reason_{case['id']}", placeholder="ليه بترجعها؟")
                if st.button("↩️ عودة للتداول", key=f"return_{case['id']}", use_container_width=True):
                    if سبب_العودة:
                        case['حالة'] = 'متداولة'
                        case['سبب_العودة'] = سبب_العودة
                        case['تاريخ_الحكم'] = ""
                        case['منطوق_الحكم'] = ""
                        case['مسندة_ل_الحكم'] = ""
                        save_data(data)
                        st.success("تم ارجاع القضية للتداول")
                        st.session_state.page = "الحصر"
                        st.rerun()
                    else:
                        st.error("لازم تكتب سبب العودة")
                st.markdown(بيانات_العودة, unsafe_allow_html=True)
            
            st.markdown("</tr></table></div>", unsafe_allow_html=True)
            # ================================================
# ============ الجزء السادس: البحث ============
# ================================================
elif st.session_state.page == "بحث":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#4DA8DA; text-align:center'>🔍 البحث عن دعوى</h2>", unsafe_allow_html=True)

    if st.button("⬅️ العودة للرئيسية", use_container_width=True, key="back_from_search"):
        st.session_state.page = "الرئيسية"
        st.rerun()

    st.markdown("<div style='background:#1E2A47; padding:20px; border-radius:15px; border:2px solid #4DA8DA; margin-bottom:20px'>", unsafe_allow_html=True)
    search_type = st.selectbox("نوع البحث", [
        "اسم المدعي او المدعى عليه",
        "رقم الدعوى وسنتها - اول درجة",
        "رقم الاستئناف وسنته - ثاني درجة",
        "رقم الطعن بالنقض وسنته",
        "رقم الدعوى وسنتها - المحكمة الادارية",
        "رقم الدعوى وسنتها - محكمة القضاء الاداري",
        "رقم الطعن وسنته - القضاء الاداري بهيئة استئنافية",
        "رقم الدعوى وسنتها - المحكمة الادارية العليا"
    ], key="search_type_select")

    search_value = st.text_input("اكتب كلمة البحث", placeholder="للاسم: اكتب الاسم | للرقم: اكتب الرقم والسنة مثال 1234 2024", key="search_input")
    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🔍 بحث", use_container_width=True, type="primary", key="do_search"):
        if not search_value.strip():
            st.error("اكتب حاجة تبحث بيها")
        else:
            results = []
            parts = search_value.split()
            رقم_بحث = parts[0] if len(parts) > 0 else ""
            سنة_بحث = parts[1] if len(parts) > 1 else ""

            for case in data["cases"]:
                match = False
                نوع_الدعوى = str(case.get('نوع_الدعوى', '')).lower()

                if search_type == "اسم المدعي او المدعى عليه":
                    if search_value in str(case.get('مدعي','')) or search_value in str(case.get('مدعي_عليه','')):
                        match = True

                elif "اول درجة" in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and ("دعوى" in نوع_الدعوى or "عمال" in نوع_الدعوى):
                        match = True

                elif "استئناف" in search_type and "استئنافية" not in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and "استئناف" in نوع_الدعوى:
                        match = True

                elif "الطعن بالنقض" in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and "طعن" in نوع_الدعوى and "اداري" not in نوع_الدعوى:
                        match = True

                elif "المحكمة الادارية" in search_type and "العليا" not in search_type and "استئنافية" not in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and "ادارية" in نوع_الدعوى and "عليا" not in نوع_الدعوى:
                        match = True

                elif "القضاء الاداري" in search_type and "استئنافية" not in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and "قضاء اداري" in نوع_الدعوى:
                        match = True

                elif "استئنافية" in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and "استئنافية" in نوع_الدعوى:
                        match = True

                elif "الادارية العليا" in search_type:
                    if رقم_بحث == str(case.get('رقم','')) and سنة_بحث == str(case.get('سنة','')) and "ادارية عليا" in نوع_الدعوى:
                        match = True

                if match:
                    results.append(case)

            if not results:
                st.warning("لم يتم العثور على نتائج")
            else:
                st.success(f"تم العثور على {len(results)} نتيجة")
                for idx, case in enumerate(results, 1):
                    رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
                    محكمة_كاملة = f"{case.get('نوع_الدعوى','')} - {case.get('محكمة_اسم','')}"
                    if case.get('مأمورية',''):
                        محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
                    if case.get('دائرة',''):
                        محكمة_كاملة += f"<br>دائرة {case.get('دائرة','')}"

                    خصوم = f"<div style='background:#FFF3CD; padding:8px; border-radius:8px; color:#000; margin-bottom:5px; text-align:center'><b>المدعى:</b><br>{case.get('مدعي','')}</div><div style='background:#CFF4FC; padding:8px; border-radius:8px; color:#000; text-align:center'><b>المدعى عليه:</b><br>{case.get('مدعي_عليه','')}</div>"

                    if case.get('حالة') == 'منتهية':
                        row_class = "row-judgment"
                        حالة_لون = "#FF5252"
                        مكان = "📁 الأرشيف"
                        الصفحة_المطلوبة = "الأرشيف" # <-- اتعدل هنا بالالف واللام
                    else:
                        row_class = "row1" if idx % 2 == 1 else "row2"
                        حالة_لون = "#4CAF50"
                        مكان = "📋 الحصر العام"
                        الصفحة_المطلوبة = "الحصر"

                    st.markdown("<div class='table-container'>", unsafe_allow_html=True)
                    st.markdown(f"""
                    <table class='case-table'>
                    <tr><th>م</th><th>الرقم والسنة</th><th>نوع الدعوى والمحكمة</th><th>الخصوم</th><th>الموضوع</th><th>اخر جلسة</th><th>الحالة</th><th>المكان</th></tr>
                    <tr class='{row_class}'>
                        <td>{idx}</td>
                        <td>{رقم_كامل}</td>
                        <td>{محكمة_كاملة}</td>
                        <td>{خصوم}</td>
                        <td>{case.get('موضوع','')}</td>
                        <td style='color:#FFD700; font-weight:900'>{case.get('تاريخ_جلسة','-')}</td>
                        <td style='color:{حالة_لون}; font-weight:900'>{case.get('حالة','متداولة')}</td>
                        <td style='color:#4DA8DA; font-weight:900'>{مكان}</td>
                    </tr>
                    </table>
                    """, unsafe_allow_html=True)

                    c1, c2, c3 = st.columns([4,1,4])
                    with c2:
                        if st.button(f"📂 فتح في {مكان}", key=f"open_smart_btn_{case['id']}", use_container_width=True):
                            st.session_state.selected_case_id = case['id']
                            st.session_state.open_from_search = True # العلامة عشان الصفحة التانية تعرف
                            st.session_state.page = الصفحة_المطلوبة
                            st.rerun()
                    st.markdown("</div>", unsafe_allow_html=True)
                    # ========== الجزء السابع: مركز التنبيهات ==========
elif st.session_state.page == "تنبيهات":
    st.markdown("<h1 style='text-align: center; color: #C9A961;'>🔔 مركز التنبيهات</h1>", unsafe_allow_html=True)
    
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): 
        st.session_state.page = "الرئيسية"
        st.rerun()
    
    # ====== 1. تسجيل الايميل ======
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("<div class='card-title'>📧 تسجيل الايميل لاستلام التنبيهات</div>", unsafe_allow_html=True)
        user_email = st.text_input("ادخل ايميلك", placeholder="example@gmail.com", key="email_alert")
        if st.button("تسجيل الايميل", type="primary", use_container_width=True):
            if user_email:
                st.session_state['saved_email'] = user_email
                st.success(f"✅ تم حفظ الايميل {user_email}. هيجيلك تنبيهات قريب")
            else:
                st.warning("دخل الايميل الاول")
    
    alerts = get_alert_cases()
    st.markdown(f"<h3 style='text-align:center; color:#FFFFFF;'>📅 تاريخ اليوم: {datetime.now().strftime('%Y-%m-%d')}</h3>", unsafe_allow_html=True)
    
    # ====== 2. الجلسات خلال 7 ايام ======
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='text-align:center; color:#C9A961;'>⚖️ جلسات خلال 7 ايام القادمة</h2>", unsafe_allow_html=True)
    
    if alerts["sessions"]:
        for case in alerts["sessions"]:
            with st.container(border=True):
                # نفس جدول الحصر
                رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
                محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
                if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
                دائرة_كاملة = f"{case.get('دائرة','')} عمال" if case.get('دائرة','') else ""
                محكمة_كاملة += f"<br>{دائرة_كاملة}"
                خصوم = f"{case.get('مدعي','')}<br>ضد<br>{case.get('مدعي_عليه','')}"
                
                st.markdown(f"<h4 style='color:#FFD700; text-align:center;'>⚠️ فاضل {case['days_left']} يوم على الجلسة</h4>", unsafe_allow_html=True)
                
                table_html = f"<div class='table-container'><table class='case-table'><tr><th>الرقم والسنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>تاريخ الجلسة</th></tr><tr class='row1'><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td>{case.get('تاريخ_جلسة','')}</td></tr></table></div>"
                st.markdown(table_html, unsafe_allow_html=True)
                
                if st.button("📂 فتح القضية", key=f"open_alert_{case['id']}", use_container_width=True):
                    st.session_state.selected_case_id = case['id']
                    st.session_state.page = "تفاصيل"
                    st.rerun()
    else:
        st.success("✅ مفيش جلسات خلال 7 ايام")
    
    # ====== 3. الطعون خلال 15 يوم ======
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='text-align:center; color:#C9A961;'>📄 طعون خلال 15 يوم</h2>", unsafe_allow_html=True)
    
    if alerts["appeals"]:
        for case in alerts["appeals"]:
            with st.container(border=True):
                رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
                محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
                if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
                دائرة_كاملة = f"{case.get('دائرة','')} عمال" if case.get('دائرة','') else ""
                محكمة_كاملة += f"<br>{دائرة_كاملة}"
                خصوم = f"{case.get('مدعي','')}<br>ضد<br>{case.get('مدعي_عليه','')}"
                
                st.markdown(f"<h4 style='color:#FF4500; text-align:center;'>⏰ فاضل {case['days_left_appeal']} يوم على اخر ميعاد للطعن</h4>", unsafe_allow_html=True)
                
                table_html = f"<div class='table-container'><table class='case-table'><tr><th>الرقم والسنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>تاريخ الحكم</th></tr><tr class='row-judgment'><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td>{case.get('تاريخ_الحكم','')}</td></tr></table></div>"
                st.markdown(table_html, unsafe_allow_html=True)
                
                if st.button("📂 فتح القضية", key=f"open_appeal_{case['id']}", use_container_width=True):
                    st.session_state.selected_case_id = case['id']
                    st.session_state.page = "تفاصيل"
                    st.rerun()
    else:
        st.success("✅ مفيش طعون قريبة")
        # ========= صفحة المكتبة القانونية =========
elif st.session_state.page == "المكتبة":
    st.markdown('<h1 style="text-align: center; color: #FFD700;">المكتبة 📚<br>القانونية</h1>', unsafe_allow_html=True)
    
    # 1. البحث العام في كل المواد
    st.markdown("### 🔍 البحث عن أي مادة قانونية")
    search_query = st.text_input("ابحث باسم الموضوع")
    
    col1, col2 = st.columns(2)
    with col1:
        search_number = st.text_input("رقم القانون / القرار / التعليمات")
    with col2:
        search_year = st.text_input("السنة")
    
    if st.button("بحث", use_container_width=True, type="primary"):
        st.session_state.search_filters = {"q": search_query, "num": search_number, "year": search_year}
        st.rerun()

    st.divider()

    # 2. ال 22 قسم بالالوان
    LIBRARY_SECTIONS = {
        "القوانين": "#FF6B6B", "القرارات الوزارية": "#4ECDC4", "قرارات الهيئة": "#45B7D1",
        "المنشورات الوزارية": "#96CEB4", "منشورات الهيئة": "#FFEAA7", "الكتب الدورية": "#DDA0DD",
        "تعليمات الهيئة": "#98D8C8", "رسائل الهيئة": "#F7DC6F", "المرصد الفني": "#BB8FCE",
        "فتاوى لجنة الشئون القانونية بالوزارة": "#85C1E2", "فتاوى الادارة المركزية للشئون القانونية": "#F8B500",
        "احكام المحكمة الدستورية العليا": "#E74C3C", "احكام محكمة النقض": "#3498DB",
        "احكام المحكمة الإدارية العليا": "#2ECC71", "احكام المحاكم الاستئنافية": "#9B59B6",
        "احكام محاكم القضاء الإدارى": "#1ABC9C", "احكام المحاكم الابتدائية": "#E67E22",
        "احكام المحكمة الإدارية": "#34495E", "منشورات القضاء العادى": "#16A085",
        "منشورات مجلس الدولة": "#8E44AD", "فتاوى الجمعية العمومية": "#27AE60",
        "صحف طعون": "#C0392B", "صحف استئنافات": "#2980B9", "صحف دعاوى": "#8E44AD",
        "مذكرات دفاع": "#D35400", "أخرى": "#7F8C8D"
    }

    # 3. عرض الاقسام كأزرار ملونة
    st.markdown("### 📁 الاقسام")
    cols = st.columns(4)
    for i, (section, color) in enumerate(LIBRARY_SECTIONS.items()):
        with cols[i % 4]:
            if st.button(f"{section}", key=f"sec_{section}", use_container_width=True, 
                        help=f"عرض {section}"):
                st.session_state.selected_section = section
                st.rerun()
            st.markdown(f'<div style="background:{color};height:5px;border-radius:5px;"></div>', unsafe_allow_html=True)

    # 4. لو اختار قسم او بحث
    if "selected_section" in st.session_state or "search_filters" in st.session_state:
        st.divider()
        
        library_data = st.session_state.data.get("library", [])
        
        # فلترة
        if "selected_section" in st.session_state:
            sec = st.session_state.selected_section
            st.subheader(f"📂 {sec}")
            files = [f for f in library_data if f.get("section") == sec]
        else:
            sec = "نتائج البحث"
            st.subheader("🔍 نتائج البحث")
            f = st.session_state.search_filters
            files = [item for item in library_data if 
                     f["q"].lower() in item.get("name","").lower() and
                     f["num"] in item.get("number","") and
                     f["year"] in item.get("year","")]
        
        # زر الاضافة
        if st.button("➕ اضافة مادة قانونية", key="add_doc", type="secondary"):
            st.session_state.show_upload = True

        # فورم الاضافة
        if st.session_state.get("show_upload", False):
            with st.form("form_add_doc"):
                section_select = st.selectbox("اختر القسم", list(LIBRARY_SECTIONS.keys()))
                doc_name = st.text_input("بيان المستند - الاسم اللي هيظهر في المكتبة")
                doc_number = st.text_input("الرقم")
                doc_year = st.text_input("السنة")
                doc_link = st.text_input("رابط المستند")
                
                if st.form_submit_button("💾 حفظ بصفة دائمة"):
                    new_doc = {
                        "id": secrets.token_hex(6),
                        "name": doc_name, 
                        "section": section_select,
                        "number": doc_number,
                        "year": doc_year,
                        "link": doc_link
                    }
                    st.session_state.data.setdefault("library", []).append(new_doc)
                    save_data(st.session_state.data)
                    st.success("تم الحفظ")
                    st.session_state.show_upload = False
                    st.rerun()

        st.divider()
        # 5. عرض المستندات مع 4 ازرار
        if files:
            for doc in files:
                color = LIBRARY_SECTIONS.get(doc.get("section"), "#7F8C8D")
                st.markdown(f'<div style="border-left:5px solid {color}; padding:10px; margin:5px 0; background:#1e1e1e;">', unsafe_allow_html=True)
                st.write(f"**{doc.get('name')}**")
                st.caption(f"رقم: {doc.get('number','-')} | سنة: {doc.get('year','-')} | القسم: {doc.get('section')}")
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    if st.button("📖 فتح", key=f"open_{doc['id']}", use_container_width=True):
                        st.info(f"الرابط: {doc.get('link')}")
                with col2:
                    if st.button("⬇️ تحميل", key=f"dl_{doc['id']}", use_container_width=True):
                        st.info("التحميل متاح للعضو المشترك")
                with col3:
                    if st.button("✏️ تعديل", key=f"edit_{doc['id']}", use_container_width=True):
                        st.warning("وظيفة التعديل قريبا")
                with col4:
                    if st.button("🗑️ حذف", key=f"del_{doc['id']}", use_container_width=True):
                        st.session_state.data["library"] = [d for d in st.session_state.data["library"] if d["id"] != doc["id"]]
                        save_data(st.session_state.data)
                        st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("مفيش مستندات. دوس ➕ اضافة مادة قانونية")

    st.divider()
    # زر العودة
    if st.button("⬅️ العودة للصفحة الرئيسية", use_container_width=True):
        st.session_state.page = "الرئيسية"
        for k in ["selected_section", "show_upload", "search_filters"]:
            st.session_state.pop(k, None)
        st.rerun()
        # ================================
# ================================================
# ============ الجزء الثامن: التقارير ============
# ================================================
elif st.session_state.page == "تقارير":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#D4AF37; text-align:center; font-family:Cairo'>📑 مركز التقارير الحكومية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()

    tab1, tab2, tab3, tab4 = st.tabs(["📊 بيان الدعاوى المتداولة", "⚖️ بيان الاحكام", "👤 بيان العضو السنوي", "📈 الإحصائيات"])

    # ===== دالة الهيدر الرسمي =====
    def report_header(region, title):
        st.markdown(f"""
        <div style='text-align:center; color:#D4AF37; border:4px double #D4AF37; padding:20px; background: linear-gradient(135deg, #0A1428 0%, #1E2A47 100%); border-radius:15px; margin-bottom:20px; font-family:Cairo;'>
        <h2 style='margin:5px'>الهيئة القومية للتأمين الاجتماعى</h2>
        <h3 style='margin:5px'>الإدارة المركزية للإدارات القانونية</h3>
        <h3 style='margin:5px'>الإدارة العامة للقضايا</h3>
        <h3 style='margin:5px'>ديوان عام {region}</h3>
        <hr style='border-color:#D4AF37'>
        <h3 style='margin:10px'>{title}</h3>
        </div>
        """, unsafe_allow_html=True)

    # ===== دالة تصدير PDF =====
    def generate_pdf_report(cases, report_type, region, from_date, to_date, lawyer):
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        
        elements = []
        style_ar = ParagraphStyle('Arabic', fontName='Cairo', fontSize=10, alignment=2, leading=16)
        style_header = ParagraphStyle('Header', fontName='Cairo', fontSize=14, alignment=1, textColor=colors.HexColor('#D4AF37'))

        elements.append(Paragraph(arabic(f"الهيئة القومية للتأمين الاجتماعي"), style_header))
        elements.append(Paragraph(arabic(f"{report_type}"), style_header))
        elements.append(Paragraph(arabic(f"ديوان عام: {region} | الفترة: {from_date} الى {to_date}"), style_ar))
        elements.append(Paragraph(arabic(f"المحامي: {lawyer}"), style_ar))
        elements.append(Spacer(1, 12))

        table_data = []
        if "العضو" in report_type:
            headers = [arabic(h) for h in ['م', 'رقم', 'سنة', 'حالة', 'موضوع', 'اخر اجراء']]
            table_data.append(headers)
            for i, c in enumerate(cases, 1):
                اجراء = c.get('تاريخ_جلسة','') if c.get('حالة')=='متداولة' else c.get('تاريخ_الحكم','')
                row = [arabic(str(i)), arabic(c.get('رقم','')), arabic(c.get('سنة','')), arabic(c.get('حالة','')), arabic(c.get('موضوع','')), arabic(اجراء)]
                table_data.append(row)
        elif "الاحكام" in report_type:
            headers = [arabic(h) for h in ['م', 'رقم', 'سنة', 'تاريخ الحكم', 'المنطوق', 'النتيجة']]
            table_data.append(headers)
            for i, c in enumerate(cases, 1):
                row = [arabic(str(i)), arabic(c.get('رقم','')), arabic(c.get('سنة','')), arabic(c.get('تاريخ_الحكم','')), arabic(c.get('منطوق_الحكم','')), arabic(c.get('مسندة_ل_الحكم',''))]
                table_data.append(row)
        else: # المتداولة
            headers = [arabic(h) for h in ['م', 'رقم', 'سنة', 'دائرة', 'موضوع', 'اخر جلسة']]
            table_data.append(headers)
            for i, c in enumerate(cases, 1):
                row = [arabic(str(i)), arabic(c.get('رقم','')), arabic(c.get('سنة','')), arabic(c.get('دائرة','')), arabic(c.get('موضوع','')), arabic(c.get('تاريخ_جلسة',''))]
                table_data.append(row)

        t = Table(table_data, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#D4AF37')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,-1), 'Cairo'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('GRID', (0,0), (-1,-1), 1, colors.black)
        ]))
        elements.append(t)
        elements.append(Spacer(1, 30))
        elements.append(Paragraph(arabic(f"تحرر في: {datetime.now().strftime('%Y-%m-%d')}"), style_ar))
        elements.append(Paragraph(arabic("عضو الادارة..................    مدير الإدارة.................."), style_ar))
        
        doc.build(elements)
        buffer.seek(0)
        return buffer

    # ========= تبويب 1: الدعاوى المتداولة =========
    with tab1:
        st.markdown("<div style='background:#1E2A47; padding:20px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px; font-family:Cairo'>", unsafe_allow_html=True)
        region = st.text_input("ديوان عام منطقة", key="region1")
        col1, col2, col3 = st.columns(3)
        with col1: from_date = st.date_input("من الفترة", key="from1")
        with col2: to_date = st.date_input("حتى الفترة", key="to1")
        with col3: lawyer = st.text_input("طرف الاستاذ/ المحامي", key="lawyer1")
        topic = st.text_input("موضوع الدعوى للفلترة - اتركها فاضية لعرض الكل", key="topic1")
        st.markdown("</div>", unsafe_allow_html=True)

        if st.button("🔍 عرض بيان الدعاوى المتداولة", use_container_width=True, type="primary", key="show1"):
            cases = [c for c in data["cases"] if c.get('حالة') == 'متداولة']
            if cases: cases = [c for c in cases if c.get('تاريخ_جلسة') and from_date <= datetime.strptime(c['تاريخ_جلسة'], '%Y-%m-%d').date() <= to_date]
            if topic: cases = [c for c in cases if topic in str(c.get('موضوع',''))]
            cases = sorted(cases, key=lambda x: x.get("تاريخ_جلسة","9999-12-31"), reverse=True)

            report_header(region, f"بيان بالدعاوى المتداولة خلال الفترة من {from_date} حتى {to_date} طرف الاستاذ/ {lawyer} المحامي")
            
            if not cases: st.warning("لا توجد دعاوى متداولة في الفترة المحددة")
            else:
                html = "<table class='case-table'><tr><th>م</th><th>رقم القضية</th><th>السنة</th><th>الدائرة والنوع</th><th>المحكمة</th><th>الخصوم</th><th>الموضوع</th><th>اخر جلسة</th></tr>"
                for i, c in enumerate(cases, 1):
                    محكمة = f"{c.get('محكمة_اسم','')}"
                    if c.get('مأمورية'): محكمة += f"<br>مأمورية {c.get('مأمورية')}"
                    دائرة = f"{c.get('دائرة','')} {c.get('نوع','')}"
                    خصوم = f"<div style='background:#FFF3CD; padding:5px; border-radius:5px; color:#000'><b>المدعى:</b> {c.get('مدعي','')}</div><div style='background:#CFF4FC; padding:5px; border-radius:5px; color:#000'><b>المدعى عليه:</b> {c.get('مدعي_عليه','')}</div>"
                    جلسة = f"<b style='color:#FFD700'>{c.get('تاريخ_جلسة','')}</b><br>{c.get('سبب','')}"
                    html += f"<tr><td>{i}</td><td>{c.get('رقم','')}</td><td>{c.get('سنة','')}</td><td>{دائرة}</td><td>{محكمة}</td><td>{خصوم}</td><td>{c.get('موضوع','')}</td><td>{جلسة}</td></tr>"
                html += "</table>"
                st.markdown(f"<div class='table-container'>{html}</div>", unsafe_allow_html=True)

            # زر التصدير
            pdf_buffer = generate_pdf_report(cases, "بيان الدعاوى المتداولة", region, from_date, to_date, lawyer)
            st.download_button("⬇️ تحميل التقرير PDF", data=pdf_buffer, file_name=f"report_motadawla_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf", use_container_width=True)
            st.markdown(f"<p style='text-align:right; color:#D4AF37; margin-top:30px; font-size:16px; font-family:Cairo'>تفضلوا بقبول وافر الاحترام<br><br>عضو الادارة..................    مدير الإدارة..................</p>", unsafe_allow_html=True)

    # ========= تبويب 2: الاحكام =========
    with tab2:
        st.markdown("<div style='background:#1E2A47; padding:20px; border-radius:15px; border:2px solid #FF5252; margin-bottom:15px; font-family:Cairo'>", unsafe_allow_html=True)
        region2 = st.text_input("ديوان عام منطقة", key="region2")
        الحكم_نوع = st.selectbox("نوع الاحكام", ["جميع الاحكام", "الاحكام للصالح", "الاحكام للضد"], key="hokm_type")
        col1, col2, col3 = st.columns(3)
        with col1: from_date2 = st.date_input("من الفترة", key="from2")
        with col2: to_date2 = st.date_input("حتى الفترة", key="to2")
        with col3: lawyer2 = st.text_input("طرف الاستاذ/ المحامي", key="lawyer2")
        topic2 = st.text_input("موضوع الدعوى للفلترة - اتركها فاضية لعرض الكل", key="topic2")
        st.markdown("</div>", unsafe_allow_html=True)

        if st.button("🔍 عرض بيان الاحكام", use_container_width=True, type="primary", key="show2"):
            cases = [c for c in data["cases"] if c.get('حالة') == 'منتهية' and c.get('مسندة_ل_الحكم') in ['الصالح','الضد']]
            if الحكم_نوع == "الاحكام للصالح": cases = [c for c in cases if c.get('مسندة_ل_الحكم') == 'الصالح']
            if الحكم_نوع == "الاحكام للضد": cases = [c for c in cases if c.get('مسندة_ل_الحكم') == 'الضد']
            if cases: cases = [c for c in cases if c.get('تاريخ_الحكم') and from_date2 <= datetime.strptime(c['تاريخ_الحكم'], '%Y-%m-%d').date() <= to_date2]
            if topic2: cases = [c for c in cases if topic2 in str(c.get('موضوع',''))]
            cases = sorted(cases, key=lambda x: x.get("تاريخ_الحكم","9999-12-31"), reverse=True)

            report_header(region2, f"بيان ب{الحكم_نوع} خلال الفترة من {from_date2} حتى {to_date2} طرف الاستاذ/ {lawyer2} المحامي")

            if not cases: st.warning("لا توجد احكام في الفترة المحددة")
            else:
                html = "<table class='case-table'><tr><th>م</th><th>رقم القضية</th><th>السنة</th><th>الدائرة</th><th>المحكمة</th><th>الخصوم</th><th>الموضوع</th><th>تاريخ الحكم</th><th>المنطوق</th><th>النتيجة</th></tr>"
                for i, c in enumerate(cases, 1):
                    محكمة = f"{c.get('محكمة_اسم','')}"
                    if c.get('مأمورية'): محكمة += f"<br>مأمورية {c.get('مأمورية')}"
                    دائرة = f"{c.get('دائرة','')} {c.get('نوع','')}"
                    خصوم = f"<div style='background:#FFF3CD; padding:5px; border-radius:5px; color:#000'><b>المدعى:</b> {c.get('مدعي','')}</div><div style='background:#CFF4FC; padding:5px; border-radius:5px; color:#000'><b>المدعى عليه:</b> {c.get('مدعي_عليه','')}</div>"
                    لون = "#4CAF50" if c.get('مسندة_ل_الحكم') == 'الصالح' else "#FF5252"
                    html += f"<tr><td>{i}</td><td>{c.get('رقم','')}</td><td>{c.get('سنة','')}</td><td>{دائرة}</td><td>{محكمة}</td><td>{خصوم}</td><td>{c.get('موضوع','')}</td><td><b style='color:#FFD700'>{c.get('تاريخ_الحكم','')}</b></td><td>{c.get('منطوق_الحكم','')}</td><td style='color:{لون}; font-weight:900'>{c.get('مسندة_ل_الحكم','')}</td></tr>"
                html += "</table>"
                st.markdown(f"<div class='table-container'>{html}</div>", unsafe_allow_html=True)

            # زر التصدير
            pdf_buffer = generate_pdf_report(cases, "بيان الاحكام", region2, from_date2, to_date2, lawyer2)
            st.download_button("⬇️ تحميل التقرير PDF", data=pdf_buffer, file_name=f"report_ahkam_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf", use_container_width=True)
            st.markdown(f"<p style='text-align:right; color:#FF5252; margin-top:30px; font-size:16px; font-family:Cairo'>تفضلوا بقبول وافر الاحترام<br><br>عضو الادارة..................    مدير الإدارة..................</p>", unsafe_allow_html=True)

    # ========= تبويب 3: بيان العضو السنوي =========
    with tab3:
        st.markdown("<div style='background:#1E2A47; padding:20px; border-radius:15px; border:2px solid #4CAF50; margin-bottom:15px; font-family:Cairo'>", unsafe_allow_html=True)
        region3 = st.text_input("ديوان عام منطقة", value="القاهرة", key="region3")
        col1, col2 = st.columns(2)
        with col1: from_date3 = st.date_input("من بداية السنة", datetime(datetime.now().year, 1, 1), key="from3")
        with col2: to_date3 = st.date_input("حتى نهاية السنة", datetime(datetime.now().year, 12, 31), key="to3")
        عضو_الاسم = st.text_input("اسم العضو / المحامي", key="lawyer3")
        st.markdown("</div>", unsafe_allow_html=True)

        if st.button("🔍 استخراج بيان العضو السنوي", use_container_width=True, type="primary", key="show3"):
            cases = [c for c in data["cases"] if عضو_الاسم and (عضو_الاسم in str(c.get('المحامي_المسند','')) or عضو_الاسم in str(c.get('lawyer','')))]
            if cases: cases = [c for c in cases if c.get('تاريخ_جلسة') and from_date3 <= datetime.strptime(c['تاريخ_جلسة'], '%Y-%m-%d').date() <= to_date3]
            
            report_header(region3, f"بيان اجمالي باعمال الاستاذ/ {عضو_الاسم} خلال عام {from_date3.year}")
            
            if not cases: st.warning("لا توجد قضايا مسندة للعضو في الفترة المحددة")
            else:
                متداولة = len([c for c in cases if c.get('حالة')=='متداولة'])
                منتهية = len([c for c in cases if c.get('حالة')=='منتهية'])
                للصالح = len([c for c in cases if c.get('مسندة_ل_الحكم')=='الصالح'])
                
                c1,c2,c3,c4 = st.columns(4)
                c1.metric("اجمالي القضايا", len(cases))
                c2.metric("متداولة", متداولة)
                c3.metric("منتهية", منتهية)
                c4.metric("احكام للصالح", للصالح)

                html = "<table class='case-table'><tr><th>م</th><th>رقم</th><th>سنة</th><th>الحالة</th><th>دائرة</th><th>محكمة</th><th>موضوع</th><th>اخر اجراء</th></tr>"
                for i, c in enumerate(cases, 1):
                    محكمة = f"{c.get('محكمة_اسم','')}"
                    if c.get('مأمورية'): محكمة += f"<br>مأمورية {c.get('مأمورية')}"
                    اجراء = c.get('تاريخ_جلسة','') if c.get('حالة')=='متداولة' else c.get('تاريخ_الحكم','')
                    html += f"<tr><td>{i}</td><td>{c.get('رقم','')}</td><td>{c.get('سنة','')}</td><td>{c.get('حالة','')}</td><td>{c.get('دائرة','')}</td><td>{محكمة}</td><td>{c.get('موضوع','')}</td><td>{اجراء}</td></tr>"
                html += "</table>"
                st.markdown(f"<div class='table-container'>{html}</div>", unsafe_allow_html=True)

            # زر التصدير
            pdf_buffer = generate_pdf_report(cases, "بيان العضو السنوي", region3, from_date3, to_date3, عضو_الاسم)
            st.download_button("⬇️ تحميل التقرير السنوي PDF", data=pdf_buffer, file_name=f"تقرير_العضو_{عضو_الاسم}_{datetime.now().year}.pdf", mime="application/pdf", use_container_width=True)
            st.markdown(f"<p style='text-align:right; color:#4CAF50; margin-top:30px; font-size:16px; font-family:Cairo'>تفضلوا بقبول وافر الاحترام<br><br>عضو الادارة..................    مدير الإدارة..................</p>", unsafe_allow_html=True)

    # ========= تبويب 4: الاحصائيات =========
    with tab4:
        st.markdown("<h3 style='color:#D4AF37; text-align:center; font-family:Cairo'>📊 الإحصائيات العددية</h3>", unsafe_allow_html=True)
        st.markdown("<div style='background:#1E2A47; padding:20px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px; font-family:Cairo'>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: stat_from = st.date_input("من تاريخ", key="s1")
        with col2: stat_to = st.date_input("حتى تاريخ", key="s2")
        st.markdown("</div>", unsafe_allow_html=True)
        
        if st.button("استخراج الإحصائيات", use_container_width=True, type="primary"):
            all_cases = data["cases"]
            متداولة = [c for c in all_cases if c.get('حالة') == 'متداولة' and c.get('تاريخ_جلسة') and stat_from <= datetime.strptime(c['تاريخ_جلسة'], '%Y-%m-%d').date() <= stat_to]
            احكام = [c for c in all_cases if c.get('حالة') == 'منتهية' and c.get('تاريخ_الحكم') and stat_from <= datetime.strptime(c['تاريخ_الحكم'], '%Y-%m-%d').date() <= stat_to]
            للصالح = [c for c in احكام if c.get('مسندة_ل_الحكم') == 'الصالح']
            للضد = [c for c in احكام if c.get('مسندة_ل_الحكم') == 'الضد']
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("عدد القضايا المتداولة", len(متداولة))
            c2.metric("عدد الاحكام الصادرة", len(احكام))
            c3.metric("عدد الاحكام للصالح", len(للصالح))
            c4.metric("عدد الاحكام للضد", len(للضد))
