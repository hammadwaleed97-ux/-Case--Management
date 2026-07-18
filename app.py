# ======== الجزء الاول: الاساسيات ============
# ============================================
import streamlit as st
import pandas as pd
import json
import os
import io
import smtplib
import secrets
import arabic_reshaper # جديد
from bidi.algorithm import get_display # جديد
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF

st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

# دالة عشان تظبط العربي وتوصله
def fix_arabic(text):
    reshaped_text = arabic_reshaper.reshape(str(text))
    return get_display(reshaped_text)

# ====== دالة التصدير للاكسل RTL صح ======
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='التقرير')
        worksheet = writer.sheets['التقرير']
        worksheet.sheet_view.rightToLeft = True
        for col in worksheet.columns:
            worksheet.column_dimensions[col[0].column_letter].width = 20
    return output.getvalue()

# ====== دالة التصدير للورد ======
def to_word(df, title, region):
    doc = Document()
    doc.add_heading(fix_arabic('الهيئة القومية للتأمين الاجتماعى'), 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading(fix_arabic('الإدارة المركزية للإدارات القانونية'), 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading(fix_arabic('الإدارة العامة للقضايا'), 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading(fix_arabic(f'ديوان عام {region}'), 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading(fix_arabic(title), 2).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    df = df.iloc[:, ::-1] # نعكس الاعمدة

    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        hdr_cells[i].text = fix_arabic(col_name)
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = fix_arabic(val)
            row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(fix_arabic(f'\nتفضلوا بقبول وافر الاحترام\nعضو الادارة.................. مدير الإدارة..................\nتحر في {datetime.now().strftime("%Y-%m-%d")}'))
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

# ====== دالة التصدير للـ PDF ======
def to_pdf(df, title, region):
    pdf = FPDF(orientation='L', unit='mm', format='A4') # عرض
    pdf.add_page()
    pdf.add_font('Cairo', '', 'Cairo-Regular.ttf', uni=True)

    pdf.set_font('Cairo', '', 14)
    pdf.cell(0, 8, fix_arabic('الهيئة القومية للتأمين الاجتماعى'), 0, 1, 'C')
    pdf.cell(0, 8, fix_arabic('الإدارة المركزية للإدارات القانونية'), 0, 1, 'C')
    pdf.cell(0, 8, fix_arabic('الإدارة العامة للقضايا'), 0, 1, 'C')
    pdf.cell(0, 8, fix_arabic(f'ديوان عام {region}'), 0, 1, 'C')
    pdf.ln(3)
    pdf.set_font('Cairo', '', 12)
    pdf.cell(0, 8, fix_arabic(title), 0, 1, 'C')
    pdf.ln(3)

    df = df.iloc[:, ::-1] # نعكس الاعمدة

    pdf.set_font('Cairo', '', 7)
    col_width = pdf.w / (len(df.columns) + 1)
    for col in df.columns:
        pdf.cell(col_width, 7, fix_arabic(col), 1, 0, 'C')
    pdf.ln()
    for _, row in df.iterrows():
        for item in row:
            pdf.cell(col_width, 7, fix_arabic(item), 1, 0, 'C')
        pdf.ln()

    pdf.ln(5)
    pdf.set_font('Cairo', '', 11)
    pdf.cell(0, 8, fix_arabic('تفضلوا بقبول وافر الاحترام'), 0, 1, 'R')
    pdf.cell(0, 8, fix_arabic('عضو الادارة.................. مدير الإدارة..................'), 0, 1, 'R')
    pdf.cell(0, 8, fix_arabic(f'تحر في {datetime.now().strftime("%Y-%m-%d")}'), 0, 1, 'R')

    return bytes(pdf.output(dest='S')) # <-- التعديل المهم عشان التحميل

# ====== دالة حفظ صحيفة الدعوى ======
def create_paper_pdf(case_data):
    if not os.path.exists("papers"): os.makedirs("papers")
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.add_font('Cairo', '', 'Cairo-Regular.ttf', uni=True)
    pdf.set_font('Cairo', '', 14); pdf.set_right_margin(15)
    pdf.cell(0,10,fix_arabic(f"صحيفة {case_data.get('مسندة_ل','')}"),ln=1,align='R')
    pdf.ln(5)
    pdf.cell(0,10,fix_arabic(f"محكمة: {case_data.get('محكمة_اسم','')}"),ln=1,align='R')
    pdf.cell(0,10,fix_arabic(f"رقم: {case_data.get('رقم','')} لسنة {case_data.get('سنة','')}"),ln=1,align='R')
    pdf.cell(0,10,fix_arabic(f"المدعي: {case_data.get('مدعي','')}"),ln=1,align='R')
    pdf.cell(0,10,fix_arabic(f"ضد: {case_data.get('مدعي_عليه','')}"),ln=1,align='R')
    pdf.multi_cell(0,10,fix_arabic(f"الموضوع: {case_data.get('موضوع','')}"),align='R')
    name = f"papers/صحيفة_{case_data.get('رقم')}_{case_data.get('سنة')}.pdf"; pdf.output(name); return name

import base64
from io import BytesIO

def print_case_report(case):
    # 1- تحديد الخصوم حسب النوع
    نوع = case.get('نوع', '').lower()
    if 'استئناف' in نوع:
        طرف1_عنوان = "المستأنف"
        طرف2_عنوان = "المستأنف ضده"
    elif 'طعن' in نوع:
        طرف1_عنوان = "الطاعن"
        طرف2_عنوان = "المطعون ضده"
    else:  # دعوى عادية
        طرف1_عنوان = "المدعي"
        طرف2_عنوان = "المدعى عليه"

    html = f"""
    <html dir="rtl" lang="ar">
    <head>
    <meta charset="UTF-8">
    <style>
        @page {{ size: A4; margin: 1.5cm; }}
        body {{ font-family: 'Arial'; direction: rtl; text-align: right; color: #000; background: #f8f9fa; }}
        .header {{ text-align: center; padding: 25px; margin-bottom: 25px; background: linear-gradient(135deg, #1E2A47 0%, #D4AF37 100%); color: #FFF; border-radius: 15px; box-shadow: 0 5px 15px rgba(0,0,0,0.2); }}
        .logo {{ font-size: 22px; font-weight: 900; color: #FFF; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); }}
        .sub {{ font-size: 16px; color: #FFF9E6; margin: 8px 0; }}
        .title {{ text-align: center; font-size: 26px; font-weight: 900; color: #1E2A47; margin: 25px 0; border: 3px solid #D4AF37; padding: 15px; border-radius: 15px; background: linear-gradient(90deg, #FFF9E6, #FFF); box-shadow: 0 3px 10px rgba(212,175,55,0.3); }}
        .section {{ padding: 20px; border-radius: 15px; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); border: 2px solid transparent; }}
        .section-title {{ font-weight: 900; font-size: 20px; color: #FFF; margin-bottom: 20px; text-align: center; padding: 12px; border-radius: 10px; }}
        .sec1 {{ background: linear-gradient(135deg, #1E2A47, #3498db); border-color: #1E2A47; }}
        .sec2 {{ background: linear-gradient(135deg, #27ae60, #2ecc71); border-color: #27ae60; }}
        .sec3 {{ background: linear-gradient(135deg, #8e44ad, #9b59b6); border-color: #8e44ad; }}
        .sec4 {{ background: linear-gradient(135deg, #c0392b, #e74c3c); border-color: #c0392b; }}
        .row {{ display: flex; justify-content: space-between; margin-bottom: 12px; background: linear-gradient(90deg, #fff, #f8f9fa); padding: 12px; border-radius: 8px; border-right: 4px solid #D4AF37; }}
        .label {{ font-weight: 900; color: #1E2A47; width: 35%; font-size: 15px; }}
        .value {{ width: 65%; color: #000; font-weight: 700; font-size: 15px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 15px; border-radius: 10px; overflow: hidden; box-shadow: 0 3px 10px rgba(0,0,0,0.1); }}
        th {{ background: linear-gradient(135deg, #1E2A47, #34495e); color: #D4AF37; padding: 12px; border: none; text-align: center; font-size: 16px; font-weight: 900; }}
        td {{ padding: 12px; border-bottom: 1px solid #ddd; text-align: center; background: #fff; }}
        tr:nth-child(even) td {{ background: #f8f9fa; }}
        tr:hover td {{ background: #FFF9E6; }}
    </style>
    </head>
    <body>
    
    <div class="header">
        <div class="logo">الهيئة القومية للتأمين الاجتماعي</div>
        <div class="sub">الإدارة المركزية للإدارات القانونية</div>
        <div class="sub">الإدارة العامة للشئون القانونية منطقة: _____________</div>
    </div>

    <div class="title">📄 تقرير تفاصيل القضية رقم {case.get('رقم')} لسنة {case.get('سنة')}</div>

    <div class="section sec1">
        <div class="section-title sec1">1- بيانات القضية</div>
        <div class="row"><div class="label">رقم القضية:</div><div class="value">{case.get('رقم')}</div></div>
        <div class="row"><div class="label">السنة:</div><div class="value">{case.get('سنة')}</div></div>
        <div class="row"><div class="label">النوع:</div><div class="value">{case.get('نوع')}</div></div>
        <div class="row"><div class="label">المحكمة:</div><div class="value">{case.get('محكمة_اسم')} {f'- مأمورية {case.get("مأمورية")}' if case.get('مأمورية') else ''}</div></div>
        <div class="row"><div class="label">الدائرة:</div><div class="value">{case.get('دائرة')}</div></div>
        <div class="row"><div class="label">الحالة:</div><div class="value">{case.get('حالة')}</div></div>
        <div class="row"><div class="label">الموضوع:</div><div class="value">{case.get('موضوع')}</div></div>
    </div>

    <div class="section sec2">
        <div class="section-title sec2">2- بيانات الخصوم</div>
        <div class="row"><div class="label">{طرف1_عنوان}:</div><div class="value">{case.get('مدعي')}</div></div>
        <div class="row"><div class="label">{طرف2_عنوان}:</div><div class="value">{case.get('مدعي_عليه')}</div></div>
    </div>
    """

    # 3- الجلسات والإجراءات
    if case.get("جلسات"):
        html += """
        <div class="section sec3">
            <div class="section-title sec3">3- الجلسات والإجراءات</div>
            <table>
                <tr><th>م</th><th>الرول</th><th>الجلسات</th><th>الإجراءات</th><th>ملاحظات</th></tr>
        """
        for i, ج in enumerate(case["جلسات"], 1):
            html += f"<tr><td>{i}</td><td>{ج.get('الرول')}</td><td>{ج.get('تاريخ')}</td><td>{ج.get('الاجراء')}</td><td>{ج.get('ملاحظات')}</td></tr>"
        html += "</table></div>"

    # 4- الحكم
    if case.get('حالة') == 'منتهية':
        html += f"""
        <div class="section sec4">
            <div class="section-title sec4">4- منطوق الحكم</div>
            <div class="row"><div class="label">تاريخ الحكم:</div><div class="value">{case.get('تاريخ_الحكم')}</div></div>
            <div class="row"><div class="label">مسندة لـ:</div><div class="value">{case.get('مسندة_ل_الحكم')}</div></div>
            <div class="row"><div class="label">المنطوق:</div><div class="value">{case.get('منطوق_الحكم')}</div></div>
        </div>
        """
    
    html += "</body></html>"
    return html
# ====== دالة التحميل والحفظ الوحيدة ======
DATA_FILE = "cases_data.json"
TOKENS_FILE = "tokens.json"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER,exist_ok=True)

def load_data():
    if not os.path.exists(DATA_FILE):
        return {"cases":[],"library":[]}
    try:
        with open(DATA_FILE,"r",encoding="utf-8") as f:
            data=json.load(f)
        if not isinstance(data,dict): data={}
        data.setdefault("cases",[])
        data.setdefault("library",[])
        return data
    except Exception:
        return {"cases":[],"library":[]}

def save_data(data):
    data.setdefault("cases",[])
    data.setdefault("library",[])
    with open(DATA_FILE,"w",encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def load_tokens():
    if os.path.exists(TOKENS_FILE):
        try:
            with open(TOKENS_FILE,"r",encoding="utf-8") as f: return json.load(f)
        except: pass
    return {"tokens":[]}

def save_tokens(tokens_data):
    with open(TOKENS_FILE,"w",encoding="utf-8") as f:
        json.dump(tokens_data, f, ensure_ascii=False, indent=4)

# ====== دوال التنبيهات ======
def get_alert_cases():
    data = load_data()
    today = datetime.now().date()
    all_cases = data.get("cases", [])
    alerts = {"sessions": [], "appeals": []}
    for case in all_cases:
        if case.get("حالة") == "متداولة" and case.get("تاريخ_جلسة"):
            try:
                session_date = datetime.strptime(case["تاريخ_جلسة"], "%Y-%m-%d").date()
                days_left = (session_date - today).days
                if 0 <= days_left <= 7:
                    case_copy = case.copy(); case_copy["days_left"] = days_left; alerts["sessions"].append(case_copy)
            except Exception: pass
        if case.get("حالة") == "منتهية" and case.get("مسندة_ل_الحكم") == "الضد" and case.get("تاريخ_الحكم"):
            try:
                judgment_date = datetime.strptime(case["تاريخ_الحكم"], "%Y-%m-%d").date()
                appeal_days = 40 if case.get("نوع") == "دعوى" else 60
                last_appeal_day = judgment_date + timedelta(days=appeal_days)
                notify_start = last_appeal_day - timedelta(days=15)
                days_left_appeal = (last_appeal_day - today).days
                if notify_start <= today <= last_appeal_day and days_left_appeal >= 0:
                    case_copy = case.copy(); case_copy["days_left_appeal"] = days_left_appeal; alerts["appeals"].append(case_copy)
            except Exception: pass
    return alerts

LIBRARY_SECTIONS = {
    "القوانين": "#FF4500", "القرارات الوزارية": "#FF8C00", "قرارات الهيئة": "#FFD700",
    "المنشورات الوزارية": "#ADFF2F", "منشورات الهيئة": "#32CD32", "الكتب الدورية": "#20B2AA",
    "تعليمات الهيئة": "#00CED1", "رسائل الهيئة": "#1E90FF", "المرصد الفنى": "#4169E1",
    "فتاوى لجنة الشئون القانونية بالوزارة": "#8A2BE2", "فتاوى الادارة المركزية للشئون القانونية": "#9400D3",
    "احكام المحكمة الدستورية العليا": "#DC143C", "احكام محكمة النقض": "#B22222", "احكام المحكمة الإدارية العليا": "#8B0000",
    "احكام المحاكم الاستئنافية": "#A0522D", "احكام محاكم القضاء الإدارى": "#D2691E", "احكام المحاكم الابتدائية": "#CD853F",
    "احكام المحكمة الإدارية": "#DEB887", "منشورات القضاء العادى": "#5F9EA0", "منشورات مجلس الدولة": "#4682B4",
    "فتاوى الجمعية العمومية": "#7B68EE", "صحف طعون": "#6A5ACD", "صحف استئنافات": "#483D8B",
    "صحف دعاوى": "#E6E6FA", "مذكرات دفاع": "#FFF0F5", "أخرى": "#808080"
}

SENDER_EMAIL=""; SENDER_PASSWORD=""; APP_URL="https://qpyqpsmkqcvdou4imbfunp.streamlit.app/"
ANWA3_MOSTANDAT = ["صحيفة دعوى","صحيفة استئناف","صحيفة طعن","مذكرة دفاع","حافظة مستندات","تقرير خبير","تقرير طب شرعى","تقرير لجنة طبية","صحيفة تجديد من الشطب","صحيفة تعجيل من الوقف","صورة حكم تمهيدى","أخرى"]

if "page" not in st.session_state: st.session_state.page="الرئيسية"
if "selected_case_id" not in st.session_state: st.session_state.selected_case_id=None

# ============= التصميم النهائي =============
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');
*{font-family:'Cairo',sans-serif!important;}
html,body{direction:rtl;color:#FFF!important;}
.stApp{background:linear-gradient(180deg,#0A1428 0%,#1E2A47 100%);}
.marquee{background:linear-gradient(90deg,#D4AF37,#FFD700,#D4AF37);color:#0A1428;padding:12px;font-weight:900;font-size:16px;white-space:nowrap;overflow:hidden;border-radius:0 0 15px 15px;}
.marquee span{display:inline-block;animation:marquee 15s linear infinite;}
@keyframes marquee{0%{transform:translateX(-100%);}100%{transform:translateX(100%);}}
.main-title{color:#D4AF37;text-align:center;font-size:36px;font-weight:900;padding:15px 0;}
h1,h2,h3{color:#D4AF37!important;text-align:center!important;}
div[data-testid="column"]{display:flex;justify-content:center;}
[data-testid="stForm"] label,.stMarkdown{color:#FFF!important;font-weight:700;}
.stButton>button{width:100%!important;max-width:400px!important;border:none!important;border-radius:15px!important;font-size:18px!important;font-weight:900!important;padding:16px!important;color:#000!important;}
</style>
""", unsafe_allow_html=True)

st.markdown("""<div class="marquee"><span>مع تحيات وليد حماد - الإدارة العامة للشئون القانونية بديوان عام منطقة البحيرة بالهيئة القومية للتأمين الاجتماعي</span></div>""", unsafe_allow_html=True)
st.markdown('<div class="main-title">⚖️ إدارة القضايا ⚖️</div>', unsafe_allow_html=True)

# =========================================
# =======================================
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
        # =============================
# ====== الجزء الثاني: تسجيل القضية ============
elif st.session_state.page == "تسجيل":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#D4AF37; text-align:center'> تسجيل القضية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", key="back_add", use_container_width=True):
        st.session_state.page = "الرئيسية"
        st.rerun()

    st.markdown("<label style='color:#FFF; font-weight:700; text-align:right; width:100%; display:block;'>نوع القضية</label>", unsafe_allow_html=True)
    نوع = st.selectbox("", ["دعوى", "استئناف", "طعن"], key="case_type_add")
    
    with st.form("form_case_add", clear_on_submit=True):  # <-- دي بتفضي الفورم لوحدها
        # 1- بيانات المحكمة
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        محكمة_اسم = st.text_input("اسم المحكمة", key="court_name_add")
        مأمورية = st.text_input("المأمورية", key="mamoria_add") if نوع == "استئناف" else ""
        st.markdown("</div>", unsafe_allow_html=True)

        # 2- بيانات القضية
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>2- بيانات القضية</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: رقم = st.text_input("رقم القضية / الاستئناف / الطعن", key="case_num_add")
        with col2: سنة = st.text_input("السنة القضائية", key="case_year_add")
        دائرة = st.text_input("الدائرة", key="circle_add")
        st.markdown("</div>", unsafe_allow_html=True)

        # 3- بيانات الخصوم
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>3- بيانات الخصوم</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: مدعي = st.text_input("اسم المدعى / المستأنف / الطاعن", key="plaintiff_add")
        with col2: مدعي_عليه = st.text_input("اسم المدعى عليه / المستأنف ضده / المطعون ضده", key="defendant_add")
        موضوع = st.text_area("موضوع القضية", height=100, key="subject_add")
        st.markdown("</div>", unsafe_allow_html=True)

        # 4- بيانات الجلسة
        st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>4- بيانات الجلسة</div>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1: تاريخ_جلسة = st.date_input("تاريخ أول جلسة", value=datetime.now().date(), key="session_date_add")
        with col2: الرول = st.text_input("الرول", key="roll_add")
        الاجراء = st.text_input("الاجراء", key="reason_add") # <-- 1. غيرنا الاسم هنا
        ملاحظات = st.text_area("ملاحظات", height=100, key="notes_add")
        st.markdown("</div>", unsafe_allow_html=True)

        if st.form_submit_button("💾 حفظ القضية", use_container_width=True, type="primary"):
            if not رقم or not سنة: 
                st.error("❌ من فضلك ادخل رقم القضية والسنة")
            else:
                case_for_pdf = {"نوع":نوع,"رقم":رقم,"سنة":سنة,"دائرة":دائرة,"محكمة_اسم":محكمة_اسم,"مدعي":مدعي,"مدعي_عليه":مدعي_عليه,"موضوع":موضوع,"تاريخ_جلسة":str(تاريخ_جلسة)}
                paper_path = create_paper_pdf(case_for_pdf)

                new_case = {
                    "id": len(data["cases"])+1, "نوع": نوع, "محكمة_اسم": محكمة_اسم, "مأمورية": مأمورية, 
                    "رقم": رقم, "سنة": سنة, "دائرة": دائرة, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه, 
                    "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة), "الرول": الرول, "الاجراء": الاجراء, # <-- 2. وهنا
                    "ملاحظات": ملاحظات, "جلسات": [], "مستندات": [paper_path], "حالة": "متداولة"
                }
                if الرول or الاجراء: # <-- 3. وهنا
                    new_case["جلسات"].append({"تاريخ":str(تاريخ_جلسة),"الرول":الرول,"الاجراء":الاجراء,"ملاحظات":ملاحظات}) # <-- 4. وهنا
                
                data["cases"].append(new_case)
                save_data(data)
                
                st.success(f"✅ تم الحفظ بنجاح -ونقلت للحصر العام- جاهز لتسجيل قضية جديدة")
                # ========================
# ====== الجزء الثالث: الحصر العام ============
# ================================================
elif st.session_state.page == "الحصر":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📊 الحصر العام الخارجي</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()

    if st.session_state.get('open_from_search', False):
        st.session_state.open_from_search = False
        st.info("جاري فتح القضية من البحث...")

    if not data["cases"]:
        st.info("لا توجد قضايا مسجلة")
    else:
        for i, case in enumerate(data["cases"]):
            if "id" not in case: case["id"] = i + 1
            if "مستندات" not in case: case["مستندات"] = []

        save_data(data)

        # ======= تحديث اخر جلسة والاجراء من الجلسات =======
        for case in data["cases"]:
            if "جلسات" in case and case["جلسات"]:
                جلسات_مرتبة = sorted(case["جلسات"], key=lambda x: x.get("تاريخ","9999-12-31"), reverse=True)
                اخر_جلسة = جلسات_مرتبة[0]
                case["تاريخ_جلسة"] = اخر_جلسة.get("تاريخ","")
                case["الاجراء"] = اخر_جلسة.get("الاجراء","") # <-- هنا بقت الاجراء
                case["الحالة"] = اخر_جلسة.get("الحالة", case.get("الحالة","متداولة"))
        save_data(data)
        # ============================================

        # ======= التعديل 1: نشيل المنتهية من الحصر =======
        active_cases = [c for c in data["cases"] if c.get('حالة')!= 'منتهية']
        # ==================================================

        sorted_cases = sorted(active_cases, key=lambda x: x.get("تاريخ_جلسة","9999-12-31"))
        total = len(data["cases"]) # اجمالي الكل زي ما هو
        today = datetime.now().date()
        start_week = today - timedelta(days=(today.weekday() + 2) % 7) # السبت
        end_week = start_week + timedelta(days=5) # الخميس

        # ======= التعديل 2: جلسات الاسبوع من السبت للخميس =======
        this_week = len([c for c in data["cases"] if c.get('تاريخ_جلسة') and start_week <= datetime.strptime(c['تاريخ_جلسة'],'%Y-%m-%d').date() <= end_week])
        # =========================================================

        # ====== التعديل 3: المحجوزة للحكم بدل المنتهية =======
        reserved = len([c for c in data["cases"] if any(k in str(c.get('الاجراء','')) for k in ['حكم', 'للحكم', 'الحكم', 'محجوزة للحكم'])])
        # =======================================================

        st.markdown(f"<div style='background:#1E2A47; padding:20px; border-radius:15px; border:2px solid #D4AF37; text-align:center; margin-bottom:20px'>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1: st.markdown(f"<div style='font-size:28px; font-weight:900; color:#D4AF37'>📊 {total}</div><div style='font-size:18px; color:#FFF; font-weight:700'>اجمالي القضايا</div>", unsafe_allow_html=True)
        with col2: st.markdown(f"<div style='font-size:28px; font-weight:900; color:#4DA8DA'>📅 {this_week}</div><div style='font-size:18px; color:#FFF; font-weight:700'>جلسات هذا الاسبوع</div>", unsafe_allow_html=True)
        with col3: st.markdown(f"<div style='font-size:28px; font-weight:900; color:#FF5252'>⚖️ {reserved}</div><div style='font-size:18px; color:#FFF; font-weight:700'>محجوزة للحكم</div>", unsafe_allow_html=True) # اتغيرت هنا
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("""
        <style>
 .case-table {width:100%; border-collapse: collapse; font-size:11px; color:white; text-align:center; margin-bottom:5px;}
 .case-table th {background:#D4AF37; color:#0B1426; padding:6px; font-weight:900;}
 .case-table td {background:#1E2A47; padding:6px; border:1px solid #D4AF37; vertical-align:top;}
 .plaintiff {background:#FFF3CD; color:#000; font-weight:700; border-radius:6px; padding:6px; font-size:11px;}
 .plaintiff-hey2a {background:#DC3545!important; color:#FFF!important; font-weight:900; border-radius:6px; padding:6px; font-size:11px;}
 .defendant {background:#CFF4FC; color:#000; font-weight:700; border-radius:6px; padding:6px; font-size:11px;}
 .date-gold {color:#FFD700; font-weight:900;}
 .status-green {color:#4CAF50; font-weight:900;}
        </style>
        """, unsafe_allow_html=True)

        for idx, case in enumerate(sorted_cases, 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
            if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = f"{case.get('دائرة', '' )}" if case.get('دائرة', '' ) else ""
            if دائرة_كاملة: محكمة_كاملة += f"<br>دائرة {دائرة_كاملة}"

            نوع = case.get('نوع','')
            if نوع == "استئناف": لقب1, لقب2 = "المستأنف:", "المستأنف ضده:"
            elif نوع == "طعن": لقب1, لقب2 = "الطاعن:", "المطعون ضده:"
            else: لقب1, لقب2 = "المدعى:", "المدعى عليه:"

            if "الهيئة" in str(case.get('مدعي','')):
                طرف1_html = f"<div class='plaintiff-hey2a'><b>{لقب1}</b><br>{case.get('مدعي','')}</div>"
            else:
                طرف1_html = f"<div class='plaintiff'><b>{لقب1}</b><br>{case.get('مدعي','')}</div>"
            طرف2_html = f"<div class='defendant'><b>{لقب2}</b><br>{case.get('مدعي_عليه','')}</div>"
            خصوم = طرف1_html + "<div style='height:4px'></div>" + طرف2_html

            table_html = "<table class='case-table'><tr>"
            headers = ["م", "الرقم والسنة", "المحكمة والدائرة", "الخصوم", "الموضوع", "اخر جلسة", "الاجراء", "الحالة"] # <-- هنا بقت الاجراء
            for h in headers: table_html += f"<th>{h}</th>"
            table_html += "</tr>"
            table_html += f"<tr><td>{idx}</td><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td class='date-gold'>{case.get('تاريخ_جلسة','')}</td><td>{case.get('الاجراء','')}</td><td class='status-green'>{case.get('حالة','متداولة')}</td></tr></table>" # <-- وهنا كمان
            st.markdown(table_html, unsafe_allow_html=True)

            c1, c2, c3 = st.columns([4,1,4])
            with c2:
                if st.button("فتح", key=f"open_{case['id']}", use_container_width=True):
                    st.session_state.selected_case_id = case['id']; st.session_state.page = "تفاصيل"; st.rerun()

# ======================================
# ====================================
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

    if st.button("⬅️ العودة للحصر", use_container_width=True): st.session_state.page = "الحصر"; st.rerun()

    # 1- بيانات القضية
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:15px'>1- بيانات القضية</div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1: st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>رقم القضية</div><div style='color:#FFF; font-weight:900; font-size:22px'>{case.get('رقم')}</div></div>", unsafe_allow_html=True)
    with col2: st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>السنة</div><div style='color:#FFF; font-weight:900; font-size:22px'>{case.get('سنة')}</div></div>", unsafe_allow_html=True)
    with col3: دائرة_نص = f"{case.get('دائرة')}" if case.get('دائرة') else ""; st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>الدائرة</div><div style='color:#FFF; font-weight:900; font-size:18px'>{دائرة_نص}</div></div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1: st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>النوع</div><div style='color:#FFF; font-weight:900; font-size:18px'>{case.get('نوع')}</div></div>", unsafe_allow_html=True)
    with col2: محكمة_كاملة = f"{case.get('محكمة_اسم')}";
    if case.get('مأمورية'): محكمة_كاملة += f" - مأمورية {case.get('مأمورية')}"; st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>المحكمة</div><div style='color:#FFF; font-weight:700; font-size:14px'>{محكمة_كاملة}</div></div>", unsafe_allow_html=True)
    with col3: st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>الحالة</div><div style='color:#4CAF50; font-weight:900; font-size:18px'>{case.get('حالة')}</div></div>", unsafe_allow_html=True)
    st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>الموضوع</div><div style='color:#FFF; font-weight:700; font-size:16px'>{case.get('موضوع')}</div></div>", unsafe_allow_html=True)
    st.markdown("<style>div[data-testid='stExpander'] summary p{color:#D4AF37!important; font-weight:900!important;}</style>", unsafe_allow_html=True)
    with st.expander("✏️ تعديل بيانات القضية"):
        with st.form("edit_case_form"):
            col1, col2, col3 = st.columns(3)
            with col1: رقم_جديد = st.text_input("رقم القضية", value=case.get('رقم','')); سنة_جديد = st.text_input("السنة", value=case.get('سنة','')); نوع_جديد = st.text_input("النوع", value=case.get('نوع',''))
            with col2: محكمة_جديد = st.text_input("اسم المحكمة", value=case.get('محكمة_اسم','')); مأمورية_جديد = st.text_input("المأمورية", value=case.get('مأمورية','')); دائرة_جديد = st.text_input("الدائرة", value=case.get('دائرة',''))
            with col3: مدعي_جديد = st.text_input("المدعي", value=case.get('مدعي','')); مدعي_عليه_جديد = st.text_input("المدعي عليه", value=case.get('مدعي_عليه','')); حالة_جديد = st.selectbox("الحالة", ["متداولة", "مؤجلة", "منتهية", "شطب"], index=["متداولة", "مؤجلة", "منتهية", "شطب"].index(case.get('حالة','متداولة')) if case.get('حالة') in ["متداولة", "مؤجلة", "منتهية", "شطب"] else 0)
            موضوع_جديد = st.text_area("الموضوع", value=case.get('موضوع',''), height=100)
            if st.form_submit_button("💾 حفظ التعديلات", use_container_width=True, type="primary"):
                case['رقم']=رقم_جديد; case['سنة']=سنة_جديد; case['نوع']=نوع_جديد; case['محكمة_اسم']=محكمة_جديد; case['مأمورية']=مأمورية_جديد; case['دائرة']=دائرة_جديد; case['مدعي']=مدعي_جديد; case['مدعي_عليه']=مدعي_عليه_جديد; case['حالة']=حالة_جديد; case['موضوع']=موضوع_جديد
                save_data(data); st.success("✅ تم حفظ التعديلات"); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    # 2- بيانات الخصوم
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>2- بيانات الخصوم</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1: st.markdown(f"<div style='background:#FFF3CD; padding:10px; border-radius:10px; color:#000; text-align:center'><b>المدعى:</b><br>{case.get('مدعي')}</div>", unsafe_allow_html=True)
    with col2: st.markdown(f"<div style='background:#CFF4FC; padding:10px; border-radius:10px; color:#000; text-align:center'><b>المدعى عليه:</b><br>{case.get('مدعي_عليه')}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # 3- الجلسات والإجراءات
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>3- الجلسات والإجراءات</div>", unsafe_allow_html=True)
    if case.get("جلسات"):
        for i, ج in enumerate(case["جلسات"]):
            st.markdown(f"<div style='background:#142038; padding:15px; border-radius:12px; border:2px solid #D4AF37; margin-bottom:10px; text-align:right; direction:rtl'><div style='display:flex; justify-content:flex-end; margin-bottom:10px'><div style='background:#D4AF37; color:#000; padding:5px 15px; border-radius:8px; font-weight:900; font-size:16px'>جلسة {i+1}</div></div><div style='margin-bottom:8px'><span style='color:#D4AF37; font-weight:900'>التاريخ:</span> <span style='color:#FFF'>{ج.get('تاريخ')}</span></div><div style='margin-bottom:8px'><span style='color:#D4AF37; font-weight:900'>الرول:</span> <span style='color:#FFF'>{ج.get('الرول')}</span></div><div style='margin-bottom:8px'><span style='color:#D4AF37; font-weight:900'>الاجراء:</span> <span style='color:#FFF'>{ج.get('الاجراء')}</span></div><div><span style='color:#D4AF37; font-weight:900'>ملاحظات:</span> <span style='color:#FFF'>{ج.get('ملاحظات')}</span></div></div>", unsafe_allow_html=True)
            if st.button("✏️ تعديل الجلسة", key=f"edit_session_{i}", use_container_width=True):
                st.session_state.edit_session_index = i; st.rerun()
        if 'edit_session_index' in st.session_state and st.session_state.edit_session_index is not None:
            idx = st.session_state.edit_session_index; جلسة = case["جلسات"][idx]
            with st.form("edit_session_form"):
                st.warning(f"تعديل الجلسة رقم {idx+1}")
                تاريخ_تعديل = st.date_input("التاريخ", value=datetime.strptime(جلسة.get('تاريخ'),'%Y-%m-%d'))
                رول_تعديل = st.text_input("الرول", value=جلسة.get('الرول','')); اجراء_تعديل = st.text_input("الاجراء", value=جلسة.get('الاجراء','')); ملاحظات_تعديل = st.text_area("الملاحظات", value=جلسة.get('ملاحظات',''))
                c1,c2 = st.columns(2)
                with c1:
                    if st.form_submit_button("💾 حفظ تعديل الجلسة", use_container_width=True):
                        case["جلسات"][idx] = {"تاريخ":str(تاريخ_تعديل),"الرول":رول_تعديل,"الاجراء":اجراء_تعديل,"ملاحظات":ملاحظات_تعديل}
                        جلسات_مرتبة = sorted(case["جلسات"], key=lambda x: x.get("تاريخ","9999-12-31"), reverse=True)
                        case["تاريخ_جلسة"] = جلسات_مرتبة[0].get("تاريخ",""); case["الاجراء"] = جلسات_مرتبة[0].get("الاجراء","")
                        save_data(data); st.session_state.edit_session_index = None; st.success("تم التعديل"); st.rerun()
                with c2:
                    if st.form_submit_button("❌ الغاء", use_container_width=True): st.session_state.edit_session_index = None; st.rerun()
    else: st.info("لا توجد جلسات مسجلة")
    st.markdown("<style>div[data-testid='stExpander'] summary p{color:white!important; font-weight:900!important;}</style>", unsafe_allow_html=True)
    with st.expander("اضافة جلسة جديدة"):
        with st.form("add_session"):
            تاريخ_جديد = st.date_input("تاريخ الجلسة", value=datetime.now()); رول_جديد = st.text_input("الرول"); الاجراء_جديد = st.text_input("الاجراء"); ملاحظات_جديدة = st.text_area("ملاحظات")
            if st.form_submit_button("حفظ الجلسة"):
                case["جلسات"].append({"تاريخ":str(تاريخ_جديد),"الرول":رول_جديد,"الاجراء":الاجراء_جديد,"ملاحظات":ملاحظات_جديدة})
                جلسات_مرتبة = sorted(case["جلسات"], key=lambda x: x.get("تاريخ","9999-12-31"), reverse=True)
                case["تاريخ_جلسة"] = جلسات_مرتبة[0].get("تاريخ",""); case["الاجراء"] = جلسات_مرتبة[0].get("الاجراء","")
                save_data(data); st.success("تم اضافة الجلسة"); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    # 4- المستندات - متعدل
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>4- المستندات</div>", unsafe_allow_html=True)
    ANWA3_MOSTANDAT = ["صحيفة دعوى", "صحيفة استئناف", "صحيفة طعن", "مذكرة دفاع","حافظة مستندات", "تقرير خبير", "تقرير طب شرعى", "تقرير لجنة طبية","صحيفة تجديد من الشطب", "صحيفة تعجيل من الوقف", "صورة حكم تمهيدى", "أخرى"]
    نوع_المستند = st.selectbox("نوع المستند", ANWA3_MOSTANDAT, key="select_doc_type")
    اسم_نهائي = نوع_المستند
    if نوع_المستند == "أخرى": اسم_نهائي = st.text_input("✍️ اكتب اسم المستند", placeholder="مثال: طلب / انذار / الخ")
    with st.form("upload_form"):
        uploaded_file = st.file_uploader("اختر الملف", type=['pdf', 'jpg', 'jpeg', 'png', 'doc', 'docx'])
        if st.form_submit_button("رفع المستند"):
            if uploaded_file and اسم_نهائي and اسم_نهائي.strip()!= "":
                file_name = f"{اسم_نهائي}_{uploaded_file.name}"
                file_base64 = base64.b64encode(uploaded_file.getvalue()).decode('utf-8')
                case['مستندات'].append({"نوع": file_name, "محتوى": file_base64})
                save_data(data); st.success("✅ تم رفع المستند"); st.rerun()
            else: st.error("❌ لازم تختار ملف وتكتب اسم المستند")

    if case.get('مستندات'):
        # اصلاح الداتا القديمة
        مستندات_متصلحة = []
        for مستند in case['مستندات']:
            if isinstance(مستند, str): مستندات_متصلحة.append({"نوع": مستند, "محتوى": ""})
            else: مستندات_متصلحة.append(مستند)
        case['مستندات'] = مستندات_متصلحة; save_data(data)

        st.markdown("<div style='background:#142038; padding:15px; border-radius:12px; margin-top:10px'>", unsafe_allow_html=True)
        st.markdown("<div style='color:#D4AF37; font-weight:900; margin-bottom:10px'>المستندات المرفوعة:</div>", unsafe_allow_html=True)
        for i, مستند in enumerate(case['مستندات']):
            col1, col2, col3 = st.columns([4,1,1])
            with col1: st.write(f"📄 {مستند.get('نوع', f'ملف رقم {i+1}')}")
            with col2:
                if مستند.get('محتوى'):
                    try: file_data = base64.b64decode(مستند['محتوى']); st.download_button("📥 تحميل", data=file_data, file_name=مستند.get('نوع'), mime="application/octet-stream", key=f"dl_{i}", use_container_width=True)
                    except: st.write("❌")
            with col3:
                if st.button("🗑️ حذف", key=f"del_{i}", use_container_width=True): case['مستندات'].pop(i); save_data(data); st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # 5- جلسة الحكم - متعدل
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #FF5252; margin-bottom:15px'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#FF5252; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>5- جلسة الحكم</div>", unsafe_allow_html=True)
    if case.get('حالة') == 'منتهية':
        لون = "#4CAF50" if case.get('مسندة_ل_الحكم') == "الصالح" else "#FF5252"
        st.markdown(f"<div style='background:#142038; padding:15px; border-radius:12px; border:2px solid {لون}; margin-bottom:10px'><b style='color:{لون}'>تاريخ الحكم:</b> {case.get('تاريخ_الحكم')}<br><b style='color:{لون}'>مسندة لـ:</b> {case.get('مسندة_ل_الحكم')}<br><b style='color:{لون}'>المنطوق:</b> {case.get('منطوق_الحكم')}</div>", unsafe_allow_html=True)
        st.success("✅ تم نقلها للارشيف للمتابعة")
        with st.expander("✏️ تعديل بيانات الحكم"):
            with st.form("edit_judgment_form"):
                تاريخ_حكم_تعديل = st.date_input("تاريخ الحكم", value=datetime.strptime(case.get('تاريخ_الحكم'),'%Y-%m-%d'))
                منطوق_الحكم_تعديل = st.text_area("منطوق الحكم", value=case.get('منطوق_الحكم',''), height=150)
                مسندة_ل_تعديل = st.selectbox("مسندة لـ", ["الصالح", "الضد"], index=["الصالح", "الضد"].index(case.get('مسندة_ل_الحكم','الصالح')))
                if st.form_submit_button("💾 حفظ تعديل الحكم", use_container_width=True, type="primary"):
                    case['تاريخ_الحكم'] = str(تاريخ_حكم_تعديل); case['منطوق_الحكم'] = منطوق_الحكم_تعديل; case['مسندة_ل_الحكم'] = مسندة_ل_تعديل
                    for ج in reversed(case['جلسات']):
                        if 'الحكم' in ج.get('الاجراء',''): ج['تاريخ'] = str(تاريخ_حكم_تعديل); ج['الاجراء'] = f'الحكم - مسندة لـ {مسندة_ل_تعديل}'; ج['ملاحظات'] = منطوق_الحكم_تعديل; break
                    case['تاريخ_جلسة'] = str(تاريخ_حكم_تعديل); case['الاجراء'] = f'الحكم - مسندة لـ {مسندة_ل_تعديل}'
                    save_data(data); st.success("✅ تم تعديل الحكم"); st.rerun()
    else:
        with st.form("judgment_form"):
            st.markdown("<div style='background:#142038; padding:10px; border-radius:10px; margin-bottom:10px'>", unsafe_allow_html=True)
            st.markdown("<label style='color:#FFD700; font-weight:900; font-size:16px'>1- تاريخ الجلسة</label>", unsafe_allow_html=True); تاريخ_حكم = st.date_input("تاريخ الجلسة", value=datetime.now().date(), label_visibility="collapsed"); st.markdown("</div>", unsafe_allow_html=True)
            st.markdown("<div style='background:#142038; padding:10px; border-radius:10px; margin-bottom:10px'>", unsafe_allow_html=True)
            st.markdown("<label style='color:#FFD700; font-weight:900; font-size:16px'>2- منطوق الحكم</label>", unsafe_allow_html=True); منطوق_الحكم = st.text_area("منطوق الحكم", height=150, placeholder="اكتب منطوق الحكم هنا...", label_visibility="collapsed"); st.markdown("</div>", unsafe_allow_html=True)
            st.markdown("<div style='background:#142038; padding:10px; border-radius:10px; margin-bottom:10px'>", unsafe_allow_html=True)
            st.markdown("<label style='color:#FFD700; font-weight:900; font-size:16px'>3- مسندة لـ</label>", unsafe_allow_html=True); مسندة_ل = st.selectbox("مسندة لـ", ["الصالح", "الضد"], label_visibility="collapsed"); st.markdown("</div>", unsafe_allow_html=True)
            st.markdown("<div style='background:#142038; padding:10px; border-radius:10px; margin-bottom:10px'>", unsafe_allow_html=True)
            st.markdown("<label style='color:#FFD700; font-weight:900; font-size:16px'>4- سبب الحفظ</label>", unsafe_allow_html=True); سبب_الحفظ = st.text_area("سبب الحفظ", placeholder="عدم جدوى طعن / تم الطعن رقم...", label_visibility="collapsed"); st.markdown("</div>", unsafe_allow_html=True)
            if st.form_submit_button("💾 حفظ الحكم", use_container_width=True, type="primary"):
                if not منطوق_الحكم: st.error("❌ لازم تكتب منطوق الحكم")
                else: case['حالة'] = 'منتهية'; case['تاريخ_الحكم'] = str(تاريخ_حكم); case['منطوق_الحكم'] = منطوق_الحكم; case['مسندة_ل_الحكم'] = مسندة_ل; case['سبب_الحفظ'] = سبب_الحفظ
                case['جلسات'].append({'تاريخ':str(تاريخ_حكم),'الرول':'-','الاجراء':f'الحكم - مسندة لـ {مسندة_ل}','ملاحظات':منطوق_الحكم}); case['تاريخ_جلسة'] = str(تاريخ_حكم); case['الاجراء'] = f'الحكم - مسندة لـ {مسندة_ل}'; save_data(data); st.success(f"✅ تم نقلها للارشيف للمتابعة"); st.session_state.page = "الأرشيف"; st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    # 6- الطباعة والتحميل
    st.markdown("<div style='background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #D4AF37; margin-bottom:15px; text-align:center'>", unsafe_allow_html=True)
    st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; margin-bottom:10px'>🖨️ الطباعة والتقرير</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🖨️ معاينة للطباعة", use_container_width=True, type="primary"):
            html_report = print_case_report(case)
            st.components.v1.html(html_report, height=800, scrolling=True)
            st.success("✅ اضغط Ctrl+P للطباعة")
    with col2:
        html_report = print_case_report(case)
        st.download_button(label="📥 تحميل التقرير",data=html_report.encode('utf-8'),file_name=f"تقرير_قضية_{case.get('رقم')}_{case.get('سنة')}.html",mime="text/html",use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # 7- حذف نهائى
    st.markdown("<style>.delete-box {background:#1E2A47; padding:15px; border-radius:15px; border:2px solid #FF0000; margin-bottom:15px; text-align:center}</style>", unsafe_allow_html=True)
# =============================================
# =============================================
# =============================================
# ======= الجزء الخامس: الارشيف =======
elif st.session_state.page == "الارشيف":
    data = load_data()
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#FFFFFF; text-align:center'>📁 أرشيف الأحكام النهائية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True): st.session_state.page = "الرئيسية"; st.rerun()

    ended_cases = [c for c in data["cases"] if c.get('حالة') == 'منتهية']

    # ======= تفعيل البحث في الارشيف =======
    search_arch = st.text_input("🔍 ابحث برقم القضية او الخصوم", key="search_arch")
    if search_arch:
        ended_cases = [c for c in ended_cases if search_arch in str(c.get('رقم','')) or search_arch in str(c.get('مدعي','')) or search_arch in str(c.get('مدعي_عليه',''))]
    # ======================================

    st.markdown(f"<div style='background:#1E2A47; padding:20px; border-radius:15px; border:2px solid #D4AF37; text-align:center; margin-bottom:20px'>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:28px; font-weight:900; color:#D4AF37'>📁 {len(ended_cases)}</div><div style='font-size:18px; color:#FFF; font-weight:700'>عدد الأحكام النهائية</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # ======= ستايل الجدول ابيض =======
    st.markdown("""
    <style>
.case-table {width:100%; border-collapse: collapse; font-size:11px; color:white; text-align:center; margin-bottom:5px;}
.case-table th {background:#D4AF37; color:#0B1426; padding:6px; font-weight:900;}
.case-table td {background:#1E2A47; padding:6px; border:1px solid #D4AF37; vertical-align:top; color:#FFFFFF;}
.plaintiff {background:#FFF3CD; color:#000; font-weight:700; border-radius:6px; padding:6px; font-size:11px;}
.plaintiff-hey2a {background:#DC3545!important; color:#FFF!important; font-weight:900; border-radius:6px; padding:6px; font-size:11px;}
.defendant {background:#CFF4FC; color:#000; font-weight:700; border-radius:6px; padding:6px; font-size:11px;}
.date-gold {color:#FFD700; font-weight:900;}
.status-green {color:#4CAF50; font-weight:900;}
    </style>
    """, unsafe_allow_html=True)
    # ==================================

    if not ended_cases: st.info("لا توجد احكام نهائية")
    else:
        sorted_ended = sorted(ended_cases, key=lambda x: x.get("تاريخ_جلسة","9999-12-31"), reverse=True)
        for idx, case in enumerate(sorted_ended, 1):
            رقم_كامل = f"{case.get('رقم','')} لسنة {case.get('سنة','')}"
            محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
            if case.get('مأمورية',''): محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = f"{case.get('دائرة', '' )}" if case.get('دائرة', '' ) else ""
            if دائرة_كاملة: محكمة_كاملة += f"<br>دائرة {دائرة_كاملة}"

            نوع = case.get('نوع','')
            if نوع == "استئناف": لقب1, لقب2 = "المستأنف:", "المستأنف ضده:"
            elif نوع == "طعن": لقب1, لقب2 = "الطاعن:", "المطعون ضده:"
            else: لقب1, لقب2 = "المدعى:", "المدعى عليه:"

            if "الهيئة" in str(case.get('مدعي','')):
                طرف1_html = f"<div class='plaintiff-hey2a'><b>{لقب1}</b><br>{case.get('مدعي','')}</div>"
            else:
                طرف1_html = f"<div class='plaintiff'><b>{لقب1}</b><br>{case.get('مدعي','')}</div>"
            طرف2_html = f"<div class='defendant'><b>{لقب2}</b><br>{case.get('مدعي_عليه','')}</div>"
            خصوم = طرف1_html + "<div style='height:4px'></div>" + طرف2_html

            # نجيب اخر جلسة بس للارشيف
            منطوق = ""
            تاريخ_حكم = ""
            لصالح = ""
            if "جلسات" in case and case["جلسات"]:
                اخر_جلسة = sorted(case["جلسات"], key=lambda x: x.get("تاريخ",""), reverse=True)[0]
                تاريخ_حكم = اخر_جلسة.get('تاريخ','')
                منطوق = اخر_جلسة.get('الاجراء','')
                لصالح = اخر_جلسة.get('الحالة','')

            table_html = "<table class='case-table'><tr>"
            headers = ["م", "الرقم والسنة", "المحكمة والدائرة", "الخصوم", "الموضوع", "تاريخ الحكم", "المنطوق", "لـ"]
            for h in headers: table_html += f"<th>{h}</th>"
            table_html += "</tr>"
            table_html += f"<tr><td>{idx}</td><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td class='date-gold'>{تاريخ_حكم}</td><td>{منطوق}</td><td class='status-green'>{لصالح}</td></tr></table>"
            st.markdown(table_html, unsafe_allow_html=True)

            # ======= زر الفتح تحت الجدول =======
            c1, c2, c3 = st.columns([4,1,4])
            with c2:
                if st.button("فتح", key=f"open_arch_{case['id']}", use_container_width=True):
                    st.session_state.selected_case_id = case['id']; st.session_state.page = "تفاصيل"; st.rerun()
            # ===================================
            st.markdown("<div style='height:15px'></div>", unsafe_allow_html=True)
# ==============================================
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

st.set_page_config(page_title="ادارة القضايا", layout="wide")

DATA_FILE = "data.json"

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"cases": [], "library": []}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='تقرير')
    return output.getvalue()

def to_word(df, title, region, member, manager, general):
    doc = Document()
    doc.add_paragraph(f"ديوان عام منطقة: {region}").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(title).alignment = WD_ALIGN_PARAGRAPH.CENTER
    if not df.empty:
        table = doc.add_table(rows=1, cols=len(df.columns))
        hdr_cells = table.rows[0].cells
        for i, col in enumerate(df.columns):
            hdr_cells[i].text = str(col)
        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            for i, val in enumerate(row):
                row_cells[i].text = str(val)
    doc.add_paragraph(f"\nعضو الادارة: {member}")
    doc.add_paragraph(f"مدير الادارة: {manager}")
    doc.add_paragraph(f"مدير عام: {general}")
    f = BytesIO()
    doc.save(f)
    return f.getvalue()

def to_pdf(df, title, region, member, manager, general):
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    p.setFont("Helvetica", 10)
    p.drawString(50, 800, f"ديوان عام منطقة: {region}")
    p.drawString(50, 780, title)
    y = 750
    if not df.empty:
        headers = " | ".join(df.columns)
        p.drawString(50, y, headers)
        y -= 20
        for _, row in df.iterrows():
            p.drawString(50, y, " | ".join([str(x) for x in row]))
            y -= 20
            if y < 50:
                break
    p.save()
    return buffer.getvalue()


if st.session_state.page == "تقارير":
    data = load_data()
    all_cases = data.get("cases", [])

    active_cases = [c for c in all_cases if c.get('الحالة') == 'متداولة']
    archive_cases = [c for c in all_cases if c.get('الحالة') == 'منتهية']

    st.markdown("<h2 style='text-align:center;'>📑 مركز التقارير القضائية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", key="back_reports", use_container_width=True):
        st.session_state.page = "الرئيسية"; st.rerun()

    report_type = st.selectbox(
        "اختر نوع التقرير",
        ["1. بيان بجميع الدعاوى المتداولة","2. بيان بالدعاوى المتداولة حسب موضوع الدعوى","3. بيان بجميع الاحكام الصادرة للصالح وللضد",
         "4. بيان بالاحكام الصادرة للصالح","5. بيان بالاحكام الصادرة للضد","6. بيان بالاحكام الصادرة حسب موضوع الدعوى",
         "8. بيان عددى بالاحكام","7. بيان بالاحصائيات"],
        key="report_select"
    )

    st.info(f"المتداولة: {len(active_cases)} | المنتهية: {len(archive_cases)}")

    col1, col2, col3 = st.columns(3)
    with col1: member = st.text_input("عضو الادارة القانونية", key="member_rep")
    with col2: manager = st.text_input("مدير الادارة القانونية", key="manager_rep")
    with col3: general_manager = st.text_input("مدير عام الإدارات القانونية", key="general_rep")

    df_report = pd.DataFrame()
    report_title = ""
    region = st.text_input("ديوان عام منطقة", key="region_rep")

    if report_type.startswith("1") or report_type.startswith("2"):
        cases = active_cases
        col1, col2, col3 = st.columns(3)
        with col1: from_date = st.date_input("من الفترة", key="from_active")
        with col2: to_date = st.date_input("حتى الفترة", key="to_active")
        with col3: topic = st.text_input("موضوع الدعوى للفلترة", key="topic_active") if report_type.startswith("2") else ""

        if st.button("🔍 عرض التقرير", use_container_width=True, type="primary"):
            filtered = []
            for c in cases:
                if not c.get('تاريخ_جلسة'):
                    filtered.append(c)
                    continue
                try:
                    case_date = datetime.strptime(c['تاريخ_جلسة'], '%Y-%m-%d').date()
                    if from_date <= case_date <= to_date:
                        if not topic or topic in str(c.get('موضوع','')):
                            filtered.append(c)
                except:
                    filtered.append(c)

            df_report = pd.DataFrame([{
                "م": i+1, "رقم القضية": f"{c.get('رقم','')}/{c.get('سنة','')}", "المحكمة": c.get('محكمة_اسم',''),
                "المدعي": c.get('مدعي',''), "المدعي عليه": c.get('مدعي_عليه',''),
                "الموضوع": c.get('موضوع',''), "تاريخ الجلسة": c.get('تاريخ_جلسة',''),
                "السبب": c.get('سبب','')
            } for i,c in enumerate(filtered)])
            report_title = "بيان بالدعاوى المتداولة"

    elif report_type.startswith("3") or report_type.startswith("4") or report_type.startswith("5") or report_type.startswith("6") or report_type.startswith("8"):
        cases = [c for c in archive_cases if c.get('مسندة_الى_الحكم') in ['للصالح', 'للضد']]
        col1, col2 = st.columns(2)
        with col1: from_date = st.date_input("من الفترة", key="from_archive")
        with col2: to_date = st.date_input("حتى الفترة", key="to_archive")
        topic = st.text_input("موضوع الدعوى للفلترة", key="topic_archive") if report_type.startswith("6") else ""

        if st.button("🔍 عرض التقرير", use_container_width=True, type="primary"):
            filtered = []
            for c in cases:
                if not c.get('تاريخ_الحكم'):
                    continue
                try:
                    case_date = datetime.strptime(c['تاريخ_الحكم'], '%Y-%m-%d').date()
                    if from_date <= case_date <= to_date:
                        if report_type.startswith("4") and c.get('مسندة_الى_الحكم')!= 'للصالح': continue
                        if report_type.startswith("5") and c.get('مسندة_الى_الحكم')!= 'للضد': continue
                        if topic and topic not in str(c.get('موضوع','')): continue
                        filtered.append(c)
                except:
                    pass

            if report_type.startswith("8"):
                total_ahkam = len(filtered)
                saleh = len([c for c in filtered if c.get('مسندة_الى_الحكم') == 'للصالح'])
                ded = len([c for c in filtered if c.get('مسندة_الى_الحكم') == 'للضد'])
                df_report = pd.DataFrame({
                    "البيان": ["اجمالي الاحكام", "احكام للصالح", "احكام للضد"],
                    "العدد": [total_ahkam, saleh, ded]
                })
            else:
                df_report = pd.DataFrame([{
                    "م": i+1, "رقم القضية": f"{c.get('رقم','')}/{c.get('سنة','')}", "الحكم": c.get('مسندة_الى_الحكم',''),
                    "الموضوع": c.get('موضوع',''), "تاريخ الحكم": c.get('تاريخ_الحكم',''), "المحكمة": c.get('محكمة_اسم','')
                } for i,c in enumerate(filtered)])
            report_title = "بيان بالاحكام"

    elif report_type.startswith("7"):
        st.markdown("<h3 style='text-align:center;'>📊 الاحصائيات العامة</h3>", unsafe_allow_html=True)
        total_active = len(active_cases)
        total_archive = len([c for c in archive_cases if c.get('مسندة_الى_الحكم') in ['للصالح', 'للضد']])
        saleh = len([c for c in archive_cases if c.get('مسندة_الى_الحكم') == 'للصالح'])
        ded = len([c for c in archive_cases if c.get('مسندة_الى_الحكم') == 'للضد'])
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("المتداولة", total_active)
        with col2: st.metric("الاحكام", total_archive)
        with col3: st.metric("للصالح", saleh)
        with col4: st.metric("للضد", ded)

    if not df_report.empty:
        st.success(f"تم العثور على {len(df_report)} سجل")
        st.dataframe(df_report, use_container_width=True)

        c1, c2, c3, c4 = st.columns(4)
        with c1: st.download_button("⬇️ Excel", data=to_excel(df_report), file_name="تقرير.xlsx", use_container_width=True)
        with c2: st.download_button("📄 Word", data=to_word(df_report, report_title, region, member, manager, general_manager), file_name="تقرير.docx", use_container_width=True)
        with c3: st.download_button("📕 PDF", data=to_pdf(df_report, report_title, region, member, manager, general_manager), file_name="تقرير.pdf", use_container_width=True)
        with c4: st.download_button("🖨️ HTML", data=df_report.to_html(index=False).encode('utf-8-sig'), file_name="تقرير.html", use_container_width=True)
    else:
        if not report_type.startswith("7"):
            st.warning("لا توجد بيانات للفترة المختارة")
