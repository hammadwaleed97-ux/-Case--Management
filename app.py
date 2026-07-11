# ================== إدارة القضايا v5.37 ====================
# ========== الإدارة العامة للشئون القانونية البحيرة ==========
# ============================================================
import streamlit as st
import pandas as pd
import json
import os
import smtplib
import secrets
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

DATA_FILE = "cases_data.json"
UPLOAD_FOLDER = "uploads"
TOKENS_FILE = "tokens.json"
if not os.path.exists(UPLOAD_FOLDER): os.makedirs(UPLOAD_FOLDER)

# ============= حط بياناتك هنا بالاحمر فقط =============
SENDER_EMAIL = "hammadwaleed97@gmail.com" # <--- احمر
SENDER_PASSWORD = "r v y q q a y j o n w h u o x r" # <--- احمر
APP_URL = "https://qpyqpsmkqcvdou4imbfunp.streamlit.app/" # ده بتاعك
# ==================================================
# ============================================
# الجزء 1: الاستيرادات والدوال الاساسية
# ============================================
ANWA3_MOSTANDAT = ["صحيفة دعوى", "صحيفة استئناف", "صحيفة طعن", "مذكرة دفاع", "حافظة مستندات", "تقرير خبير", "تقرير طب شرعى", "تقرير لجنة طبية", "صحيفة تجديد من الشطب", "صحيفة تعجيل من الوقف", "صورة حكم تمهيدى", "أخرى"]

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return {"cases": []}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=4)

def load_tokens():
    if os.path.exists(TOKENS_FILE):
        with open(TOKENS_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return {"tokens": []}

def save_tokens(tokens_data):
    with open(TOKENS_FILE, "w", encoding="utf-8") as f: json.dump(tokens_data, f, ensure_ascii=False, indent=4)

def send_verification_email(recipient_email, token):
    verify_link = f"{APP_URL}?verify_token={token}"
    subject = "تفعيل تنبيهات القضايا - الشئون القانونية البحيرة"
    body = f"مرحبا,\n\nلقد قمت بالتسجيل لتلقي تنبيهات الجلسات.\nمن فضلك فعل الاشتراك بالضغط على الرابط التالي:\n{verify_link}\n\nالرابط صالح لمدة 24 ساعة."
    msg = MIMEMultipart(); msg["From"] = SENDER_EMAIL; msg["To"] = recipient_email; msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    try: server = smtplib.SMTP("smtp.gmail.com", 587); server.starttls(); server.login(SENDER_EMAIL, SENDER_PASSWORD); server.sendmail(SENDER_EMAIL, recipient_email, msg.as_string()); server.quit(); return True
    except Exception as e: st.error(f"خطأ في ارسال الايميل: {e}"); return False

def send_session_alert_email(case_info, session_date, session_reason):
    tokens_data = load_tokens()
    verified_emails = [t['email'] for t in tokens_data['tokens'] if t['verified']]
    if not verified_emails: return
    subject = f"تنبيه جلسة جديدة - قضية رقم {case_info['رقم']} لسنة {case_info['سنة']}"
    body = f"مرحبا,\n\nتم اضافة جلسة جديدة للقضية:\n\nالرقم: {case_info['رقم']} لسنة {case_info['سنة']}\nالمحكمة: {case_info['محكمة_اسم']}\nالخصوم: {case_info['مدعي']} ضد {case_info['مدعي_عليه']}\n\nالتاريخ: {session_date}\nالسبب: {session_reason}\n\nمع تحيات\nالادارة العامة للشئون القانونية - البحيرة\n"
    for recipient_email in verified_emails:
        msg = MIMEMultipart(); msg["From"] = SENDER_EMAIL; msg["To"] = recipient_email; msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))
        try: server = smtplib.SMTP("smtp.gmail.com", 587); server.starttls(); server.login(SENDER_EMAIL, SENDER_PASSWORD); server.sendmail(SENDER_EMAIL, recipient_email, msg.as_string()); server.quit()
        except Exception as e: st.error(f"خطأ في ارسال للايميل {recipient_email}: {e}")

def verify_token(token):
    tokens_data = load_tokens(); now = datetime.now()
    for t in tokens_data["tokens"]:
        if t["token"] == token and datetime.strptime(t["expires"], "%Y-%m-%d %H:%M:%S") > now:
            if not t["verified"]: t["verified"] = True; save_tokens(tokens_data); return t["email"]
    return None
# ============================================
# نهاية الجزء 1
# ============================================
