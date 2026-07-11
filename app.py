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
