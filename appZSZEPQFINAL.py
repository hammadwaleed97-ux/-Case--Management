from pathlib import Path

import json, os, bcrypt, smtplib, random, io, requests, base64, html, secrets, re
from datetime import datetime, timedelta, date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import streamlit as st
import pandas as pd

from fpdf import FPDF
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import arabic_reshaper
from bidi.algorithm import get_display
from openpyxl.styles import Font, Alignment, PatternFill


# ===== دوال توحيد وتنسيق التواريخ (يوم . شهر . سنة / يوم / شهر / سنة) =====
def parse_date_safe(val):
    if not val:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    if not s:
        return None
    s_date = s.split("T")[0].split(" ")[0].strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(s_date, fmt).date()
        except Exception:
            continue
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        return None


def fmt_date_ar(val, default="", sep=" . "):
    d = parse_date_safe(val)
    if d is None:
        return str(val) if val not in (None, "") else default
    return f"{d.day:02d}{sep}{d.month:02d}{sep}{d.year:04d}"


def fmt_date_report(val, default=""):
    return fmt_date_ar(val, default=default, sep=" . ")


def date_sort_key(val, fallback="9999-12-31"):
    d = parse_date_safe(val)
    if d is None:
        return fallback
    return d.strftime("%Y-%m-%d")


def _clamp_day_for_ym(year, month, day):
    import calendar
    max_d = calendar.monthrange(int(year), int(month))[1]
    return min(max(int(day), 1), max_d)


def _norm_ar(text):
    s = str(text or "").strip().lower()
    s = re.sub(r"[إأآا]", "ا", s)
    s = s.replace("ى", "ي").replace("ة", "ه").replace("ـ", "")
    s = re.sub(r"\s+", " ", s)
    return s


def _case_session_dates(case):
    dates = []
    d_main = parse_date_safe(case.get("تاريخ_جلسة"))
    if d_main is not None:
        dates.append(d_main)
    for g in case.get("جلسات") or []:
        if isinstance(g, dict):
            d_g = parse_date_safe(g.get("تاريخ"))
            if d_g is not None:
                dates.append(d_g)
    return dates


def date_input_dmy(label, value=None, min_value=None, max_value=None, key=None, label_visibility="visible"):
    base_key = key or f"dmy_{re.sub(r'[^A-Za-z0-9_]+', '_', str(label))}"
    today = datetime.now().date()
    d_val = parse_date_safe(value) or today
    d_min = parse_date_safe(min_value) or date(1900, 1, 1)
    d_max = parse_date_safe(max_value) or date(2100, 12, 31)
    if d_min > d_max:
        d_min, d_max = d_max, d_min

    extra_max_key = f"{base_key}_extra_max_year"
    extra_max_year = st.session_state.get(extra_max_key, d_max.year)
    effective_max_year = max(d_max.year, extra_max_year, d_val.year)
    effective_min_year = min(d_min.year, d_val.year)

    if label_visibility != "collapsed":
        st.markdown(f"<div style='color:#FFD700;font-weight:900;font-size:14px;margin-bottom:4px;'>{html.escape(str(label))}</div>", unsafe_allow_html=True)

    years = list(range(effective_min_year, effective_max_year + 1))
    months = list(range(1, 13))
    days = list(range(1, 32))
    safe_day_idx = min(max(d_val.day - 1, 0), 30)

    c_d, c_m, c_y = st.columns([1, 1, 1.3], gap="small")
    with c_d:
        sel_d = st.selectbox(
            "اليوم",
            days,
            format_func=lambda d: f"{d:02d}",
            index=safe_day_idx,
            key=f"{base_key}_d",
            label_visibility="collapsed",
        )
    with c_m:
        sel_m = st.selectbox(
            "الشهر",
            months,
            format_func=lambda m: f"{m:02d}",
            index=months.index(d_val.month),
            key=f"{base_key}_m",
            label_visibility="collapsed",
        )
    with c_y:
        sel_y = st.selectbox(
            "السنة",
            years,
            index=years.index(d_val.year),
            key=f"{base_key}_y",
            label_visibility="collapsed",
        )

    final_d = date(int(sel_y), int(sel_m), _clamp_day_for_ym(sel_y, sel_m, sel_d))
    return fmt_date_ar(final_d)


# ===== التصميم القضائي الجديد =====
st.markdown("""
<style>
* {
    box-sizing: border-box;
}
/* وضوح الاختيارات في التبويبات والراديو والـ checkbox */
div[data-testid="stTabs"] button,
div[data-testid="stTabs"] button p,
div[data-testid="stTabs"] button span {
    color: #FFFFFF !important;
    font-weight: 900 !important;
    font-size: 14px !important;
    opacity: 1 !important;
    visibility: visible !important;
    white-space: nowrap !important;
}
div[data-testid="stTabs"] button[aria-selected="true"],
div[data-testid="stTabs"] button[aria-selected="true"] p,
div[data-testid="stTabs"] button[aria-selected="true"] span {
    color: #FFD700 !important;
    border-bottom: 3px solid #FFD700 !important;
    background: rgba(255, 215, 0, 0.14) !important;
}
div[data-testid="stRadio"] label,
div[data-testid="stRadio"] label p,
div[data-testid="stRadio"] label span,
div[data-testid="stRadio"] div[role="radiogroup"] label {
    color: #FFFFFF !important;
    font-weight: 900 !important;
    font-size: 14px !important;
    opacity: 1 !important;
    visibility: visible !important;
    white-space: normal !important;
    word-break: normal !important;
    overflow-wrap: break-word !important;
}
div[data-testid="stCheckbox"] label,
div[data-testid="stCheckbox"] label p,
div[data-testid="stCheckbox"] label span {
    color: #FFFFFF !important;
    font-weight: 900 !important;
    font-size: 14px !important;
    opacity: 1 !important;
    visibility: visible !important;
    white-space: normal !important;
    word-break: normal !important;
    overflow-wrap: break-word !important;
}
div[data-testid="stRadio"] label[data-checked="true"],
div[data-testid="stRadio"] label[data-checked="true"] p,
div[data-testid="stRadio"] label[data-checked="true"] span {
    color: #FFD700 !important;
}

/* صفحة الدخول */
.login-shell{
    max-width:760px;margin:45px auto;padding:28px;
    background:linear-gradient(145deg,#111c33,#1f3154);
    border:2px solid #D4AF37;border-radius:24px;
    box-shadow:0 20px 55px rgba(0,0,0,.45);
}
.login-shell h1,.login-shell h2,.login-shell p,.login-shell label{color:#fff!important;}
.login-shell input,.login-shell textarea{
    background:#fff!important;color:#111!important;
}
.admin-control-top{
    background:linear-gradient(135deg,#101b31,#1e2f50);
    border:2px solid #D4AF37;border-radius:20px;padding:14px;
    margin:8px 0 18px;box-shadow:0 10px 30px rgba(0,0,0,.28);
}
.judicial-card{
    min-height:118px;padding:20px 14px;border-radius:22px;
    border:2px solid #D4AF37;
    background:linear-gradient(145deg,#13213a,#213a61);
    box-shadow:0 12px 28px rgba(0,0,0,.30);
    text-align:center;margin-bottom:18px;
}
.judicial-card-title{color:#D4AF37;font-size:20px;font-weight:900;margin-top:8px;}
.judicial-card-icon{font-size:38px;line-height:1;}
.banner-admin-card{
    border:2px solid #D4AF37;border-radius:18px;padding:14px;
    background:#142038;margin-bottom:12px;
}
.status-active{color:#42d77d;font-weight:900}
.status-temp{color:#ffca28;font-weight:900}
.status-final{color:#ff5252;font-weight:900}
.incoming-messages-title{color:#D4AF37;font-size:24px;font-weight:900;text-align:center;margin:8px 0 14px;padding:10px;border-bottom:2px solid #D4AF37;}

/* إبراز الأيقونات والأزرار الملتصقة بخلفية الموقع */
button, [data-testid="stButton"] button, [data-testid="stDownloadButton"] button {
    color:#fff !important;
    font-weight:900 !important;
}
button svg, [data-testid="stButton"] svg, [data-testid="stDownloadButton"] svg, [data-testid="stFileUploader"] svg, [data-testid="stTabs"] svg, [data-testid="stExpander"] svg {
    color:#fff !important; fill:#fff !important; stroke:#fff !important; opacity:1 !important; visibility:visible !important;
}
button [data-testid="stIconMaterial"], button span, button p { opacity:1 !important; visibility:visible !important; color:#fff !important; fill:#fff !important; }
[data-testid="stButton"] button, [data-testid="stDownloadButton"] button, [data-testid="stFormSubmitButton"] button {
    color:#fff !important; background:#142038 !important; border:1px solid #D4AF37 !important;
}
[data-testid="stFileUploader"] section * { color:#fff !important; opacity:1 !important; visibility:visible !important; }
input, textarea, [data-baseweb="select"], [data-testid="stFileUploader"] {color:#111!important;}
label, [data-testid="stCheckbox"] label, [data-testid="stRadio"] label, [data-testid="stMultiSelect"] label, [data-testid="stSelectbox"] label {color:#fff!important;font-weight:900!important;}
[data-testid="stCheckbox"] p, [data-testid="stRadio"] p {color:#fff!important;font-weight:800!important;}
[data-testid="stFileUploader"] section {background:#142038!important;border:1px solid #D4AF37!important;}
[data-testid="stButton"] button svg, [data-testid="stDownloadButton"] button svg, [data-testid="stTabs"] button svg, [data-testid="stExpander"] svg, [data-testid="stFileUploader"] svg { display:inline-block !important; visibility:visible !important; opacity:1 !important; color:#fff !important; fill:#fff !important; stroke:#fff !important; }
[data-testid="stButton"] button > div, [data-testid="stDownloadButton"] button > div { display:flex !important; align-items:center !important; justify-content:center !important; gap:.45rem !important; }

[data-testid="stFileUploader"] small, [data-testid="stFileUploader"] span {color:#fff!important;}

/* توسيط الأزرار في الصفحة الرئيسية تحت بعضها بوضوح */
.home-sections-stack {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 12px;
    width: 100%;
    margin: 12px auto 20px auto;
}
</style>
""", unsafe_allow_html=True)

# ===== مزامنة السحابة والحفظ المحلي =====
BASE_DIR = Path(__file__).resolve().parent
LOCAL_DATA_DIR = str(BASE_DIR / "data")
os.makedirs(LOCAL_DATA_DIR, exist_ok=True)
LOCAL_BACKUP_DIR = BASE_DIR / "local_backups"
LOCAL_BACKUP_DIR.mkdir(parents=True, exist_ok=True)

DATA_DIR = Path(LOCAL_DATA_DIR)
DATA_DIR.mkdir(parents=True, exist_ok=True)
BANNERS_FILE = DATA_DIR / "banners.json"

def cloud_status():
    try:
        return bool(SUPABASE_URL and SUPABASE_KEY and supabase is not None)
    except Exception:
        return False

def cloud_success_message(action, section=""):
    target = f" في {section}" if section else ""
    return f"تم {action}{target} والسحابة بنجاح ✓"

def cloud_error_message(action, error):
    return f"تم {action} محليًا، ولكن تعذر إتمام العملية في السحابة: {error}"

def cloud_delete_record(table_name, record_id):
    try:
        if not cloud_status() or not record_id:
            return False, "اتصال السحابة غير متاح"
        supabase.table(table_name).delete().eq("id", record_id).execute()
        return True, ""
    except Exception as exc:
        return False, str(exc)

def cloud_upload_file(file_bytes, file_name, content_type="application/octet-stream"):
    try:
        if not cloud_status():
            return None, "اتصال السحابة غير متاح"
        safe_name = re.sub(r"[^A-Za-z0-9._/-]+", "_", str(file_name))
        storage_path = f"{datetime.now().strftime('%Y/%m')}/{safe_name}"
        supabase.storage.from_(SUPABASE_STORAGE_BUCKET).upload(
            storage_path, file_bytes,
            {"content-type": content_type, "upsert": "true"}
        )
        return storage_path, ""
    except Exception as exc:
        return None, str(exc)

def cloud_delete_file(storage_path):
    try:
        if not cloud_status() or not storage_path:
            return False, "مسار الملف السحابي غير موجود"
        supabase.storage.from_(SUPABASE_STORAGE_BUCKET).remove([storage_path])
        return True, ""
    except Exception as exc:
        return False, str(exc)

def fix_arabic(text):
    """ بتظبط العربي عشان ميطلعش متقطع """
    if not text:
        return ""
    reshaped_text = arabic_reshaper.reshape(str(text))
    bidi_text = get_display(reshaped_text)
    return bidi_text

def _secret(name, default=""):
    try:
        value = st.secrets.get(name, default)
        if value:
            return str(value)
    except Exception:
        pass
    return os.getenv(name, default)

SUPABASE_URL = _secret("SUPABASE_URL").rstrip("/")
SUPABASE_KEY = _secret("SUPABASE_KEY")
ADMIN_USERNAME = _secret("ADMIN_USERNAME", "admin")
ADMIN_DEFAULT_PASS = _secret("ADMIN_DEFAULT_PASS", "admin123")

SENDER_EMAIL = _secret("SENDER_EMAIL", "")
SENDER_PASSWORD = _secret("SENDER_PASSWORD", "")
APP_URL = _secret("APP_URL", "https://qpyqpsmkqcvdou4imbfunp.streamlit.app/")
SUPABASE_STORAGE_BUCKET = _secret("SUPABASE_STORAGE_BUCKET", "legal-files")
CLOUD_ENABLED = bool(SUPABASE_URL and SUPABASE_KEY)

_CLOUD_LAST_ERROR = ""

def _cloud_headers():
    return {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "resolution=merge-duplicates,return=minimal",
    }

def _cloud_put_db(key, payload):
    global _CLOUD_LAST_ERROR
    try:
        r = requests.post(
            f"{SUPABASE_URL}/rest/v1/app_state",
            params={"on_conflict": "key"},
            headers=_cloud_headers(),
            json={
                "key": key,
                "payload": payload,
                "updated_at": datetime.now().astimezone().isoformat(),
            },
            timeout=15,
        )
        if r.ok:
            return True
        _CLOUD_LAST_ERROR = f"Database HTTP {r.status_code}: {r.text[:500]}"
    except Exception as exc:
        _CLOUD_LAST_ERROR = f"Database: {exc}"
    return False

def _cloud_get_db(key):
    global _CLOUD_LAST_ERROR
    try:
        r = requests.get(
            f"{SUPABASE_URL}/rest/v1/app_state",
            params={"key": f"eq.{key}", "select": "key,payload,updated_at"},
            headers=_cloud_headers(),
            timeout=12,
        )
        if r.ok:
            rows = r.json()
            return rows[0] if rows else None
        _CLOUD_LAST_ERROR = f"Database HTTP {r.status_code}: {r.text[:500]}"
    except Exception as exc:
        _CLOUD_LAST_ERROR = f"Database: {exc}"
    return None

def _storage_headers(content_type="application/json"):
    return {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": content_type,
        "x-upsert": "true",
    }

def _cloud_put_storage(key, payload):
    global _CLOUD_LAST_ERROR
    try:
        storage_key = f"app_state/{key}.json"
        raw = json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")
        r = requests.post(
            f"{SUPABASE_URL}/storage/v1/object/{SUPABASE_STORAGE_BUCKET}/{storage_key}",
            headers=_storage_headers(),
            data=raw,
            timeout=20,
        )
        if r.ok:
            return True
        _CLOUD_LAST_ERROR = f"Storage HTTP {r.status_code}: {r.text[:500]}"
    except Exception as exc:
        _CLOUD_LAST_ERROR = f"Storage: {exc}"
    return False

def _cloud_get_storage(key):
    global _CLOUD_LAST_ERROR
    try:
        storage_key = f"app_state/{key}.json"
        r = requests.get(
            f"{SUPABASE_URL}/storage/v1/object/{SUPABASE_STORAGE_BUCKET}/{storage_key}",
            headers={
                "apikey": SUPABASE_KEY,
                "Authorization": f"Bearer {SUPABASE_KEY}",
            },
            timeout=15,
        )
        if r.ok:
            return {"key": key, "payload": r.json()}
        if r.status_code != 404:
            _CLOUD_LAST_ERROR = f"Storage HTTP {r.status_code}: {r.text[:500]}"
    except Exception as exc:
        _CLOUD_LAST_ERROR = f"Storage: {exc}"
    return None

def _cloud_put(key, payload):
    global _CLOUD_LAST_ERROR
    if not CLOUD_ENABLED:
        _CLOUD_LAST_ERROR = "SUPABASE_URL أو SUPABASE_KEY غير موجود في Secrets."
        return False
    if _cloud_put_db(key, payload):
        return True
    return _cloud_put_storage(key, payload)

def _cloud_get(key):
    global _CLOUD_LAST_ERROR
    if not CLOUD_ENABLED:
        _CLOUD_LAST_ERROR = "SUPABASE_URL أو SUPABASE_KEY غير موجود في Secrets."
        return None
    row = _cloud_get_db(key)
    if row is not None:
        return row
    return _cloud_get_storage(key)

def _write_local_json_atomic(file_path, data):
    """كتابة آمنة محلياً عبر ملف مؤقت ثم استبدال ذري + نسخة احتياطية محلية."""
    path = Path(file_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    serialized = json.dumps(data, ensure_ascii=False, indent=2, default=str)
    tmp_path = path.with_name(f".{path.name}.{secrets.token_hex(4)}.tmp")
    tmp_path.write_text(serialized, encoding="utf-8")
    os.replace(str(tmp_path), str(path))

    backup_dir = LOCAL_BACKUP_DIR / path.stem
    backup_dir.mkdir(parents=True, exist_ok=True)
    latest_backup = backup_dir / f"{path.name}.bak"
    latest_backup.write_text(serialized, encoding="utf-8")
    return True

def _read_local_json_safe(file_path):
    """قراءة آمنة من الملف المحلي مع الرجوع للنسخة الاحتياطية المحلية عند الحاجة."""
    path = Path(file_path)
    for candidate in (path, LOCAL_BACKUP_DIR / path.stem / f"{path.name}.bak"):
        try:
            if candidate.exists():
                return json.loads(candidate.read_text(encoding="utf-8"))
        except Exception:
            continue
    return None

def load_json_persistent(key, local_path, default):
    local_obj = _read_local_json_safe(local_path)
    cloud = _cloud_get(key)
    if cloud and cloud.get("payload") is not None:
        try:
            cloud_obj = cloud["payload"]
            if local_obj is None:
                local_obj = cloud_obj
                _write_local_json_atomic(local_path, local_obj)
        except Exception:
            pass
    if local_obj is None:
        local_obj = default
        try:
            _write_local_json_atomic(local_path, local_obj)
        except Exception:
            pass
        if CLOUD_ENABLED:
            _cloud_put(key, local_obj)
    return local_obj

def save_json_persistent(key, file_path, data):
    """
    تعديل جزء الحفظ محلياً:
    يعتمد الحفظ المحلي أولاً وبشكل آمن ومستقل، ويحفظ نسخة احتياطية محلية،
    ثم يحاول المزامنة مع السحابة دون إفساد أو إبطال الحفظ المحلي إذا تعذرت السحابة.
    """
    try:
        local_ok = _write_local_json_atomic(file_path, data)
    except Exception as exc:
        st.error(f"❌ تعذر الحفظ المحلي: {exc}")
        return False

    try:
        _cached_users_snapshot.clear()
    except Exception:
        pass
    try:
        _cached_banners_snapshot.clear()
    except Exception:
        pass

    if not CLOUD_ENABLED:
        return bool(local_ok)

    try:
        cloud_ok = bool(_cloud_put(key, data))
    except Exception as exc:
        st.info(f"💾 تم الحفظ محليًا بنجاح (تعذرت مزامنة السحابة مؤقتًا: {exc})")
        return bool(local_ok)

    if not cloud_ok:
        st.info(f'💾 تم الحفظ محليًا بنجاح (لم تتم مزامنة السحابة: {_CLOUD_LAST_ERROR or "غير متاحة حاليًا"})')
        return bool(local_ok)

    return True

def load_banners():
    return load_json_persistent("banners", BANNERS_FILE, [])

def save_banners(banners):
    return save_json_persistent("banners", BANNERS_FILE, banners)

if "RESET_CODES" not in st.session_state:
    st.session_state.RESET_CODES = {}

def send_email(to_email, subject, body):
    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = subject
        msg["From"] = SENDER_EMAIL
        msg["To"] = to_email
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"خطأ في الارسال: {e}")
        return False

USERS_FILE = os.path.join(LOCAL_DATA_DIR, "users.json")

def load_users():
    admin_pass = bcrypt.hashpw(ADMIN_DEFAULT_PASS.encode(), bcrypt.gensalt()).decode()
    default_users = [{"id":1,"username":ADMIN_USERNAME,"display_name":ADMIN_USERNAME,"password":admin_pass,"email":SENDER_EMAIL,"recovery_email":"","role":"admin","status":"active","password_set":True,"job_title":"","insurance_area":""}]
    return load_json_persistent("users", USERS_FILE, default_users)

def save_users(users):
    return save_json_persistent("users", USERS_FILE, users)

@st.cache_data(ttl=20, show_spinner=False)
def _cached_users_snapshot():
    return load_users()

@st.cache_data(ttl=20, show_spinner=False)
def _cached_banners_snapshot():
    return load_banners()

def check_login(username, password):
    users = _cached_users_snapshot()
    changed = False
    for u in users:
        if not u.get("display_name"):
            u["display_name"] = u.get("username", "")
            changed = True
    if changed:
        save_users(users)
    for user in users:
        typed = str(username or "").strip().casefold()
        login_username = str(user.get("username", "")).strip().casefold()
        display_name = str(user.get("display_name", "")).strip().casefold()
        if typed not in (login_username, display_name):
            continue
        if user.get("status", "active") != "active":
            return None
        stored = user.get("password") or ""
        if not stored:
            return None
        try:
            if bcrypt.checkpw(password.encode(), stored.encode()):
                return user
        except Exception:
            return None
    return None

def is_admin_email(email):
    users = _cached_users_snapshot()
    admin = next((u for u in users if u["role"] == "admin"), None)
    if not admin:
        return False
    return email == admin["email"] or email == admin.get("recovery_email","")

def normalize_person_name(value):
    text = str(value or "").strip().casefold().replace("ـ", "")
    return re.sub(r"\s+", " ", text)

def find_user_for_activation(users, typed_name):
    key = normalize_person_name(typed_name)
    if not key:
        return None
    for u in users:
        if normalize_person_name(u.get("username")) == key or normalize_person_name(u.get("display_name")) == key:
            return u
    compact = key.replace(" ", "")
    for u in users:
        for field in ("username", "display_name"):
            candidate = normalize_person_name(u.get(field)).replace(" ", "")
            if candidate and candidate == compact:
                return u
    return None

def membership_job_title(user):
    job = str(user.get("job") or user.get("وظيفة") or user.get("role_name") or "")
    return "مدير عام (قيادى مناطق)" if job.strip() == "مدير عام" else job

def login_page():
    st.markdown("""
    <div class="login-shell">
      <h1 style="text-align:center;color:#D4AF37!important;">⚖️ إدارة القضايا دخول السادة الأعضاء ⚖️</h1>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["🔐 تسجيل الدخول", "🪪 تفعيل حساب جديد"])

    with tab1:
        st.markdown("<div class='login-shell' style='margin-top:0;'>", unsafe_allow_html=True)
        st.markdown("<p style='color:white!important;font-weight:900;'>اسم المستخدم</p>", unsafe_allow_html=True)
        username = st.text_input("اسم المستخدم", key="login_user", label_visibility="collapsed")
        st.markdown("<p style='color:white!important;font-weight:900;'>كلمة السر</p>", unsafe_allow_html=True)
        password = st.text_input("كلمة المرور", type="password", key="login_pass", label_visibility="collapsed")

        if st.button("دخول", type="primary", use_container_width=True, key="login_btn"):
            user = check_login(username.strip(), password)
            if user:
                st.session_state.user = user
                st.session_state.role = user["role"]
                if user["role"] == "member" and not user.get("password_set", False):
                    st.session_state.page = "set_password"
                    st.session_state.temp_user_id = user.get("id")
                    st.session_state.temp_user = user.get("username")
                else:
                    st.session_state.page = "الرئيسية"
                st.rerun()
            else:
                st.error("اسم المستخدم أو كلمة السر غير صحيحة أو العضوية موقوفة")

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("<p style='color:white!important;font-weight:900;'>نسيت بياناتك؟ استرجعها بالإيميل</p>", unsafe_allow_html=True)

        admin_recover_email = st.text_input("إيميل الادمن", key="admin_recover")
        if st.button("إرسال كود للادمن", key="admin_send", use_container_width=True):
            if is_admin_email(admin_recover_email):
                code = str(random.randint(100000, 999999))
                st.session_state.RESET_CODES[admin_recover_email] = {"code": code, "role": "admin"}
                body = f"كود إعادة تعيين كلمة سر الادمن: {code}"
                if send_email(admin_recover_email, "كود استرجاع الادمن", body):
                    st.success("تم إرسال الكود على الإيميل")
                    st.session_state.show_reset_admin = True
            else:
                st.error("هذا الإيميل غير مسجل كادمن")

        member_recover_email = st.text_input("إيميل العضو", key="member_recover")
        if st.button("إرسال كود للعضو", key="member_send", use_container_width=True):
            users = _cached_users_snapshot()
            found = [u for u in users if u.get("email") == member_recover_email]
            if found:
                user = found[0]
                code = str(random.randint(100000, 999999))
                st.session_state.RESET_CODES[member_recover_email] = {"code": code, "user_id": user["id"]}
                body = f"مرحبا {user['username']}\nاسم المستخدم: {user['username']}\nكود إعادة التعيين: {code}"
                if send_email(member_recover_email, "استرجاع بيانات الدخول", body):
                    st.success("تم إرسال الكود على إيميلك")
                    st.session_state.show_reset_member = True
            else:
                st.error("الإيميل ده مش متسجل")

        if st.session_state.get("show_reset_admin") or st.session_state.get("show_reset_member"):
            email_to_reset = admin_recover_email if st.session_state.get("show_reset_admin") else member_recover_email
            code_input = st.text_input("ادخل الكود")
            new_pass = st.text_input("كلمة السر الجديدة", type="password")
            if st.button("تأكيد وتغيير كلمة السر", key="confirm_reset"):
                saved = st.session_state.RESET_CODES.get(email_to_reset, {})
                if saved.get("code") == code_input:
                    users = _cached_users_snapshot()
                    logged_user = None
                    if saved.get("role") == "admin":
                        admin = next((u for u in users if u["role"] == "admin"), None)
                        if admin:
                            hashed = bcrypt.hashpw(new_pass.encode(), bcrypt.gensalt()).decode()
                            admin["password"] = hashed
                            save_users(users)
                            logged_user = admin
                    else:
                        uid = saved.get("user_id")
                        for u in users:
                            if u["id"] == uid:
                                u["password"] = bcrypt.hashpw(new_pass.encode(), bcrypt.gensalt()).decode()
                                u["status"] = "active"
                                u["password_set"] = True
                                logged_user = u
                                break
                        save_users(users)

                    if logged_user:
                        st.session_state.RESET_CODES.clear()
                        st.session_state.show_reset_admin = False
                        st.session_state.show_reset_member = False
                        st.session_state.user = logged_user
                        st.session_state.role = logged_user["role"]
                        st.session_state.page = "الرئيسية"
                        st.success("تم تسجيل الدخول بنجاح")
                        st.rerun()
                else:
                    st.error("الكود غلط")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab2:
        st.markdown("<div class='login-shell' style='margin-top:0;'>", unsafe_allow_html=True)
        st.markdown("<h3 style='color:#D4AF37!important;text-align:center;'>تفعيل حساب جديد</h3>", unsafe_allow_html=True)
        member_name = st.text_input("اكتب الاسم الذي استخرجه لك الادمن", key="new_user")
        if st.button("🪪 تفعيل الحساب", use_container_width=True, type="primary", key="activate_account_v2"):
            users = _cached_users_snapshot()
            found_user = find_user_for_activation(users, member_name)
            if not member_name.strip():
                st.error("اكتب اسم العضو أولاً")
            elif not found_user:
                st.error("الاسم غير موجود ضمن العضويات المستخرجة. اكتب الاسم كما سجله الادمن عند استخراج العضوية.")
            elif found_user.get("role") == "admin":
                st.error("هذا حساب الادمن ولا يتم تفعيله من هنا.")
            elif found_user.get("status") in ("suspended_final", "deleted", "banned"):
                st.error("هذه العضوية موقوفة نهائياً ولا يمكن تفعيلها.")
            elif found_user.get("password_set") and found_user.get("password"):
                st.error("هذه العضوية مفعلة بالفعل. استخدم تسجيل الدخول أو استرجاع كلمة السر.")
            else:
                st.session_state.activation_user_id = found_user.get("id")
                st.session_state.activation_user_name = found_user.get("display_name") or found_user.get("username")
                st.session_state.show_activation_password = True
                st.rerun()

        if st.session_state.get("show_activation_password"):
            users = _cached_users_snapshot()
            activation_id = st.session_state.get("activation_user_id")
            target = next((u for u in users if u.get("id") == activation_id), None)
            if target:
                st.markdown(f"<div style='background:#14233d;border:2px solid #D4AF37;border-radius:14px;padding:14px;margin-top:14px;color:#fff;text-align:center;font-weight:900;'>تم العثور على العضوية: {html.escape(str(target.get('display_name') or target.get('username')))}<br>أدخل كلمة السر التي تريد استخدامها.</div>", unsafe_allow_html=True)
                first = st.text_input("كلمة السر", type="password", key="activation_pass_1")
                second = st.text_input("تأكيد كلمة السر", type="password", key="activation_pass_2")
                activation_photo = st.file_uploader("الصورة الشخصية", type=["jpg","jpeg","png"], key="activation_profile_photo")
                if st.button("💾 حفظ كلمة السر وتفعيل العضوية", type="primary", use_container_width=True, key="confirm_activation_v2"):
                    if not first:
                        st.error("اكتب كلمة السر أولاً")
                    elif first != second:
                        st.error("كلمتا السر غير متطابقتين")
                    else:
                        for u in users:
                            if u.get("id") == activation_id:
                                u["password"] = bcrypt.hashpw(first.encode(), bcrypt.gensalt()).decode()
                                u["password_set"] = True
                                if activation_photo:
                                    u["profile_photo"] = base64.b64encode(activation_photo.getvalue()).decode("utf-8")
                                    u["profile_photo_type"] = activation_photo.type
                                if u.get("status") in (None, "suspended_temp"):
                                    u["status"] = "active"
                                target = u
                                break
                        save_users(users)
                        st.session_state.user = target
                        st.session_state.role = target.get("role")
                        st.session_state.page = "الرئيسية"
                        st.session_state.pop("activation_user_id", None)
                        st.session_state.pop("activation_user_name", None)
                        st.session_state.show_activation_password = False
                        st.success("تم تفعيل العضوية وتسجيل الدخول بنجاح")
                        st.rerun()
            else:
                st.session_state.show_activation_password = False
                st.session_state.pop("activation_user_id", None)
                st.error("تعذر العثور على العضوية، اكتب الاسم مرة أخرى.")

        st.markdown("</div>", unsafe_allow_html=True)

def calculate_age(birth_date, on_date=None):
    """حساب العمر بدقة بالسنوات الكاملة."""
    parsed_b = parse_date_safe(birth_date)
    if not parsed_b:
        return None
    on_date = parse_date_safe(on_date) or datetime.now().date()
    return on_date.year - parsed_b.year - ((on_date.month, on_date.day) < (parsed_b.month, parsed_b.day))

def member_age_status(user):
    """يعيد حالة السن، مع مراعاة تاريخ مد العضوية إن وجد."""
    birth = user.get("birth_date")
    age = calculate_age(birth)
    extension = user.get("membership_extended_until")
    extended_active = False
    ext_date = parse_date_safe(extension)
    if ext_date is not None:
        extended_active = ext_date >= datetime.now().date()
    return age, extended_active

def extract_member_page():
    st.markdown("<h2 style='text-align:center; color:#C9A961'>🪪 استخراج عضوية جديدة</h2>", unsafe_allow_html=True)
    st.markdown("""
    <style>
    div[data-testid="stTextInput"] label, div[data-testid="stDateInput"] label,
    div[data-testid="stSelectbox"] label {color:white!important;font-weight:900!important;}
    </style>
    """, unsafe_allow_html=True)

    if st.button("⬅️ العودة للوحة التحكم", use_container_width=True):
        st.session_state.page = "لوحة التحكم"
        st.rerun()

    with st.container(border=True):
        st.markdown("<h3 style='color:#D4AF37;text-align:center;'>بيانات العضوية</h3>", unsafe_allow_html=True)
        new_username = st.text_input("الاسم", key="new_username")
        birth_date = date_input_dmy(
            "تاريخ الميلاد",
            value=datetime.now().date().replace(year=max(1900, datetime.now().year - 30)),
            min_value=date(1900, 1, 1),
            max_value=date(2100, 12, 31),
            key="new_member_birth_date"
        )
        insurance_area = st.text_input("المنطقة التأمينية التابع لها العضو", key="new_member_insurance_area")
        job_role = st.selectbox("الوظيفة", ["عضو قانونى", "مدير إدارة", "مدير عام", "قيادي مركز رئيسى 1", "قيادي مركز رئيسى 2", "قيادي مركز رئيسى 3"], key="new_member_job_role")
        profile_photo = st.file_uploader("الصورة الشخصية", type=["jpg","jpeg","png"], key="new_member_profile_photo")

        if st.button("🪪 استخراج العضوية", use_container_width=True, type="primary", key="extract_member_new_v3"):
            name = new_username.strip()
            area = insurance_area.strip()
            if not name:
                st.error("لازم تكتب اسم العضو")
            elif not area:
                st.error("لازم تكتب المنطقة التأمينية")
            else:
                users = _cached_users_snapshot()
                existing_user = next((u for u in users if u.get("username") == name), None)
                birth_date_str = fmt_date_ar(birth_date)
                if existing_user:
                    existing_user["birth_date"] = birth_date_str
                    existing_user["insurance_area"] = area
                    existing_user["job_title"] = job_role
                    if profile_photo:
                        existing_user["profile_photo"] = base64.b64encode(profile_photo.getvalue()).decode("utf-8")
                        existing_user["profile_photo_type"] = profile_photo.type
                    existing_user["role"] = {"عضو قانونى":"member", "مدير إدارة":"manager", "مدير عام":"general_manager", "قيادي مركز رئيسى 1":"leader1", "قيادي مركز رئيسى 2":"leader2", "قيادي مركز رئيسى 3":"leader3"}.get(job_role, "member")
                    existing_user.setdefault("membership_extended_until", None)
                    if existing_user.get("status") in ("banned", "deleted"):
                        existing_user["status"] = "active"
                    if not existing_user.get("password_set"):
                        existing_user["password"] = ""
                    save_users(users)
                    st.success(f"تم حفظ بيانات العضوية للعضو: {name}")
                else:
                    new_id = max([int(u.get('id', 0)) for u in users if str(u.get('id','')).isdigit()] or [0]) + 1
                    users.append({
                        "id": new_id,
                        "username": name,
                        "password": "",
                        "email": "",
                        "recovery_email": "",
                        "role": {"عضو قانونى":"member", "مدير إدارة":"manager", "مدير عام":"general_manager", "قيادي مركز رئيسى 1":"leader1", "قيادي مركز رئيسى 2":"leader2", "قيادي مركز رئيسى 3":"leader3"}.get(job_role, "member"),
                        "job_title": job_role,
                        "status": "active",
                        "password_set": False,
                        "birth_date": birth_date_str,
                        "insurance_area": area,
                        "membership_extended_until": None,
                        "age_action": None,
                        "age_action_at": None,
                        "transferred_to": None,
                        "profile_photo": base64.b64encode(profile_photo.getvalue()).decode("utf-8") if profile_photo else "",
                        "profile_photo_type": profile_photo.type if profile_photo else ""
                    })
                    save_users(users)
                    st.success(f"تم استخراج العضوية بنجاح: {name}")
                st.rerun()

def _transfer_member_data(source_user, target_user):
    """نقل القضايا والمكتبة للعضو الجديد مع الإبقاء على السجلات دون حذف."""
    data = load_data()
    source_name = source_user.get("username", "")
    target_name = target_user.get("username", "")
    target_id = target_user.get("id")

    changed = False
    for case in data.get("cases", []):
        if case.get("assigned_to") == source_name:
            case["assigned_to"] = target_name
            changed = True
        if case.get("user_id") == source_user.get("id"):
            case["user_id"] = target_id
            changed = True

    for item in data.get("library", []):
        if item.get("user_id") == source_user.get("id"):
            item["user_id"] = target_id
            changed = True

    if changed:
        save_data(data)
    return changed

def _age_action_dialog(user_id):
    """حوار إجراءات بلوغ الستين للعضو."""
    users = _cached_users_snapshot()
    user = next((u for u in users if u.get("id") == user_id), None)
    if not user:
        return

    age, extended_active = member_age_status(user)
    if age is None or age < 60 or extended_active:
        return

    @st.dialog("إجراءات العضوية لبلوغ سن الستين")
    def age_dialog():
        st.warning(f"⚠️ العضو **{user.get('username','')}** بلغ سن {age} سنة.")
        st.write(f"تاريخ الميلاد: {fmt_date_ar(user.get('birth_date'), 'غير مسجل')}")
        st.write(f"المنطقة التأمينية: {user.get('insurance_area','غير مسجلة')}")

        action = st.radio(
            "اختر الإجراء",
            [
                "إيقاف عضوية لبلوغ السن",
                "إيقاف عضوية لبلوغ السن مع نقل البيانات لعضو آخر",
                "مد العضوية لفترة"
            ],
            key=f"age_action_{user_id}"
        )

        if action == "إيقاف عضوية لبلوغ السن":
            if st.button("⛔ تنفيذ إيقاف العضوية", type="primary", use_container_width=True, key=f"age_stop_{user_id}"):
                user["status"] = "suspended_age"
                user["age_action"] = "إيقاف عضوية لبلوغ السن"
                user["age_action_at"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                save_users(users)
                st.success("تم إيقاف العضوية، ولم يتم حذف أي قضية أو مستند.")
                st.rerun()

        elif action == "إيقاف عضوية لبلوغ السن مع نقل البيانات لعضو آخر":
            targets = [u for u in users if u.get("role") != "admin" and u.get("id") != user.get("id") and u.get("status") == "active"]
            if not targets:
                st.error("لا يوجد عضو آخر نشط يمكن نقل البيانات إليه.")
            else:
                target_map = {u.get("username", ""): u for u in targets}
                target_name = st.selectbox("ادخل اسم العضو الآخر", list(target_map.keys()), key=f"age_target_{user_id}")
                st.info("سيتم نقل القضايا والمكتبة الخاصة بالعضو للعضو المختار، مع عدم حذف السجلات.")
                if st.button("🔄 نقل البيانات ثم إيقاف العضوية", type="primary", use_container_width=True, key=f"age_transfer_{user_id}"):
                    target = target_map[target_name]
                    _transfer_member_data(user, target)
                    user["status"] = "suspended_age"
                    user["age_action"] = "إيقاف عضوية لبلوغ السن مع نقل البيانات"
                    user["age_action_at"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                    user["transferred_to"] = target.get("username")
                    save_users(users)
                    st.success(f"تم نقل البيانات إلى {target.get('username')} ثم إيقاف العضوية.")
                    st.rerun()

        else:
            extension_date = date_input_dmy(
                "مد العضوية حتى تاريخ",
                min_value=datetime.now().date(),
                value=(datetime.now() + timedelta(days=365)).date(),
                max_value=date(2100, 12, 31),
                key=f"age_extension_{user_id}"
            )
            if st.button("📅 حفظ مد العضوية", type="primary", use_container_width=True, key=f"age_extend_{user_id}"):
                ext_str = fmt_date_ar(extension_date)
                user["status"] = "active"
                user["membership_extended_until"] = ext_str
                user["age_action"] = "مد العضوية"
                user["age_action_at"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                save_users(users)
                st.success(f"تم مد العضوية حتى {ext_str}")
                st.rerun()

    age_dialog()

def manage_users_page():
    if not st.session_state.user or st.session_state.user.get("role") != "admin":
        st.error("هذه الصفحة متاحة للادمن فقط")
        return

    st.markdown("<h2 style='color:#D4AF37;text-align:center;'>👥 إدارة الأعضاء</h2>", unsafe_allow_html=True)
    users = _cached_users_snapshot()
    members = [u for u in users if u.get("role") != "admin"]

    st.info("إيقاف العضو أو حذف حساب الدخول لا يحذف القضايا أو المستندات. بيانات العضو تظل محفوظة.")

    if st.button("✏️ تعديل بيانات العضويات", key="open_membership_editor", use_container_width=True):
        st.session_state.page = "تعديل_بيانات_العضويات"
        st.rerun()

    if not members:
        st.info("لا يوجد أعضاء حالياً.")
        return

    for user in members:
        birth = fmt_date_ar(user.get("birth_date"))
        age, extended_active = member_age_status(user)
        status = user.get("status", "active")
        if status == "active":
            status_text, cls = "نشط", "status-active"
        elif status == "suspended_temp":
            status_text, cls = "إيقاف مؤقت", "status-temp"
        elif status == "suspended_final":
            status_text, cls = "إيقاف نهائي", "status-final"
        elif status == "suspended_age":
            status_text, cls = "إيقاف لبلوغ السن", "status-final"
        else:
            status_text, cls = status, "status-final"

        with st.container(border=True):
            c1, c2, c3 = st.columns([4, 2, 4])
            with c1:
                if user.get("profile_photo"):
                    st.markdown(_photo_html(user.get("profile_photo"), user.get("profile_photo_type") or "image/png", 85), unsafe_allow_html=True)
                st.markdown(f"### 👤 {user.get('username', '')}")
                st.markdown(f"**الوظيفة:** {user.get('job_title') or role_label(user.get('role'))}")
                st.markdown(f"**تاريخ الميلاد:** {birth or 'غير مسجل'}")
                st.markdown(f"**المنطقة التأمينية:** {user.get('insurance_area') or 'غير مسجلة'}")
                if age is not None:
                    age_line = f"**العمر:** {age} سنة"
                    if extended_active:
                        age_line += f" — **العضوية ممددة حتى:** {fmt_date_ar(user.get('membership_extended_until'))}"
                    st.markdown(age_line)
                st.markdown(f"الإيميل: {user.get('email') or 'غير مسجل'}")
                st.markdown(f"الحالة: <span class='{cls}'>{status_text}</span>", unsafe_allow_html=True)

                if age is not None and age >= 60 and not extended_active and status not in ("suspended_age", "suspended_final"):
                    st.warning("⚠️ بلغ سن الستين — يلزم اتخاذ إجراء بشأن العضوية.")
                    if st.button("⚖️ إجراءات بلوغ سن الستين", key=f"age_dialog_btn_{user['id']}", use_container_width=True):
                        _age_action_dialog(user["id"])
                elif status == "suspended_age":
                    st.warning("⛔ العضوية موقوفة لبلوغ السن.")
                    if user.get("transferred_to"):
                        st.caption(f"تم نقل البيانات إلى: {user.get('transferred_to')}")

            with c2:
                if status != "active":
                    if st.button("✅ إعادة تفعيل", key=f"reactivate_{user['id']}", use_container_width=True):
                        user["status"] = "active"
                        save_users(users)
                        st.rerun()

                if age is not None and age >= 60 and not extended_active and status not in ("suspended_age", "suspended_final"):
                    if st.button("⚖️ إجراء السن", key=f"age_action_btn_{user['id']}", use_container_width=True):
                        _age_action_dialog(user["id"])

            with c3:
                a,b,c = st.columns(3)
                with a:
                    if status == "active":
                        if st.button("⏸️ إيقاف مؤقت", key=f"temp_{user['id']}", use_container_width=True):
                            user["status"] = "suspended_temp"
                            save_users(users)
                            st.rerun()
                with b:
                    if status != "suspended_final":
                        if st.button("⛔ إيقاف نهائي", key=f"final_{user['id']}", use_container_width=True):
                            user["status"] = "suspended_final"
                            user["password"] = ""
                            user["password_set"] = False
                            save_users(users)
                            st.rerun()
                with c:
                    if st.button("🗑️ حذف الحساب", key=f"delete_member_{user['id']}", use_container_width=True):
                        st.session_state[f"confirm_member_delete_{user['id']}"] = True
                        st.rerun()

            if st.session_state.get(f"confirm_member_delete_{user['id']}"):
                st.warning("تأكيد حذف حساب الدخول فقط — القضايا والمستندات ستظل محفوظة.")
                y,n = st.columns(2)
                with y:
                    if st.button("نعم، احذف الحساب", key=f"yes_delete_member_{user['id']}", type="primary", use_container_width=True):
                        users = [u for u in users if u.get("id") != user.get("id")]
                        save_users(users)
                        st.session_state.pop(f"confirm_member_delete_{user['id']}", None)
                        st.rerun()
                with n:
                    if st.button("إلغاء", key=f"no_delete_member_{user['id']}", use_container_width=True):
                        st.session_state.pop(f"confirm_member_delete_{user['id']}", None)
                        st.rerun()

# ===== الرسائل والبيانات الداخلية =====
MESSAGES_FILE = os.path.join(LOCAL_DATA_DIR, "messages_v1.json")

def load_messages():
    data = load_json_persistent("messages", MESSAGES_FILE, [])
    return data if isinstance(data, list) else []

def save_messages(items):
    return save_json_persistent("messages", MESSAGES_FILE, items)

def role_label(role):
    return {
        "admin":"الادمن", "member":"عضو قانونى", "manager":"مدير إدارة",
        "general_manager":"مدير عام",
        "leader1":"قيادي مركز رئيسى 1", "leader2":"قيادي مركز رئيسى 2",
        "leader3":"قيادي مركز رئيسى 3"
    }.get(role, role or "")

def can_message(sender_role, target_role):
    if target_role == "admin" and sender_role != "admin":
        return True
    if sender_role == "admin":
        return target_role in ("admin", "member", "manager", "general_manager", "leader1", "leader2", "leader3")
    if sender_role == "member":
        return target_role in ("admin", "member", "manager")
    if sender_role == "manager":
        return target_role in ("admin", "member", "manager", "general_manager")
    if sender_role == "leader1":
        return target_role in ("admin", "member", "manager", "general_manager", "leader1", "leader2", "leader3")
    if sender_role == "leader2":
        return target_role in ("admin", "member", "manager", "general_manager", "leader1", "leader2", "leader3")
    if sender_role == "leader3":
        return target_role in ("admin", "member", "manager", "general_manager", "leader2", "leader3")
    if sender_role == "general_manager":
        return target_role in ("admin", "member", "manager", "general_manager", "leader2", "leader3")
    return False

def person_display(u, include_role=True):
    name = u.get("display_name") or u.get("username") or "بدون اسم"
    role = role_label(u.get("role"))
    if u.get("role") == "admin":
        role = "الادمن"
    area = u.get("insurance_area") or "غير محددة"
    if include_role:
        return f"{name} — {role} — المنطقة التأمينية: {area}"
    return name

def _recipient_groups(candidates, key_prefix, preselected=None):
    preselected = set(preselected or [])
    selected_ids = []
    grouped = {}
    for u in candidates:
        area = (u.get("insurance_area") or "غير محددة").strip() or "غير محددة"
        grouped.setdefault(area, []).append(u)
    for area in sorted(grouped, key=lambda x: x.casefold()):
        st.markdown(f"<div style='color:#D4AF37;font-size:17px;font-weight:900;margin:10px 0 6px;'>المنطقة: {html.escape(area)}</div>", unsafe_allow_html=True)
        cols = st.columns(3)
        for i, u in enumerate(sorted(grouped[area], key=lambda x: (role_order(x.get("role")), str(x.get("display_name") or x.get("username") or "").casefold()))):
            uid = u.get("id")
            label = person_display(u)
            with cols[i % 3]:
                checked = st.checkbox(label, value=(uid in preselected), key=f"{key_prefix}_{uid}")
                if checked:
                    selected_ids.append(uid)
    return selected_ids

def messaging_page():
    user = st.session_state.get("user") or {}
    sr = user.get("role")
    if sr not in ("admin", "member", "manager", "general_manager", "leader1", "leader2", "leader3"):
        return

    users = _cached_users_snapshot()
    messages = load_messages()
    st.markdown("<h2 style='color:#D4AF37;text-align:center;'>📢 الرسائل والبيانات</h2>", unsafe_allow_html=True)

    target_options = [
        ("الادمن", "admin"), ("عضو قانونى", "member"), ("مدير إدارة", "manager"),
        ("مدير عام", "general_manager"), ("قيادي مركز رئيسى 1", "leader1"),
        ("قيادي مركز رئيسى 2", "leader2"), ("قيادي مركز رئيسى 3", "leader3")
    ]
    allowed_roles = [r for _, r in target_options if can_message(sr, r)]
    preselected = st.session_state.pop("message_target_ids", [])
    candidates = [u for u in users if u.get("status") == "active" and u.get("id") != user.get("id") and can_message(sr, u.get("role"))]

    if not candidates:
        st.info("لا يوجد مستلمون متاحون وفق صلاحيات الحساب")
        return

    target_kind = st.selectbox(
        "نوع المستلمين",
        ["الجميع"] + [label for label, role in target_options if role in allowed_roles] + ["محددون"],
        key=f"msg_target_kind_{sr}"
    )
    text = st.text_area("نص الرسالة أو البيان", height=130, key=f"msg_text_{sr}")

    attach_count_key = f"msg_attach_count_{sr}"
    if attach_count_key not in st.session_state:
        st.session_state[attach_count_key] = 1
    attachment_files = []
    for ai in range(st.session_state[attach_count_key]):
        af = st.file_uploader(
            "📎 المرفق" if ai == 0 else f"📎 مرفق إضافي {ai + 1}",
            type=None,
            key=f"msg_attachment_{sr}_{ai}",
            help="يمكن رفع PDF أو Word أو Excel أو صورة أو أي ملف آخر"
        )
        attachment_files.append(af)
        if af is not None and st.button("➕ إضافة مرفق آخر", key=f"msg_add_attachment_{sr}_{ai}", use_container_width=True):
            st.session_state[attach_count_key] += 1
            st.rerun()

    image_file = st.file_uploader("🖼️ الصورة", type=["jpg","jpeg","png"], key=f"msg_image_{sr}")

    if target_kind == "الجميع":
        selected_ids = [u.get("id") for u in candidates]
    elif target_kind == "محددون":
        selected_ids = _recipient_groups(candidates, f"msg_chk_{sr}", preselected)
    else:
        rr = dict(target_options).get(target_kind)
        role_candidates = [u for u in candidates if u.get("role") == rr]
        selected_ids = _recipient_groups(role_candidates, f"msg_chk_{sr}_{rr}", preselected)

    if st.button("📤 إرسال الرسالة / البيان", type="primary", use_container_width=True, key=f"send_message_{sr}"):
        if not text.strip():
            st.error("اكتب نص الرسالة أو البيان")
            return
        if not selected_ids:
            st.error("حدد مستلمًا واحدًا على الأقل")
            return
        messages.append({
            "id": secrets.token_hex(10),
            "sender_id": user.get("id"),
            "sender_name": user.get("display_name") or user.get("username"),
            "sender_role": sr,
            "sender_job_title": user.get("job_title") or role_label(sr),
            "sender_insurance_area": user.get("insurance_area") or "",
            "sender_photo": user.get("profile_photo") or "",
            "sender_photo_type": user.get("profile_photo_type") or "",
            "text": text.strip(),
            "recipient_ids": selected_ids,
            "created_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "image_data": base64.b64encode(image_file.getvalue()).decode("utf-8") if image_file else "",
            "image_type": image_file.type if image_file else "",
            "attachments": [{"data": base64.b64encode(af.getvalue()).decode("utf-8"), "name": af.name, "type": af.type or "application/octet-stream"} for af in attachment_files if af is not None],
            "attachment_data": base64.b64encode(attachment_files[0].getvalue()).decode("utf-8") if attachment_files and attachment_files[0] else "",
            "attachment_name": attachment_files[0].name if attachment_files and attachment_files[0] else "",
            "attachment_type": attachment_files[0].type if attachment_files and attachment_files[0] else "application/octet-stream",
            "in_reply_to": ""
        })
        save_messages(messages)
        st.session_state.pop("message_target_ids", None)
        st.success("تم إرسال الرسالة وحفظها")
        st.rerun()

def _photo_html(photo_data, photo_type="image/png", size=72):
    if not photo_data:
        return ""
    return f"<img src='data:{html.escape(photo_type)};base64,{photo_data}' style='width:{size}px;height:{size}px;border-radius:50%;object-fit:cover;border:2px solid #D4AF37;margin:0 auto 8px;display:block;'>"

def _visible_banners_for_user(user):
    username = user.get("username", "")
    now = datetime.now()
    out = []
    for b in load_banners():
        try:
            expire = datetime.fromisoformat(str(b.get("expire_at") or b.get("expire")))
        except Exception:
            continue
        if expire <= now:
            continue
        audience = b.get("audience", "الكل")
        if audience not in ("الكل", "كل الأعضاء") and username not in b.get("visible_to", []):
            continue
        out.append(b)
    return out

def show_top_notices():
    user = st.session_state.get("user") or {}
    uid = user.get("id")
    if not uid:
        return
    users = _cached_users_snapshot()
    by_id = {u.get("id"): u for u in users}
    notices = []
    for m in load_messages():
        if uid in m.get("recipient_ids", []):
            sender = by_id.get(m.get("sender_id"), {})
            notices.append({"kind":"message", "created_at":m.get("created_at", ""), "item":m, "sender":sender})
    for b in _visible_banners_for_user(user):
        sender = by_id.get(b.get("sender_id"), {})
        notices.append({"kind":"banner", "created_at":b.get("created_at", ""), "item":b, "sender":sender})
    if not notices:
        return
    notices.sort(key=lambda x: str(x.get("created_at", "")), reverse=True)
    total = len(notices)
    idx = int(st.session_state.get("top_notice_index", 0))
    idx = max(0, min(idx, total - 1))
    st.session_state.top_notice_index = idx
    n = notices[idx]
    item = n["item"]
    sender = n["sender"]
    sender_name = item.get("sender_name") or sender.get("display_name") or sender.get("username") or "الادمن"
    sender_role = item.get("sender_role") or sender.get("role") or "admin"
    sender_job = item.get("sender_job_title") or sender.get("job_title") or role_label(sender_role)
    sender_area = item.get("sender_insurance_area") or sender.get("insurance_area") or "غير محددة"
    photo = item.get("sender_photo") or sender.get("profile_photo") or ""
    photo_type = item.get("sender_photo_type") or sender.get("profile_photo_type") or "image/png"
    text = item.get("text", "")
    image_data = item.get("image_data", "")
    image_type = item.get("image_type", "image/png")
    title = "رسالة واردة" if n["kind"] == "message" else "بيان وارد"
    created = html.escape(fmt_date_ar(item.get("created_at", ""), str(item.get("created_at", ""))))
    sender_color = "#00E676" if sender_role == "admin" else "#D4AF37"
    card = f"<div style='background:linear-gradient(135deg,#101b31,#203b63);border:2px solid {sender_color};border-radius:18px;padding:16px;margin:0 0 10px;box-shadow:0 10px 28px rgba(0,0,0,.3);text-align:center;'>"
    card += _photo_html(photo, photo_type, 70)
    card += f"<div style='color:#D4AF37;font-size:22px;font-weight:900;'>📢 {title}</div>"
    admin_prefix = "<span style=\"color:#ff3b30;font-weight:1000;margin-left:8px;\">الادمن</span>" if sender_role == "admin" else ""
    card += f"<div style='color:{sender_color};font-size:20px;font-weight:900;margin-top:5px;'>المرسل: {admin_prefix} {html.escape(str(sender_name))}</div>"
    card += f"<div style='color:#fff;font-size:15px;font-weight:800;margin-top:4px;'>الوظيفة: {html.escape(str(sender_job))} — المنطقة التأمينية: {html.escape(str(sender_area))}</div>"
    card += f"<div style='color:#fff;font-size:18px;font-weight:700;margin-top:12px;white-space:pre-wrap;overflow-wrap:break-word;text-align:right;direction:rtl;'>{html.escape(str(text))}</div>"
    if image_data:
        card += f"<img src='data:{html.escape(image_type)};base64,{image_data}' style='max-width:100%;max-height:360px;border-radius:14px;margin-top:12px;border:1px solid #D4AF37;'>"
    attachments = item.get("attachments") or []
    if not attachments and item.get("attachment_data"):
        attachments = [{"data": item.get("attachment_data"), "name": item.get("attachment_name") or "مرفق", "type": item.get("attachment_type") or "application/octet-stream"}]
    for att in attachments:
        card += f"<div style='margin-top:10px;padding:9px;border:1px solid #D4AF37;border-radius:12px;color:#fff;font-weight:800;'>📎 {html.escape(str(att.get('name') or 'مرفق'))}</div>"
    card += f"<div style='color:#cbd5e1;font-size:12px;margin-top:10px;'>{created}</div></div>"
    st.markdown(card, unsafe_allow_html=True)
    attachments = item.get("attachments") or []
    if not attachments and item.get("attachment_data"):
        attachments = [{"data": item.get("attachment_data"), "name": item.get("attachment_name") or "مرفق", "type": item.get("attachment_type") or "application/octet-stream"}]
    for ai, att in enumerate(attachments):
        try:
            attachment_bytes = base64.b64decode(att.get("data") or "")
            st.download_button(
                f"📎 فتح / تنزيل: {att.get('name') or 'مرفق'}",
                data=attachment_bytes, file_name=att.get("name") or "مرفق",
                mime=att.get("type") or "application/octet-stream", use_container_width=True,
                key=f"download_notice_attachment_{item.get('id','x')}_{ai}"
            )
        except Exception:
            pass
    c1,c2,c3 = st.columns([1,1,1])
    with c1:
        if st.button("⬅️ السابقة", disabled=(idx >= total-1), use_container_width=True, key="top_notice_prev_v4"):
            st.session_state.top_notice_index = min(total-1, idx+1)
            st.rerun()
    with c2:
        st.markdown(f"<div style='text-align:center;color:#fff;font-weight:900;padding-top:8px;'>الرسالة {idx+1} من {total}</div>", unsafe_allow_html=True)
    with c3:
        if st.button("التالية ➡️", disabled=(idx <= 0), use_container_width=True, key="top_notice_next_v4"):
            st.session_state.top_notice_index = max(0, idx-1)
            st.rerun()

    if n["kind"] == "message":
        st.markdown("<div style='color:#D4AF37;font-size:18px;font-weight:900;text-align:center;margin-top:10px;'>الرد على الرسالة</div>", unsafe_allow_html=True)
        reply_text = st.text_area("نص الرد", height=110, key=f"top_reply_text_{item.get('id','x')}")
        reply_count_key = f"reply_attach_count_{item.get('id','x')}"
        if reply_count_key not in st.session_state:
            st.session_state[reply_count_key] = 1
        reply_attachments = []
        for ri in range(st.session_state[reply_count_key]):
            rf = st.file_uploader("📎 مرفق الرد" if ri == 0 else f"📎 مرفق رد إضافي {ri + 1}", type=None, key=f"top_reply_attachment_{item.get('id','x')}_{ri}", help="PDF أو Word أو Excel أو صورة أو أي ملف")
            reply_attachments.append(rf)
            if rf is not None and st.button("➕ إضافة مرفق آخر للرد", key=f"reply_add_attachment_{item.get('id','x')}_{ri}", use_container_width=True):
                st.session_state[reply_count_key] += 1
                st.rerun()
        if st.button("↩️ إرسال الرد", type="primary", use_container_width=True, key=f"top_reply_send_{item.get('id','x')}"):
            if not reply_text.strip() and not any(reply_attachments):
                st.error("اكتب الرد أو أرفق ملفًا")
            else:
                target_id = item.get("sender_id")
                if target_id is not None:
                    msgs = load_messages()
                    msgs.append({
                        "id": secrets.token_hex(10),
                        "sender_id": user.get("id"),
                        "sender_name": user.get("display_name") or user.get("username"),
                        "sender_role": user.get("role"),
                        "sender_job_title": user.get("job_title") or role_label(user.get("role")),
                        "sender_insurance_area": user.get("insurance_area") or "",
                        "sender_photo": user.get("profile_photo") or "",
                        "sender_photo_type": user.get("profile_photo_type") or "",
                        "text": reply_text.strip(),
                        "recipient_ids": [target_id],
                        "created_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
                        "image_data": "",
                        "image_type": "",
                        "attachments": [{"data": base64.b64encode(rf.getvalue()).decode("utf-8"), "name": rf.name, "type": rf.type or "application/octet-stream"} for rf in reply_attachments if rf is not None],
                        "attachment_data": base64.b64encode(reply_attachments[0].getvalue()).decode("utf-8") if reply_attachments and reply_attachments[0] else "",
                        "attachment_name": reply_attachments[0].name if reply_attachments and reply_attachments[0] else "",
                        "attachment_type": reply_attachments[0].type if reply_attachments and reply_attachments[0] else "application/octet-stream",
                        "in_reply_to": item.get("id", "")
                    })
                    save_messages(msgs)
                    st.success("تم إرسال الرد")
                    st.rerun()

def role_order(role):
    return {"leader1":0,"leader2":1,"leader3":2,"general_manager":3,"manager":4,"member":5,"admin":-1}.get(role,99)

def show_member_directory():
    users = [u for u in load_users() if u.get("status") != "deleted"]
    users.sort(key=lambda u: (role_order(u.get("role")), str(u.get("insurance_area") or "").casefold(), str(u.get("display_name") or u.get("username") or "").casefold()))
    current = st.session_state.get("user") or {}
    st.markdown("<div style='color:#D4AF37;font-size:21px;font-weight:900;text-align:center;margin:14px 0 8px;'>العضويات</div>", unsafe_allow_html=True)
    for area in sorted(set((u.get("insurance_area") or "غير محددة").strip() or "غير محددة" for u in users), key=lambda x:x.casefold()):
        st.markdown(f"<div style='color:#D4AF37;font-size:14px;font-weight:900;margin:7px 0 5px;'>المنطقة: {html.escape(area)}</div>", unsafe_allow_html=True)
        area_users = [u for u in users if ((u.get("insurance_area") or "غير محددة").strip() or "غير محددة") == area]
        cols = st.columns(3)
        for i,u in enumerate(area_users):
            with cols[i % 3]:
                name = html.escape(str(u.get("display_name") or u.get("username") or ""))
                job_value = membership_job_title(u) or role_label(u.get("role"))
                job = html.escape(job_value)
                photo = _photo_html(u.get("profile_photo"), u.get("profile_photo_type") or "image/png", 42)
                is_admin = u.get("role") == "admin"
                border = "#00E676" if is_admin else "#D4AF37"
                name_color = "#00E676" if is_admin else "#fff"
                admin_badge = "<div style='color:#00E676;font-size:10px;font-weight:900;'>ADMIN</div>" if is_admin else ""
                card = f"<div style='background:#142038;border:2px solid {border};border-radius:12px;padding:8px;text-align:center;min-height:116px;margin-bottom:5px;width:100%;'>{photo}<div style='color:{name_color};font-size:12px;font-weight:900;line-height:1.25;'>{name}</div>{admin_badge}<div style='color:#D4AF37;font-size:10px;font-weight:800;line-height:1.25;'>{job}</div><div style='color:#cbd5e1;font-size:9px;line-height:1.25;'>{html.escape(area)}</div></div>"
                st.markdown(card, unsafe_allow_html=True)
                if u.get("id") != current.get("id") and can_message(current.get("role"), u.get("role")):
                    if st.button("✉️ إرسال", key=f"dir_msg_v5_{current.get('id')}_{u.get('id')}", use_container_width=True):
                        st.session_state.message_target_ids = [u.get("id")]
                        st.session_state.page = "الرسائل"
                        st.rerun()

def global_people_search():
    users=_cached_users_snapshot()
    q=st.text_input("🔎 بحث سريع بالاسم", placeholder="اكتب اسم العضو أو مدير الإدارة أو المدير العام", key="global_people_search")
    if q.strip():
        matches=[u for u in users if q.strip().lower() in str(u.get("display_name") or u.get("username","")).lower()]
        if matches:
            for u in matches:
                st.markdown(f"**👤 {u.get('display_name') or u.get('username')}** — {role_label(u.get('role'))} — المنطقة التأمينية: {u.get('insurance_area') or 'لا توجد منطقة'} — الحالة: {u.get('status','')}")
        else:
            st.info("لا توجد نتيجة بهذا الاسم")

def admin_control_page():
    if not st.session_state.user or st.session_state.user.get("role") != "admin":
        st.error("هذه الصفحة متاحة للادمن فقط")
        return

    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>⚖️ لوحة التحكم</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:white;text-align:center;'>جميع أدوات إدارة الموقع في مكان واحد</p>", unsafe_allow_html=True)

    tab_users, tab_messages, tab_search, tab_account = st.tabs(["👥 الأعضاء", "📢 الرسائل والبيانات", "🔎 البحث والملف الشخصي", "🔐 حساب الادمن والإيميل"])

    with tab_users:
        manage_users_page()
        st.markdown("---")
        if st.button("➕ استخراج عضوية جديدة", use_container_width=True, key="admin_extract_from_panel"):
            st.session_state.page = "extract_member"
            st.rerun()

    with tab_messages:
        messaging_page()
        st.markdown("---")
        st.markdown("<h3 style='color:#D4AF37;'>📜 سجل البيانات والرسائل</h3>", unsafe_allow_html=True)
        users = _cached_users_snapshot()
        banners = _cached_banners_snapshot()
        now = datetime.now()

        with st.form("admin_banner_add_v2"):
            text = st.text_area("نص الرسالة أو البيان", height=90)
            color = st.color_picker("لون الرسالة أو البيان", "#D4AF37")
            banner_files = st.file_uploader("مرفقات البيان", type=None, accept_multiple_files=True, key="banner_files_add_v4")
            col1,col2 = st.columns(2)
            with col1:
                duration_value = st.number_input("مدة الإبقاء", min_value=1, max_value=3650, value=7, step=1)
            with col2:
                duration_unit = st.selectbox("وحدة المدة", ["أيام","ساعات","دقائق"])
            audience = st.radio("يظهر لـ", ["كل الأعضاء","أعضاء محددين"], horizontal=True, key="banner_audience_v2")
            visible_to = []
            if audience == "أعضاء محددين":
                banner_candidates = [u for u in users if u.get("status") == "active"]
                selected_banner_ids = _recipient_groups(banner_candidates, "banner_recipient_chk_v3")
                visible_to = [u.get("username") for u in banner_candidates if u.get("id") in selected_banner_ids]
            if st.form_submit_button("🚀 نشر الرسالة / البيان", use_container_width=True, type="primary"):
                if not text.strip():
                    st.error("اكتب نص الرسالة أو البيان")
                elif audience == "أعضاء محددين" and not visible_to:
                    st.error("حدد عضواً واحداً على الأقل")
                else:
                    minutes = duration_value * (1440 if duration_unit == "أيام" else 60 if duration_unit == "ساعات" else 1)
                    created = datetime.now()
                    banner = {
                        "id": secrets.token_hex(8),
                        "text": text.strip(),
                        "color": color,
                        "created_at": created.strftime("%d/%m/%Y %H:%M"),
                        "expire_at": (created + timedelta(minutes=minutes)).isoformat(),
                        "audience": "الكل" if audience == "كل الأعضاء" else "اعضاء محددين",
                        "visible_to": visible_to,
                        "sender_id": st.session_state.user.get("id"),
                        "sender_name": st.session_state.user.get("display_name") or st.session_state.user.get("username"),
                        "sender_role": st.session_state.user.get("role"),
                        "sender_job_title": st.session_state.user.get("job_title") or role_label(st.session_state.user.get("role")),
                        "sender_insurance_area": st.session_state.user.get("insurance_area") or "",
                        "sender_photo": st.session_state.user.get("profile_photo") or "",
                        "sender_photo_type": st.session_state.user.get("profile_photo_type") or "",
                        "image_data": "",
                        "image_type": "",
                        "attachments": [{"data": base64.b64encode(f.getvalue()).decode("utf-8"), "name": f.name, "type": f.type or "application/octet-stream"} for f in (banner_files or [])]
                    }
                    banners.append(banner)
                    save_banners(banners)
                    st.success("تم نشر الرسالة / البيان")
                    st.rerun()

        st.markdown("---")
        st.markdown(f"<h3 style='color:#D4AF37;'>الرسائل والبيانات الظاهرة أعلى الصفحة: {len(banners)}</h3>", unsafe_allow_html=True)

        for i,b in enumerate(banners):
            expire_raw = b.get("expire_at") or b.get("expire")
            expire_dt = parse_date_safe(expire_raw)
            try:
                expire_full_dt = datetime.fromisoformat(str(expire_raw))
            except Exception:
                expire_full_dt = None
            expired = bool(expire_full_dt and expire_full_dt <= datetime.now())
            audience_text = "الكل" if b.get("audience") in ("الكل","كل الأعضاء") else "، ".join(b.get("visible_to", []))
            expire_display = fmt_date_ar(expire_raw, "غير محدد")
            with st.container(border=True):
                st.markdown(
                    f"<div class='banner-admin-card'><div style='color:#D4AF37;font-size:20px;font-weight:900;'>📢 {b.get('text','')}</div>"
                    f"<div style='color:white;margin-top:8px;'>الجمهور: {audience_text}</div>"
                    f"<div style='color:#ddd;'>تنتهي: {expire_display} | الحالة: {'منتهية' if expired else 'فعالة'}</div>"
                    + "".join(f"<div style='color:#fff;margin-top:5px;'>📎 {html.escape(str(a.get('name') or 'مرفق'))}</div>" for a in (b.get('attachments') or []))
                    + "</div>",
                    unsafe_allow_html=True
                )
                e1,e2 = st.columns(2)
                with e1:
                    if st.button("✏️ تعديل", key=f"edit_banner_{b.get('id',i)}", use_container_width=True):
                        st.session_state["edit_banner_id"] = b.get("id")
                        st.rerun()
                with e2:
                    if st.button("🗑️ حذف", key=f"delete_banner_{b.get('id',i)}", use_container_width=True):
                        banners = [x for x in banners if x.get("id") != b.get("id")]
                        save_banners(banners)
                        st.rerun()

                if st.session_state.get("edit_banner_id") == b.get("id"):
                    with st.form(f"edit_banner_form_{b.get('id',i)}"):
                        nt = st.text_area("نص الرسالة أو البيان", value=b.get("text",""), key=f"et_{b.get('id',i)}")
                        nc = st.color_picker("اللون", value=b.get("color","#D4AF37"), key=f"ec_{b.get('id',i)}")
                        nfiles = st.file_uploader("مرفقات البيان", type=None, accept_multiple_files=True, key=f"eimg_{b.get('id',i)}")
                        na = st.radio("المستلمون", ["كل الأعضاء","أعضاء محددين"], index=0 if b.get("audience") in ("الكل","كل الأعضاء") else 1, horizontal=True, key=f"ea_{b.get('id',i)}")
                        nvis = []
                        if na == "أعضاء محددين":
                            edit_candidates = [u for u in users if u.get("status") == "active"]
                            default_ids = [u.get("id") for u in edit_candidates if u.get("username") in b.get("visible_to", [])]
                            selected_edit_ids = _recipient_groups(edit_candidates, f"banner_edit_recipient_{b.get('id',i)}", default_ids)
                            nvis = [u.get("username") for u in edit_candidates if u.get("id") in selected_edit_ids]
                        new_expire = date_input_dmy("تاريخ الانتهاء", value=expire_dt or datetime.now().date(), max_value=date(2100, 12, 31), key=f"ed_{b.get('id',i)}")
                        ec1,ec2 = st.columns(2)
                        with ec1:
                            if st.form_submit_button("💾 حفظ التعديل", use_container_width=True):
                                b["text"] = nt.strip()
                                b["color"] = nc
                                b["audience"] = "الكل" if na == "كل الأعضاء" else "اعضاء محددين"
                                b["visible_to"] = nvis
                                if nfiles:
                                    b["attachments"] = b.get("attachments") or []
                                    b["attachments"].extend([{"data": base64.b64encode(f.getvalue()).decode("utf-8"), "name": f.name, "type": f.type or "application/octet-stream"} for f in nfiles])
                                old_dt = expire_full_dt or datetime.now()
                                parsed_exp = parse_date_safe(new_expire) or datetime.now().date()
                                b["expire_at"] = datetime.combine(parsed_exp, old_dt.time()).isoformat()
                                save_banners(banners)
                                st.session_state.pop("edit_banner_id", None)
                                st.rerun()
                        with ec2:
                            if st.form_submit_button("إلغاء", use_container_width=True):
                                st.session_state.pop("edit_banner_id", None)
                                st.rerun()

    with tab_search:
        st.markdown("<h3 style='color:#D4AF37;'>🔎 البحث عن عضو أو مدير إدارة أو مدير عامأو قيادي</h3>", unsafe_allow_html=True)
        global_people_search()
        st.markdown("---")
        st.markdown("<h3 style='color:#D4AF37;'>👤 الملف الشخصي</h3>", unsafe_allow_html=True)
        admin_user = next((u for u in load_users() if u.get("role") == "admin"), st.session_state.user)
        st.write(f"**الاسم:** {admin_user.get('username','')}")
        st.write(f"**الوظيفة:** {role_label(admin_user.get('role'))}")
        st.write(f"**البريد الإلكتروني:** {admin_user.get('email') or 'غير مسجل'}")
        st.write(f"**الحالة:** {admin_user.get('status','')}")

    with tab_account:
        st.markdown("<h3 style='color:#D4AF37;'>👤 بيانات حساب الادمن</h3>", unsafe_allow_html=True)
        users = _cached_users_snapshot()
        admin_user = next((u for u in users if u.get("role") == "admin"), None)
        if admin_user:
            with st.form("admin_profile_edit_form_v4"):
                display_name = st.text_input("الاسم", value=admin_user.get("display_name") or admin_user.get("username") or "admin")
                job_title = st.text_input("الوظيفة", value=admin_user.get("job_title") or "")
                insurance_area = st.text_input("المنطقة التأمينية", value=admin_user.get("insurance_area") or "")
                login_name = st.text_input("اسم الدخول", value=admin_user.get("username") or "admin")
                admin_photo = st.file_uploader("الصورة الشخصية", type=["jpg","jpeg","png"], key="admin_profile_photo_upload")
                if st.form_submit_button("💾 حفظ بيانات الادمن", use_container_width=True, type="primary"):
                    login_name = login_name.strip() or "admin"
                    if login_name != admin_user.get("username") and any(u.get("username") == login_name and u.get("id") != admin_user.get("id") for u in users):
                        st.error("اسم الدخول مستخدم بالفعل")
                    else:
                        admin_user["display_name"] = display_name.strip() or login_name
                        admin_user["username"] = login_name
                        admin_user["job_title"] = job_title.strip()
                        admin_user["insurance_area"] = insurance_area.strip()
                        if admin_photo:
                            admin_user["profile_photo"] = base64.b64encode(admin_photo.getvalue()).decode("utf-8")
                            admin_user["profile_photo_type"] = admin_photo.type
                        save_users(users)
                        st.session_state.user = admin_user
                        st.success("تم حفظ اسم الادمن والوظيفة والمنطقة التأمينية وبيانات الدخول")
                        st.rerun()
        st.info("تظل صفة الحساب ظاهرة دائمًا باسم: الادمن. ويمكن لمن لديه صلاحية مراسلة الادمن رؤيته بالاسم والوظيفة والمنطقة التأمينية.")
        recovery_settings_page()
        st.markdown("---")
        change_password_page()

    st.markdown("---")
    if st.button("🚪 تسجيل الخروج", use_container_width=True, key="admin_control_logout_v3"):
        st.session_state.user = None
        st.session_state.role = None
        st.session_state.page = "login"
        st.rerun()

    if st.button("⬅️ العودة للرئيسية", use_container_width=True, key="back_admin_panel"):
        st.session_state.page = "الرئيسية"
        st.rerun()

def recovery_settings_page():
    st.markdown("<h2 style='text-align:center; color:#C9A961'>تأكيد البريد الالكتروني</h2>", unsafe_allow_html=True)
    users = _cached_users_snapshot()
    user = next((u for u in users if u["id"] == st.session_state.user["id"]), None)
    st.info("هذا البريد سيستخدم لاسترجاع حسابك في حالة الفقد")
    email = st.text_input("البريد الالكتروني", value=user.get("email",""))
    if user["role"] == "admin":
        recovery_email = st.text_input("ايميل استرجاع اضافي للادمن", value=user.get("recovery_email",""))
    else:
        recovery_email = user.get("recovery_email","")
    if st.button("حفظ البريد", use_container_width=True):
        user["email"] = email
        if user["role"] == "admin":
            user["recovery_email"] = recovery_email
        save_users(users); st.session_state.user = user; st.success("تم حفظ البريد بنجاح")

def change_password_page():
    st.markdown("<h1 style='text-align:center; color:#C9A961'>تغيير كلمة السر</h1>", unsafe_allow_html=True)
    old_pass = st.text_input("كلمة السر القديمة", type="password")
    new_pass = st.text_input("كلمة السر الجديدة", type="password")
    if st.button("تغيير", use_container_width=True):
        if bcrypt.checkpw(old_pass.encode(), st.session_state.user["password"].encode()):
            users = _cached_users_snapshot()
            for user in users:
                if user["id"] == st.session_state.user["id"]:
                    user["password"] = bcrypt.hashpw(new_pass.encode(), bcrypt.gensalt()).decode()
                    save_users(users); st.session_state.user = user; st.success("تم تغيير الباسورد"); st.rerun()
        else:
            st.error("كلمة السر القديمة غلط")

def set_password_page():
    st.markdown("<div class='login-shell'>", unsafe_allow_html=True)
    st.markdown("<h1 style='text-align:center;color:#D4AF37!important;'>🔐 تفعيل العضوية</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center;color:#fff!important;font-weight:900;'>أدخل كلمة السر التي تريد استخدامها لحسابك</p>", unsafe_allow_html=True)
    new_pass = st.text_input("كلمة السر الجديدة", type="password", key="first_new_pass")
    confirm_pass = st.text_input("تأكيد كلمة السر", type="password", key="first_confirm_pass")
    profile_photo = st.file_uploader("الصورة الشخصية", type=["jpg","jpeg","png"], key="first_profile_photo")
    if st.button("حفظ كلمة السر والدخول", type="primary", use_container_width=True, key="save_first_password"):
        if not new_pass:
            st.error("اكتب كلمة السر أولاً")
            return
        if new_pass != confirm_pass:
            st.error("كلمتا السر غير متطابقتين")
            return
        users = _cached_users_snapshot()
        target_id = st.session_state.get("temp_user_id")
        target_name = str(st.session_state.get("temp_user", "")).strip().casefold()
        target = None
        for user in users:
            if target_id is not None and user.get("id") == target_id:
                target = user
                break
            if normalize_person_name(user.get("username")) == normalize_person_name(target_name) or normalize_person_name(user.get("display_name")) == normalize_person_name(target_name):
                target = user
                break
        if not target:
            st.error("تعذر العثور على العضوية. ارجع لتسجيل الدخول وأعد إدخال اسم العضو كما هو.")
            return
        target["password"] = bcrypt.hashpw(new_pass.encode(), bcrypt.gensalt()).decode()
        target["password_set"] = True
        if profile_photo:
            target["profile_photo"] = base64.b64encode(profile_photo.getvalue()).decode("utf-8")
            target["profile_photo_type"] = profile_photo.type
        if target.get("status") in (None, "suspended_temp"):
            target["status"] = "active"
        save_users(users)
        st.session_state.user = target
        st.session_state.role = target.get("role")
        st.session_state.page = "الرئيسية"
        st.session_state.pop("temp_user", None)
        st.session_state.pop("temp_user_id", None)
        st.success("تم تفعيل العضوية وتسجيل الدخول بنجاح")
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

def personal_control_page():
    user = st.session_state.get("user") or {}
    role = user.get("role")
    if not user or role == "admin":
        return

    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>⚙️ لوحة تحكمى</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:white;text-align:center;font-weight:800;'>ملفك الشخصي وصلاحياتك وحسابك ورسائلك</p>", unsafe_allow_html=True)

    tab_profile, tab_permissions, tab_messages, tab_account = st.tabs(
        ["👤 الملف الشخصي", "🛡️ صلاحياتى", "📢 الرسائل والبيانات", "🔐 الحساب"]
    )

    with tab_profile:
        st.markdown("<h3 style='color:#D4AF37;'>بيانات العضوية</h3>", unsafe_allow_html=True)
        st.write(f"**الاسم:** {user.get('username','')}")
        st.write(f"**الوظيفة:** {user.get('job_title') or role_label(role)}")
        st.write(f"**تاريخ الميلاد:** {fmt_date_ar(user.get('birth_date'), 'غير مسجل')}")
        st.write(f"**المنطقة التأمينية:** {user.get('insurance_area') or 'غير مسجلة'}")
        st.write(f"**البريد الإلكتروني:** {user.get('email') or 'غير مسجل'}")
        st.write(f"**الحالة:** {user.get('status','')}")
        st.info("إيقاف حساب الدخول لا يحذف القضايا أو المستندات أو الرسائل المحفوظة.")

    with tab_permissions:
        st.markdown("<h3 style='color:#D4AF37;'>الصلاحيات المقررة لحسابك</h3>", unsafe_allow_html=True)
        targets = [
            ("الادمن","admin"), ("عضو قانونى","member"), ("مدير إدارة","manager"),
            ("مدير عام","general_manager"), ("قيادي مركز رئيسى 1","leader1"),
            ("قيادي مركز رئيسى 2","leader2"), ("قيادي مركز رئيسى 3","leader3")
        ]
        allowed = [label for label, target_role in targets if can_message(role, target_role)]
        if allowed:
            st.success("يمكنك إرسال رسالة أو بيان إلى: " + "، ".join(allowed))
        else:
            st.info("لا توجد صلاحية إرسال رسائل لهذا الحساب.")

    with tab_messages:
        messaging_page()

    with tab_account:
        st.markdown("<h3 style='color:#D4AF37;'>👤 بيانات حساب الادمن</h3>", unsafe_allow_html=True)
        users = _cached_users_snapshot()
        admin_user = next((u for u in users if u.get("role") == "admin"), None)
        if admin_user:
            with st.form("admin_profile_edit_form_v4"):
                display_name = st.text_input("الاسم", value=admin_user.get("display_name") or admin_user.get("username") or "admin")
                job_title = st.text_input("الوظيفة", value=admin_user.get("job_title") or "")
                insurance_area = st.text_input("المنطقة التأمينية", value=admin_user.get("insurance_area") or "")
                login_name = st.text_input("اسم الدخول", value=admin_user.get("username") or "admin")
                if st.form_submit_button("💾 حفظ بيانات الادمن", use_container_width=True, type="primary"):
                    login_name = login_name.strip() or "admin"
                    if login_name != admin_user.get("username") and any(u.get("username") == login_name and u.get("id") != admin_user.get("id") for u in users):
                        st.error("اسم الدخول مستخدم بالفعل")
                    else:
                        admin_user["display_name"] = display_name.strip() or login_name
                        admin_user["username"] = login_name
                        admin_user["job_title"] = job_title.strip()
                        admin_user["insurance_area"] = insurance_area.strip()
                        save_users(users)
                        st.session_state.user = admin_user
                        st.success("تم حفظ اسم الادمن والوظيفة والمنطقة التأمينية وبيانات الدخول")
                        st.rerun()
        st.info("تظل صفة الحساب ظاهرة دائمًا باسم: الادمن. ويمكن لمن لديه صلاحية مراسلة الادمن رؤيته بالاسم والوظيفة والمنطقة التأمينية.")
        recovery_settings_page()
        st.markdown("---")
        change_password_page()

    st.markdown("---")
    if st.button("⬅️ العودة للرئيسية", use_container_width=True, key="personal_control_back_v3"):
        st.session_state.page = "الرئيسية"
        st.rerun()
    if st.button("🚪 تسجيل الخروج", use_container_width=True, key="personal_control_logout_v3"):
        st.session_state.user = None
        st.session_state.role = None
        st.session_state.page = "login"
        st.rerun()

# ===== عرض الرسائل والبيانات أعلى الصفحة =====
def _early_load_banners():
    try:
        data = load_json_persistent("banners", BANNERS_FILE, [])
        return data if isinstance(data, list) else []
    except Exception:
        return []

def show_banners():
    if not st.session_state.get("user"):
        return
    username = st.session_state.user.get("username", "")
    now = datetime.now()
    for b in _early_load_banners():
        try:
            expire = datetime.fromisoformat(str(b.get("expire_at") or b.get("expire")))
        except Exception:
            continue
        if expire <= now:
            continue
        audience = b.get("audience", "الكل")
        if audience not in ("الكل", "كل الأعضاء") and username not in b.get("visible_to", []):
            continue
        st.markdown(
            f"<div style='background:linear-gradient(90deg,{b.get('color','#D4AF37')},#ffffff22);color:#fff;padding:16px;border-radius:16px;margin:0 0 14px;font-weight:900;font-size:20px;text-align:center;border:2px solid {b.get('color','#D4AF37')};'>📢 {b.get('text','')}</div>",
            unsafe_allow_html=True
        )
# ======= الجزء الاول: الاساسيات ============
# ============================================
import arabic_reshaper
from bidi.algorithm import get_display
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from fpdf import FPDF

st.markdown('\n<style id="visible-controls-v4">\nbutton,.stButton button,.stDownloadButton button,\ndiv[data-testid="stFileUploader"] button,\ndiv[data-testid="stTabs"] button,\ndiv[data-testid="stExpander"] button {\n    opacity:1!important;\n    visibility:visible!important;\n}\nbutton p,button span,.stButton button p,.stButton button span,\n.stDownloadButton button p,.stDownloadButton button span {\n    opacity:1!important;\n    visibility:visible!important;\n    color:#fff!important;\n}\n</style>\n', unsafe_allow_html=True)
st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

# ===== تثبيت ظهور أسماء الأزرار والأيقونات بوضوح ومنع تداخل الجداول =====
st.markdown("""
<style>
.stButton > button,
[data-testid="stButton"] > button,
[data-testid="stDownloadButton"] > button,
[data-testid="stFormSubmitButton"] > button {
    background: linear-gradient(135deg,#16243d,#243b63) !important;
    color: #ffffff !important;
    border: 2px solid #D4AF37 !important;
    border-radius: 15px !important;
    font-weight: 900 !important;
    min-height: 54px !important;
    opacity: 1 !important;
    box-shadow: 0 6px 16px rgba(0,0,0,.22) !important;
    white-space: normal !important;
    word-break: normal !important;
    overflow-wrap: break-word !important;
}
.stButton > button:hover,
[data-testid="stButton"] > button:hover,
[data-testid="stDownloadButton"] > button:hover {
    background: linear-gradient(135deg,#203455,#2e4c7c) !important;
    color:#ffffff !important;
    border-color:#f0cf62 !important;
}
.stButton > button p,
[data-testid="stButton"] > button p,
[data-testid="stDownloadButton"] > button p,
.stButton > button span,
[data-testid="stButton"] > button span {
    color:#ffffff !important;
    opacity:1 !important;
    font-weight:900 !important;
    white-space: normal !important;
    word-break: normal !important;
    overflow-wrap: break-word !important;
}
.stButton > button svg,
[data-testid="stButton"] > button svg,
[data-testid="stDownloadButton"] > button svg {
    color:#ffffff !important;
    fill:#ffffff !important;
    stroke:#ffffff !important;
    opacity:1 !important;
}
div[data-testid="stTabs"] button,
div[data-testid="stTabs"] button p,
div[data-testid="stTabs"] button span {
    color:#ffffff !important;
    font-weight:900 !important;
    opacity:1 !important;
    visibility:visible !important;
}
div[data-testid="stTabs"] button[aria-selected="true"],
div[data-testid="stTabs"] button[aria-selected="true"] p,
div[data-testid="stTabs"] button[aria-selected="true"] span {
    color:#FFD700 !important;
    background:rgba(255,215,0,0.14) !important;
}
div[data-testid="stRadio"] label,
div[data-testid="stRadio"] label p,
div[data-testid="stRadio"] label span {
    color:#ffffff !important;
    font-weight:900 !important;
    opacity:1 !important;
    visibility:visible !important;
}
div[data-testid="stRadio"] label[data-checked="true"],
div[data-testid="stRadio"] label[data-checked="true"] p,
div[data-testid="stRadio"] label[data-checked="true"] span {
    color:#FFD700 !important;
}
div[data-testid="stTextInput"] label,
div[data-testid="stPasswordInput"] label,
div[data-testid="stDateInput"] label,
div[data-testid="stSelectbox"] label,
div[data-testid="stTextArea"] label {
    color:#ffffff !important;
    font-weight:900 !important;
    white-space: normal !important;
}
.table-scroll-wrap {
    width: 100%;
    overflow-x: auto;
    -webkit-overflow-scrolling: touch;
    margin-bottom: 10px;
}
.case-table {
    width: 100%;
    min-width: 780px;
    border-collapse: collapse;
    font-size: 12px;
    color: #fff;
    text-align: center;
    direction: rtl;
}
.case-table th {
    background: #D4AF37;
    color: #0B1426;
    padding: 8px 6px;
    font-weight: 900;
    border: 1px solid #8B6914;
    white-space: nowrap;
}
.case-table td {
    background: #1E2A47;
    padding: 8px 6px;
    border: 1px solid #D4AF37;
    vertical-align: middle;
    white-space: normal;
    overflow-wrap: break-word;
    word-break: normal;
    line-height: 1.45;
}
.nowrap-cell {
    white-space: nowrap !important;
}
</style>
""", unsafe_allow_html=True)

def fix_arabic(text):
    if not text:
        return ""
    text = str(text)
    reshaped_text = arabic_reshaper.reshape(text)
    bidi_text = get_display(reshaped_text)
    return bidi_text

def _report_font_path():
    for p in (
        "Cairo-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf",
    ):
        if os.path.exists(p):
            return p
    return None

def report_download_name(prefix, member_name="", specialization="", region="", ext=""):
    parts = [str(prefix or "").strip(), str(member_name or "").strip(), str(specialization or "").strip(), str(region or "").strip()]
    clean = []
    for p in parts:
        p = re.sub(r'[\\/:*?"<>|]+', "_", p)
        if p:
            clean.append(p)
    name = " - ".join(clean) or "تقرير"
    return name + (f".{ext}" if ext else "")

def _set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def _set_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))

def _set_row_header(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:tblHeader'))

def _set_table_fixed_layout(table, total_width_inches=6.5):
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    tblPr = table._tbl.tblPr
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'fixed')
    tblPr.append(tblLayout)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(int(total_width_inches * 1440)))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)
    bidi_tbl = OxmlElement('w:bidiVisual')
    tblPr.append(bidi_tbl)

def _format_lawyer_party(lawyer_name):
    s = str(lawyer_name or "").strip()
    if not s:
        return ""
    s = re.sub(r'^طرف\s+الاستاذ(?:/|\s+المحامي)?\s*[:\-]?\s*', '', s).strip()
    s = re.sub(r'^الأستاذ(?:/|\s+المحامي)?\s*[:\-]?\s*', '', s).strip()
    s = re.sub(r'^الاستاذ(?:/|\s+المحامي)?\s*[:\-]?\s*', '', s).strip()
    s = re.sub(r'^المحامي(?:/|\s+الأستاذ)?\s*[:\-]?\s*', '', s).strip()
    s = re.sub(r'\s+المحامي$', '', s).strip()
    return f"طرف الاستاذ / {s} المحامي"

def _report_signatures_html(عضو_قانوني, مدير_ادارة, مدير_عام, تحرير_تاريخ=None, dark_theme=True):
    dt_str = fmt_date_report(تحرير_تاريخ or datetime.now().date())
    main_color = "#B8860B" if dark_theme else "#1E2A47"
    sub_color = "#FFFFFF" if dark_theme else "#000000"
    gm_color = "#ff6b6b" if dark_theme else "#c0392b"
    return f"""
    <div style="margin-top:25px; color:{main_color}; font-size:12px; width:100%; direction:rtl;">
        <p style="text-align:center; margin-bottom:18px; font-size:13px; font-weight:700;">وتفضلوا بقبول وافر الاحترام والتقدير،</p>
        <table style="width:100%; table-layout:fixed; border-collapse:collapse; border:none; direction:rtl;">
            <tr>
                <td style="width:42%; text-align:right; vertical-align:top; border:none; padding:6px 8px; overflow-wrap:break-word; word-break:normal;">
                    <div style="font-weight:900; font-size:13px; color:{main_color};">العضو القانوني</div>
                    <div style="color:{sub_color}; font-weight:700; margin-top:4px;">{html.escape(str(عضو_قانوني or ''))}</div>
                    <div style="margin-top:12px; color:{sub_color};">....................</div>
                    <div style="margin-top:14px; font-weight:800; color:{sub_color};">تحرر في : {html.escape(dt_str)}</div>
                </td>
                <td style="width:16%; border:none;"></td>
                <td style="width:42%; text-align:left; vertical-align:top; border:none; padding:6px 8px; overflow-wrap:break-word; word-break:normal;">
                    <div style="font-weight:900; font-size:13px; color:{main_color};">مدير إدارة القضايا</div>
                    <div style="color:{sub_color}; font-weight:700; margin-top:4px;">{html.escape(str(مدير_ادارة or ''))}</div>
                    <div style="margin-top:12px; color:{sub_color};">....................</div>
                </td>
            </tr>
        </table>
        <div style="text-align:center; margin-top:20px; padding:6px 10px; overflow-wrap:break-word; word-break:normal;">
            <div style="font-weight:900; font-size:13px; color:{gm_color};">مدير عام الإدارات القانونية</div>
            <div style="color:{sub_color}; font-weight:700; margin-top:4px;">{html.escape(str(مدير_عام or ''))}</div>
            <div style="margin-top:12px; color:{sub_color};">....................</div>
        </div>
    </div>
    """

def to_excel(df, title="التقرير"):
    if df is None or df.empty:
        df = pd.DataFrame([{"البيان": "لا توجد بيانات"}])
    df = df.fillna("").astype(str).copy()
    has_serial = "م" in df.columns
    if not has_serial:
        df.insert(0, "م", range(1, len(df) + 1))
    formatted_rows = []
    for _, r in df.iterrows():
        new_r = {}
        for col in df.columns:
            val = str(r[col] or "")
            if re.fullmatch(r"\d{1,2}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{4}", val.strip()):
                val = fmt_date_report(val.strip())
            new_r[col] = val
        formatted_rows.append(new_r)
    df_fmt = pd.DataFrame(formatted_rows)
    df_rtl = df_fmt.iloc[:, ::-1] if len(df_fmt.columns) > 1 else df_fmt
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_rtl.to_excel(writer, index=False, sheet_name="التقرير")
        ws = writer.sheets["التقرير"]
        ws.sheet_view.rightToLeft = True
        ws.freeze_panes = "A2"
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
                cell.font = Font(name="Arial", size=10)
        for cell in ws[1]:
            cell.font = Font(name="Arial", size=10, bold=True)
        for col in ws.columns:
            letter = col[0].column_letter
            length = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[letter].width = min(max(length + 3, 10), 45)
    return output.getvalue()

def to_word(df, title, region, member_name="", specialization="", manager_name="", general_manager=""):
    if df is None or df.empty:
        df = pd.DataFrame([{"البيان": "لا توجد بيانات"}])
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Inches(0.6)
        sec.bottom_margin = Inches(0.6)
        sec.left_margin = Inches(0.6)
        sec.right_margin = Inches(0.6)

    def rtl(text, size=9.5, bold=False, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=2):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.1
        p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
        r = p.add_run(str(text))
        r.font.name = "Arial"
        r.font.size = Pt(size)
        r.bold = bold
        return p

    rtl("الهيئة القومية للتأمين الاجتماعي", 13, True, WD_ALIGN_PARAGRAPH.CENTER, 2)
    rtl("الإدارة المركزية للإدارات القانونية", 10.5, True, WD_ALIGN_PARAGRAPH.CENTER, 1)
    rtl("الإدارة العامة للقضايا", 10.5, True, WD_ALIGN_PARAGRAPH.CENTER, 1)
    if region:
        rtl(f"ديوان عام منطقة {region}", 10, True, WD_ALIGN_PARAGRAPH.CENTER, 2)
    for tline in str(title or "").splitlines():
        if tline.strip():
            rtl(tline.strip(), 10.5, True, WD_ALIGN_PARAGRAPH.CENTER, 4)

    num_cols = max(1, len(df.columns))
    usable_w = 7.3
    col_width = Inches(usable_w / num_cols)

    table = doc.add_table(rows=1, cols=num_cols)
    table.style = "Table Grid"
    _set_table_fixed_layout(table, usable_w)
    _set_row_header(table.rows[0])
    _set_row_cant_split(table.rows[0])

    for i, col in enumerate(df.columns):
        cell = table.rows[0].cells[i]
        cell.width = col_width
        _set_cell_margins(cell, top=50, bottom=50, left=60, right=60)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
        r = p.add_run(str(col))
        r.font.name = "Arial"
        r.font.size = Pt(8.5)
        r.bold = True

    for _, row in df.iterrows():
        cells = table.add_row().cells
        _set_row_cant_split(table.rows[-1])
        for i, val in enumerate(row):
            cell = cells[i]
            cell.width = col_width
            _set_cell_margins(cell, top=45, bottom=45, left=60, right=60)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.05
            p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
            r = p.add_run(str(val))
            r.font.name = "Arial"
            r.font.size = Pt(8.5)

    rtl(f"تحريرًا في : {fmt_date_report(datetime.now().date())}", 9, False, WD_ALIGN_PARAGRAPH.RIGHT, 4)
    rtl("وتفضلوا بقبول وافر الاحترام والتقدير،", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, 6)

    sig_tbl = doc.add_table(rows=1, cols=3)
    _set_table_fixed_layout(sig_tbl, usable_w)
    _set_row_cant_split(sig_tbl.rows[0])

    right_cell = sig_tbl.rows[0].cells[0]
    gap_cell = sig_tbl.rows[0].cells[1]
    left_cell = sig_tbl.rows[0].cells[2]

    right_cell.width = Inches(2.75)
    gap_cell.width = Inches(1.8)
    left_cell.width = Inches(2.75)

    _set_cell_margins(right_cell, top=40, bottom=40, left=60, right=60)
    _set_cell_margins(gap_cell, top=40, bottom=40, left=60, right=60)
    _set_cell_margins(left_cell, top=40, bottom=40, left=60, right=60)

    p_r = right_cell.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.space_before = Pt(0)
    p_r.paragraph_format.space_after = Pt(0)
    p_r.paragraph_format.line_spacing = 1.15
    p_r._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    run_r = p_r.add_run(f"العضو القانوني\n{member_name}\n....................")
    run_r.font.name = "Arial"
    run_r.font.size = Pt(9.5)
    run_r.bold = True

    p_l = left_cell.paragraphs[0]
    p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_l.paragraph_format.space_before = Pt(0)
    p_l.paragraph_format.space_after = Pt(0)
    p_l.paragraph_format.line_spacing = 1.15
    p_l._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    run_l = p_l.add_run(f"مدير إدارة القضايا\n{manager_name}\n....................")
    run_l.font.name = "Arial"
    run_l.font.size = Pt(9.5)
    run_l.bold = True

    rtl("مدير عام الإدارات القانونية", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, 1)
    rtl(general_manager, 9.5, False, WD_ALIGN_PARAGRAPH.CENTER, 1)
    rtl("....................", 9.5, False, WD_ALIGN_PARAGRAPH.CENTER, 2)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def _pdf_split_lines(pdf, text, col_w, font_name, font_size):
    s = "" if text is None else str(text)
    if not s:
        return [""]
    pdf.set_font(font_name, "", font_size)
    max_w = max(col_w - 2.4, 4.0)
    lines = []

    def _break_long_word(word):
        chunks = []
        cur = ""
        for ch in word:
            test = cur + ch
            try:
                w = pdf.get_string_width(fix_arabic(test))
            except Exception:
                w = len(test) * (font_size * 0.35)
            if w <= max_w or not cur:
                cur = test
            else:
                chunks.append(cur)
                cur = ch
        if cur:
            chunks.append(cur)
        return chunks

    for para in s.split("\n"):
        if not para:
            lines.append("")
            continue
        words = para.split(" ")
        cur = ""
        for word in words:
            candidate = f"{cur} {word}".strip() if cur else word
            try:
                w = pdf.get_string_width(fix_arabic(candidate))
            except Exception:
                w = len(candidate) * (font_size * 0.35)
            if w <= max_w:
                cur = candidate
            else:
                if cur:
                    lines.append(cur)
                sub_chunks = _break_long_word(word)
                for sc in sub_chunks[:-1]:
                    lines.append(sc)
                cur = sub_chunks[-1] if sub_chunks else ""
        if cur:
            lines.append(cur)
    return lines or [""]

def _protect_pdf_dates(text):
    """يحافظ على فواصل التاريخ (DD . MM . YYYY) بوضوح عند تحويل النص للـ PDF."""
    s = "" if text is None else str(text)
    return re.sub(r'(\d{1,2})\s*[./]\s*(\d{1,2})\s*[./]\s*(\d{4})', r'\1 . \2 . \3', s)

def to_pdf(df, title, region, member_name="", specialization="", manager_name="", general_manager=""):
    if df is None or df.empty:
        df = pd.DataFrame([{"البيان": "لا توجد بيانات"}])
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    fp = _report_font_path()
    font = "ArabicReport" if fp else "Helvetica"
    if fp:
        pdf.add_font(font, "", fp, uni=True)

    def ar(v):
        s = _protect_pdf_dates(v)
        try:
            return fix_arabic(s)
        except Exception:
            return s

    pdf.set_font(font, "", 12)
    pdf.cell(0, 7, ar("الهيئة القومية للتأمين الاجتماعي"), 0, 1, "C")
    pdf.set_font(font, "", 10)
    pdf.cell(0, 6, ar("الإدارة المركزية للإدارات القانونية"), 0, 1, "C")
    pdf.cell(0, 6, ar("الإدارة العامة للقضايا"), 0, 1, "C")
    if region:
        pdf.cell(0, 6, ar(f"ديوان عام منطقة {region}"), 0, 1, "C")
    for tline in _pdf_split_lines(pdf, _protect_pdf_dates(title), 275, font, 10.5):
        pdf.set_font(font, "", 10.5)
        pdf.cell(0, 6, ar(tline), 0, 1, "C")
    pdf.ln(2)

    if len(df.columns):
        df_rtl = df.iloc[:, ::-1] if len(df.columns) > 1 else df
        body_font_size = 7.2
        line_h = 3.8
        x_left_margin = 10.0
        usable_w = 277.0
        width = usable_w / len(df_rtl.columns)

        def _draw_header():
            pdf.set_font(font, "", body_font_size)
            pdf.set_x(x_left_margin)
            for col in df_rtl.columns:
                pdf.cell(width, 6.5, ar(col), 1, 0, "C")
            pdf.ln()

        _draw_header()
        for _, row in df_rtl.iterrows():
            cell_lines = [_pdf_split_lines(pdf, _protect_pdf_dates(val), width, font, body_font_size) for val in row]
            max_lines = max((len(cl) for cl in cell_lines), default=1)
            row_h = max(max_lines * line_h + 1.4, 6.2)
            if pdf.get_y() + row_h > pdf.h - 12:
                pdf.add_page()
                _draw_header()
            x_start = x_left_margin
            y_start = pdf.get_y()
            for idx_c, cl in enumerate(cell_lines):
                cx = x_start + idx_c * width
                pdf.rect(cx, y_start, width, row_h)
                for idx_l, ltxt in enumerate(cl):
                    ly = y_start + 0.7 + idx_l * line_h
                    pdf.set_xy(cx + 1.0, ly)
                    pdf.cell(max(width - 2.0, 2.0), line_h, ar(ltxt), 0, 0, "R")
            pdf.set_xy(x_start, y_start + row_h)

    pdf.ln(4)
    pdf.set_font(font, "", 8.5)
    pdf.cell(0, 5.5, ar(f"تحريرًا في : {fmt_date_report(datetime.now().date())}"), 0, 1, "R")
    pdf.ln(2)
    pdf.cell(0, 5.5, ar("وتفضلوا بقبول وافر الاحترام والتقدير،"), 0, 1, "C")
    pdf.ln(3)
    y_sig = pdf.get_y()
    pdf.set_xy(165, y_sig)
    pdf.cell(122, 5.5, ar("العضو القانوني"), 0, 0, "R")
    pdf.set_xy(10, y_sig)
    pdf.cell(122, 5.5, ar("مدير إدارة القضايا"), 0, 1, "L")

    pdf.set_xy(165, y_sig + 5.2)
    pdf.cell(122, 5.5, ar(member_name or ""), 0, 0, "R")
    pdf.set_xy(10, y_sig + 5.2)
    pdf.cell(122, 5.5, ar(manager_name or ""), 0, 1, "L")

    pdf.set_xy(165, y_sig + 10.4)
    pdf.cell(122, 5.5, ar("...................."), 0, 0, "R")
    pdf.set_xy(10, y_sig + 10.4)
    pdf.cell(122, 5.5, ar("...................."), 0, 1, "L")

    pdf.set_y(y_sig + 17.5)
    pdf.cell(0, 5.5, ar("مدير عام الإدارات القانونية"), 0, 1, "C")
    pdf.cell(0, 5.5, ar(general_manager or ""), 0, 1, "C")
    pdf.cell(0, 5.5, ar("...................."), 0, 1, "C")
    raw = pdf.output()
    return raw.encode("latin-1", errors="replace") if isinstance(raw, str) else bytes(raw)
import streamlit as st
import json
import os
import re
import html
import hashlib
from datetime import datetime, timedelta, date
import pandas as pd
import io
import base64
import secrets

# تحميل ملف البيانات الرئيسي محلياً وسحابياً
DATA_FILE = os.path.join(LOCAL_DATA_DIR, "cases_data.json")
LOCAL_CASES_BACKUP = LOCAL_BACKUP_DIR / "cases_data_v2.json"

# ============================================
# ===== دوال توحيد التواريخ وتحميل وحفظ البيانات =====
# ============================================

def _load_cases_local_only():
    for candidate in (Path(DATA_FILE), LOCAL_CASES_BACKUP):
        try:
            if candidate.exists():
                loaded = json.loads(candidate.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    loaded.setdefault("cases", [])
                    loaded.setdefault("library", [])
                    return loaded
        except Exception:
            continue
    return {"cases": [], "library": []}


def _save_cases_local_only(data):
    serialized = json.dumps(data, ensure_ascii=False, indent=2, default=str)
    for target in (Path(DATA_FILE), LOCAL_CASES_BACKUP):
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = target.with_name(f".{target.name}.{secrets.token_hex(4)}.tmp")
        tmp_path.write_text(serialized, encoding="utf-8")
        os.replace(str(tmp_path), str(target))


def load_data():
    local_data = _load_cases_local_only()
    cloud_row = _cloud_get("cases_data_v1")
    if cloud_row and isinstance(cloud_row.get("payload"), dict):
        cloud_data = cloud_row["payload"]
        cloud_cases = cloud_data.get("cases") or []
        cloud_lib = cloud_data.get("library") or []
        local_cases = local_data.get("cases") or []
        local_lib = local_data.get("library") or []

        if not local_cases and cloud_cases:
            local_data["cases"] = cloud_cases
        elif cloud_cases:
            seen_ids = {str(c.get("id")) for c in local_data["cases"] if c.get("id") is not None}
            for c in cloud_cases:
                cid = str(c.get("id"))
                if cid and cid not in seen_ids:
                    local_data["cases"].append(c)
                    seen_ids.add(cid)

        if not local_lib and cloud_lib:
            local_data["library"] = cloud_lib
        elif cloud_lib:
            seen_lids = {str(item.get("id")) for item in local_data["library"] if item.get("id") is not None}
            for item in cloud_lib:
                lid = str(item.get("id"))
                if lid and lid not in seen_lids:
                    local_data["library"].append(item)
                    seen_lids.add(lid)

        try:
            _save_cases_local_only(local_data)
        except Exception:
            pass

    local_data.setdefault("cases", [])
    local_data.setdefault("library", [])
    return local_data


def save_data(data):
    """
    حفظ آمن محلياً أولاً مع نسخة احتياطية مستقلة، ثم مزامنة السحابة بأمان.
    """
    data.setdefault("cases", [])
    data.setdefault("library", [])
    try:
        _save_cases_local_only(data)
    except Exception as e:
        st.error(f"❌ تعذر الحفظ المحلي للبيانات: {e}")
        return False

    if not CLOUD_ENABLED:
        return True

    try:
        ok = bool(_cloud_put("cases_data_v1", data))
    except Exception as exc:
        st.info(f"💾 تم الحفظ محليًا بنجاح (تعذرت مزامنة السحابة مؤقتًا: {exc})")
        return True

    if not ok:
        st.info(f'💾 تم الحفظ محليًا بنجاح (لم تتم مزامنة السحابة: {_CLOUD_LAST_ERROR or "غير متاحة حاليًا"})')
    return True


def _normalize_case_dates(case_dict):
    for f in ("تاريخ_جلسة", "تاريخ_حكم", "تاريخ_قرار", "تاريخ_ورود"):
        if case_dict.get(f):
            case_dict[f] = fmt_date_ar(case_dict[f])
    for g in case_dict.get("جلسات") or []:
        if isinstance(g, dict) and g.get("تاريخ"):
            g["تاريخ"] = fmt_date_ar(g["تاريخ"])
    for h in case_dict.get("أحكام") or []:
        if isinstance(h, dict) and h.get("تاريخ"):
            h["تاريخ"] = fmt_date_ar(h["تاريخ"])
    for q in case_dict.get("قرارات") or []:
        if isinstance(q, dict) and q.get("تاريخ"):
            q["تاريخ"] = fmt_date_ar(q["تاريخ"])
    return case_dict


def get_user_cases(data, user):
    if not user:
        return []
    role = user.get("role")
    if role == "admin":
        return data.get("cases", [])
    username = user.get("username", "")
    uid = user.get("id")
    return [
        c for c in data.get("cases", [])
        if c.get("assigned_to") == username or c.get("user_id") == uid
    ]


def check_duplicate_case(cases, case_number, case_year, court_name):
    for c in cases:
        if (
            str(c.get("رقم_القضية", "")).strip() == str(case_number).strip()
            and str(c.get("سنة_القضية", "")).strip() == str(case_year).strip()
            and str(c.get("المحكمة", "")).strip() == str(court_name).strip()
        ):
            return True
    return False


def _save_uploaded_file_bytes(file_obj):
    if not file_obj:
        return None
    try:
        raw = file_obj.getvalue()
        return {
            "name": file_obj.name,
            "type": file_obj.type or "application/octet-stream",
            "data": base64.b64encode(raw).decode("utf-8"),
        }
    except Exception:
        return None


def _get_case_files_list(case):
    files = list(case.get("ملفات") or [])
    if not files and case.get("مستند"):
        m = case.get("مستند")
        if isinstance(m, dict) and m.get("data"):
            files.append(m)
        elif isinstance(m, str) and m.strip():
            files.append({"name": m.strip(), "type": "application/octet-stream", "data": ""})
    return files


def _render_case_files_section(case, key_prefix="case_file"):
    files = _get_case_files_list(case)
    if not files:
        return
    st.markdown("**📎 المرفقات والمستندات:**")
    for idx, f_item in enumerate(files):
        if isinstance(f_item, dict):
            fname = f_item.get("name") or f"مرفق_{idx+1}"
            fdata = f_item.get("data") or ""
            ftype = f_item.get("type") or "application/octet-stream"
            if fdata:
                try:
                    raw_bytes = base64.b64decode(fdata)
                    st.download_button(
                        f"⬇️ تنزيل: {fname}",
                        data=raw_bytes,
                        file_name=fname,
                        mime=ftype,
                        key=f"{key_prefix}_{case.get('id', 'x')}_{idx}",
                        use_container_width=True,
                    )
                except Exception:
                    st.caption(f"• {fname}")
            else:
                st.caption(f"• {fname}")
        else:
            st.caption(f"• {f_item}")


# ============================================
# ===== التنسيق العام وتوحيد مظهر الواجهة =====
# ============================================

st.markdown("""
<style>
.main, .stApp {
    background: linear-gradient(160deg, #091122 0%, #101e38 55%, #0c172c 100%) !important;
    color: #ffffff !important;
}
.block-container {
    padding-top: 1.2rem !important;
    padding-bottom: 2.2rem !important;
    max-width: 1120px !important;
}
/* توحيد أزرار الصفحة الرئيسية تحت بعضها في وسط الصفحة */
.home-btn-wrap .stButton > button {
    width: 100% !important;
    min-height: 62px !important;
    font-size: 18px !important;
    font-weight: 900 !important;
    border-radius: 18px !important;
    border: 2px solid #D4AF37 !important;
    background: linear-gradient(135deg, #13223e 0%, #1f365e 100%) !important;
    color: #ffffff !important;
    margin-bottom: 4px !important;
}
.home-btn-wrap .stButton > button:hover {
    background: linear-gradient(135deg, #1d335c 0%, #2c4c85 100%) !important;
    border-color: #ffe07d !important;
}
</style>
""", unsafe_allow_html=True)


def _ensure_session_defaults():
    if "user" not in st.session_state:
        st.session_state.user = None
    if "role" not in st.session_state:
        st.session_state.role = None
    if "page" not in st.session_state:
        st.session_state.page = "login"
    if "data_cache" not in st.session_state or not isinstance(st.session_state.data_cache, dict):
        st.session_state.data_cache = load_data()


_ensure_session_defaults()


def _refresh_data_cache():
    st.session_state.data_cache = load_data()
    return st.session_state.data_cache


# ============================================
# ===== دالة الاختيار الآمن للتقارير (بدون خطأ StreamlitWidgetAlreadyInstantiatedError) =====
# ============================================

def _selection_checkboxes(options, key_prefix, active_idx_key, horizontal=False):
    """
    اختيار عنصر واحد من قائمة عبر checkboxes بدون تعديل st.session_state الخاص بالـ widget
    بعد إنشائه في نفس الرن، مما يمنع StreamlitWidgetAlreadyInstantiatedError نهائياً.
    """
    if not options:
        return None, 0

    if active_idx_key not in st.session_state:
        st.session_state[active_idx_key] = 0

    prev_idx = int(st.session_state.get(active_idx_key, 0))
    prev_idx = max(0, min(prev_idx, len(options) - 1))

    # مفتاح رن لتغيير حالة الـ checkboxes عند تغيير الاختيار السابق
    epoch_key = f"{key_prefix}_epoch"
    if epoch_key not in st.session_state:
        st.session_state[epoch_key] = 0
    epoch = int(st.session_state[epoch_key])

    checked_values = []
    if horizontal:
        # في الموبايل نجعلها عمودية أو متناسقة دون تزاحم
        for idx, opt in enumerate(options):
            cb_key = f"{key_prefix}_cb_{epoch}_{idx}"
            val = st.checkbox(
                str(opt),
                value=(idx == prev_idx),
                key=cb_key,
            )
            checked_values.append(bool(val))
    else:
        for idx, opt in enumerate(options):
            cb_key = f"{key_prefix}_cb_{epoch}_{idx}"
            val = st.checkbox(
                str(opt),
                value=(idx == prev_idx),
                key=cb_key,
            )
            checked_values.append(bool(val))

    selected_indices = [i for i, v in enumerate(checked_values) if v]
    new_idx = prev_idx

    if len(selected_indices) == 1:
        new_idx = selected_indices[0]
    elif len(selected_indices) > 1:
        # إذا اختار المستخدم أكثر من صندوق، نأخذ الأحدث مقارنة بالاختيار السابق
        different = [i for i in selected_indices if i != prev_idx]
        new_idx = different[0] if different else selected_indices[0]
    else:
        # إذا ألغى المستخدم الاختيار الحالي نُبقيه مفعلاً في الرن التالي
        new_idx = prev_idx

    st.session_state[active_idx_key] = new_idx
    if checked_values != [i == new_idx for i in range(len(options))]:
        st.session_state[epoch_key] = epoch + 1
        st.rerun()

    return options[new_idx], new_idx


# ============================================
# ===== الصفحة الرئيسية (الأيقونات تحت بعضها في منتصف الصفحة) =====
# ============================================

def home_page():
    user = st.session_state.get("user") or {}
    uname = user.get("display_name") or user.get("username") or "عضو"
    role = user.get("role") or "member"

    st.markdown(
        f"<div style='text-align:center;color:#FFD700;font-size:24px;font-weight:900;margin-bottom:4px;'>أهلاً {html.escape(str(uname))}</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<div style='text-align:center;color:#D4AF37;font-size:28px;font-weight:900;margin-bottom:14px;'>⚖️ إدارة القضايا</div>",
        unsafe_allow_html=True,
    )

    data = _refresh_data_cache()
    my_cases = get_user_cases(data, user)
    total_cases = len(my_cases)
    active_cases = len([c for c in my_cases if c.get("حالة_الدعوى", "متداولة") == "متداولة"])
    judgments_count = len([c for c in my_cases if c.get("حكم_صدر") or c.get("حالة_الدعوى") == "حكم صدر"])
    today_str = fmt_date_ar(datetime.now().date())
    today_sessions = len([
        c for c in my_cases
        if fmt_date_ar(c.get("تاريخ_جلسة")) == today_str
    ])

    s1, s2 = st.columns(2)
    with s1:
        st.markdown(
            f"<div class='judicial-card'><div class='judicial-card-icon'>📊</div><div style='color:#fff;font-size:28px;font-weight:900;'>{total_cases}</div><div class='judicial-card-title'>إجمالي الدعاوى</div></div>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<div class='judicial-card'><div class='judicial-card-icon'>⚖️</div><div style='color:#fff;font-size:28px;font-weight:900;'>{judgments_count}</div><div class='judicial-card-title'>أحكام صادرة</div></div>",
            unsafe_allow_html=True,
        )
    with s2:
        st.markdown(
            f"<div class='judicial-card'><div class='judicial-card-icon'>📂</div><div style='color:#fff;font-size:28px;font-weight:900;'>{active_cases}</div><div class='judicial-card-title'>دعاوى متداولة</div></div>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<div class='judicial-card'><div class='judicial-card-icon'>📅</div><div style='color:#fff;font-size:28px;font-weight:900;'>{today_sessions}</div><div class='judicial-card-title'>جلسات اليوم</div></div>",
            unsafe_allow_html=True,
        )

    st.markdown(
        "<div style='text-align:center;color:#D4AF37;font-size:22px;font-weight:900;margin:10px 0 12px;'>الأقسام</div>",
        unsafe_allow_html=True,
    )

    sections = [
        ("⚖️ تسجيل دعوى جديدة", "إضافة_دعوى"),
        ("📋 سجل الدعاوى والحصر", "سجل_الدعاوى"),
        ("🔔 تنبيهات الجلسات والقرارات", "تنبيهات_الجلسات"),
        ("📑 التقارير الحكومية", "التقارير"),
        ("📁 الأرشيف والأحكام", "الأرشيف"),
        ("📚 المكتبة القانونية", "المكتبة"),
        ("🔍 البحث المتقدم", "البحث"),
    ]

    # وضع الأزرار تحت بعضها عمودياً في منتصف الصفحة بعرض متناسق
    _, center_col, _ = st.columns([1, 4, 1])
    with center_col:
        st.markdown("<div class='home-btn-wrap'>", unsafe_allow_html=True)
        for label, target_page in sections:
            if st.button(label, key=f"home_sec_{target_page}", use_container_width=True):
                st.session_state.page = target_page
                st.rerun()

        if role == "admin":
            if st.button("⚙️ لوحة التحكم", key="home_admin_panel", use_container_width=True):
                st.session_state.page = "لوحة التحكم"
                st.rerun()
        else:
            if st.button("⚙️ لوحة تحكمى", key="home_personal_panel", use_container_width=True):
                st.session_state.page = "لوحة تحكمى"
                st.rerun()

        if st.button("📢 الرسائل والبيانات", key="home_messages_btn", use_container_width=True):
            st.session_state.page = "الرسائل"
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    show_member_directory()


# ============================================
# ===== صفحات الدعاوى والجلسات والأرشيف والمكتبة والبحث =====
# ============================================

def add_case_page():
    st.markdown("<h2 style='text-align:center;color:#D4AF37;'>⚖️ تسجيل دعوى جديدة</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True, key="back_home_add_case"):
        st.session_state.page = "الرئيسية"
        st.rerun()

    data = _refresh_data_cache()
    user = st.session_state.get("user") or {}
    users = _cached_users_snapshot()
    member_names = [u.get("username") for u in users if u.get("role") != "admin" and u.get("username")]

    with st.form("add_case_form", clear_on_submit=False):
        case_number = st.text_input("رقم القضية", key="new_case_num")
        case_year = st.text_input("سنة القضية", value=str(datetime.now().year), key="new_case_year")
        court_name = st.text_input("المحكمة", key="new_case_court")
        circuit_name = st.text_input("الدائرة", key="new_case_circuit")
        case_type = st.selectbox("نوع الدعوى", ["مدني", "إداري", "تأديبي", "معاشات واشتراكات", "استئناف", "نقض", "إدارية عليا", "أخرى"], key="new_case_type")
        subject = st.text_area("موضوع الدعوى / الطلبات", height=90, key="new_case_subject")
        plaintiff = st.text_input("المدعي / الطاعن", key="new_case_plaintiff")
        defendant = st.text_input("المدعى عليه / المطعون ضده", key="new_case_defendant")
        lawyer_party = st.text_input("طرف الأستاذ / المحامي", key="new_case_lawyer")
        session_date = date_input_dmy("تاريخ الجلسة القادمة", value=datetime.now().date(), key="new_case_session_date")
        session_decision = st.text_area("قرار / مجريات الجلسة", height=80, key="new_case_decision")
        case_status = st.selectbox("حالة الدعوى", ["متداولة", "مؤجلة", "محجوزة للحكم", "حكم صدر", "منتهية / أرشيف"], key="new_case_status")
        judgment_text = st.text_area("منطوق الحكم / القرار (إن وجد)", height=80, key="new_case_judgment")
        judgment_date = date_input_dmy("تاريخ الحكم / القرار", value=datetime.now().date(), key="new_case_judgment_date")
        region_name = st.text_input("المنطقة التأمينية", value=user.get("insurance_area") or "البحيرة", key="new_case_region")

        if user.get("role") == "admin" and member_names:
            assigned_to = st.selectbox("إسناد الدعوى للعضو القانوني", member_names, key="new_case_assigned")
        else:
            assigned_to = user.get("username") or "admin"

        uploaded_files = st.file_uploader("📎 إرفاق مستندات الدعوى", accept_multiple_files=True, key="new_case_files")

        submitted = st.form_submit_button("💾 حفظ الدعوى", type="primary", use_container_width=True)
        if submitted:
            if not str(case_number).strip() or not str(court_name).strip():
                st.error("برجاء إدخال رقم القضية والمحكمة على الأقل")
            elif check_duplicate_case(data.get("cases", []), case_number, case_year, court_name):
                st.warning("⚠️ هذه الدعوى مسجلة مسبقاً بنفس الرقم والسنة والمحكمة.")
            else:
                files_payload = []
                for uf in (uploaded_files or []):
                    saved_f = _save_uploaded_file_bytes(uf)
                    if saved_f:
                        files_payload.append(saved_f)

                new_case = _normalize_case_dates({
                    "id": secrets.token_hex(8),
                    "رقم_القضية": str(case_number).strip(),
                    "سنة_القضية": str(case_year).strip(),
                    "المحكمة": str(court_name).strip(),
                    "الدائرة": str(circuit_name).strip(),
                    "نوع_الدعوى": str(case_type).strip(),
                    "موضوع_الدعوى": str(subject).strip(),
                    "المدعي": str(plaintiff).strip(),
                    "المدعى_عليه": str(defendant).strip(),
                    "طرف_المحامي": str(lawyer_party).strip(),
                    "تاريخ_جلسة": session_date,
                    "قرار_الجلسة": str(session_decision).strip(),
                    "حالة_الدعوى": str(case_status).strip(),
                    "حكم_صدر": str(judgment_text).strip(),
                    "تاريخ_حكم": judgment_date if str(judgment_text).strip() else "",
                    "المنطقة": str(region_name).strip(),
                    "assigned_to": assigned_to,
                    "user_id": user.get("id"),
                    "ملفات": files_payload,
                    "جلسات": [{"تاريخ": session_date, "قرار": str(session_decision).strip()}] if session_date else [],
                    "created_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
                })
                data.setdefault("cases", []).insert(0, new_case)
                save_data(data)
                st.success("✅ تم تسجيل الدعوى وحفظها بنجاح")
                st.rerun()


def cases_register_page():
    st.markdown("<h2 style='text-align:center;color:#D4AF37;'>📋 سجل الدعاوى والحصر</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True, key="back_home_cases_reg"):
        st.session_state.page = "الرئيسية"
        st.rerun()

    data = _refresh_data_cache()
    user = st.session_state.get("user") or {}
    my_cases = get_user_cases(data, user)

    if not my_cases:
        st.info("لا توجد دعاوى مسجلة حالياً.")
        return

    status_filter = st.selectbox("تصفية حسب الحالة", ["الكل", "متداولة", "مؤجلة", "محجوزة للحكم", "حكم صدر", "منتهية / أرشيف"], key="reg_status_filter")
    search_txt = st.text_input("بحث برقم القضية أو المحكمة أو الموضوع أو المحامي", key="reg_search_txt")

    filtered = []
    for c in my_cases:
        if status_filter != "الكل" and c.get("حالة_الدعوى", "متداولة") != status_filter:
            continue
        if search_txt.strip():
            blob = " ".join([
                str(c.get("رقم_القضية", "")),
                str(c.get("سنة_القضية", "")),
                str(c.get("المحكمة", "")),
                str(c.get("الدائرة", "")),
                str(c.get("موضوع_الدعوى", "")),
                str(c.get("المدعي", "")),
                str(c.get("المدعى_عليه", "")),
                str(c.get("طرف_المحامي", "")),
            ])
            if _norm_ar(search_txt) not in _norm_ar(blob):
                continue
        filtered.append(c)

    st.markdown(f"**عدد الدعاوى المعروضة:** {len(filtered)}")

    for c in filtered:
        cid = c.get("id")
        with st.container(border=True):
            st.markdown(f"### ⚖️ قضية رقم {c.get('رقم_القضية', '')} لسنة {c.get('سنة_القضية', '')} — {c.get('المحكمة', '')}")
            st.write(f"**الدائرة:** {c.get('الدائرة', '') or 'غير محدد'} | **النوع:** {c.get('نوع_الدعوى', '') or 'غير محدد'} | **الحالة:** {c.get('حالة_الدعوى', 'متداولة')}")
            st.write(f"**الخصوم:** {c.get('المدعي', '')} × {c.get('المدعى_عليه', '')}")
            if c.get("طرف_المحامي"):
                st.write(f"**طرف الأستاذ / المحامي:** {c.get('طرف_المحامي', '')}")
            if c.get("موضوع_الدعوى"):
                st.write(f"**الموضوع:** {c.get('موضوع_الدعوى', '')}")
            st.write(f"**تاريخ الجلسة:** {fmt_date_ar(c.get('تاريخ_جلسة'), 'غير محدد')} | **القرار:** {c.get('قرار_الجلسة', '') or '—'}")
            if c.get("حكم_صدر"):
                st.write(f"**الحكم / القرار:** {c.get('حكم_صدر', '')} ({fmt_date_ar(c.get('تاريخ_حكم'), '')})")
            _render_case_files_section(c, key_prefix="reg_file")

            with st.expander("✏️ تعديل بيانات الدعوى أو إضافة جلسة / حكم"):
                new_num = st.text_input("رقم القضية", value=str(c.get("رقم_القضية", "")), key=f"ed_num_{cid}")
                new_yr = st.text_input("سنة القضية", value=str(c.get("سنة_القضية", "")), key=f"ed_yr_{cid}")
                new_court = st.text_input("المحكمة", value=str(c.get("المحكمة", "")), key=f"ed_crt_{cid}")
                new_circ = st.text_input("الدائرة", value=str(c.get("الدائرة", "")), key=f"ed_cir_{cid}")
                new_subj = st.text_area("الموضوع", value=str(c.get("موضوع_الدعوى", "")), key=f"ed_sub_{cid}")
                new_sdate = date_input_dmy("تاريخ الجلسة", value=c.get("تاريخ_جلسة") or datetime.now().date(), key=f"ed_sdt_{cid}")
                new_sdec = st.text_area("قرار الجلسة", value=str(c.get("قرار_الجلسة", "")), key=f"ed_sdc_{cid}")
                statuses = ["متداولة", "مؤجلة", "محجوزة للحكم", "حكم صدر", "منتهية / أرشيف"]
                cur_st = c.get("حالة_الدعوى", "متداولة")
                st_idx = statuses.index(cur_st) if cur_st in statuses else 0
                new_st = st.selectbox("حالة الدعوى", statuses, index=st_idx, key=f"ed_st_{cid}")
                new_jdg = st.text_area("منطوق الحكم / القرار", value=str(c.get("حكم_صدر", "")), key=f"ed_jdg_{cid}")
                new_jdate = date_input_dmy("تاريخ الحكم", value=c.get("تاريخ_حكم") or datetime.now().date(), key=f"ed_jdt_{cid}")

                if st.button("💾 حفظ التعديلات", key=f"save_ed_{cid}", type="primary", use_container_width=True):
                    c["رقم_القضية"] = new_num.strip()
                    c["سنة_القضية"] = new_yr.strip()
                    c["المحكمة"] = new_court.strip()
                    c["الدائرة"] = new_circ.strip()
                    c["موضوع_الدعوى"] = new_subj.strip()
                    c["تاريخ_جلسة"] = new_sdate
                    c["قرار_الجلسة"] = new_sdec.strip()
                    c["حالة_الدعوى"] = new_st
                    c["حكم_صدر"] = new_jdg.strip()
                    c["تاريخ_حكم"] = new_jdate if new_jdg.strip() else c.get("تاريخ_حكم", "")
                    sessions = list(c.get("جلسات") or [])
                    if new_sdate or new_sdec.strip():
                        sessions.append({"تاريخ": new_sdate, "قرار": new_sdec.strip()})
                    c["جلسات"] = sessions
                    _normalize_case_dates(c)
                    save_data(data)
                    st.success("✅ تم حفظ التعديلات")
                    st.rerun()

            if st.button("🗑️ حذف الدعوى", key=f"del_case_{cid}", use_container_width=True):
                data["cases"] = [x for x in data.get("cases", []) if x.get("id") != cid]
                save_data(data)
                st.rerun()


def sessions_alerts_page():
    st.markdown("<h2 style='text-align:center;color:#D4AF37;'>🔔 تنبيهات الجلسات والقرارات</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True, key="back_home_sessions"):
        st.session_state.page = "الرئيسية"
        st.rerun()

    data = _refresh_data_cache()
    user = st.session_state.get("user") or {}
    my_cases = get_user_cases(data, user)

    today = datetime.now().date()
    upcoming = []
    for c in my_cases:
        d_s = parse_date_safe(c.get("تاريخ_جلسة"))
        if d_s is not None:
            diff = (d_s - today).days
            upcoming.append((diff, d_s, c))

    upcoming.sort(key=lambda x: x[1])
    if not upcoming:
        st.info("لا توجد جلسات مسجلة حالياً.")
        return

    for diff, d_s, c in upcoming:
        if diff == 0:
            badge = "🔴 جلسة اليوم"
        elif 0 < diff <= 7:
            badge = f"🟡 خلال {diff} يوم"
        elif diff < 0:
            badge = f"⚪ جلسة سابقة ({fmt_date_ar(d_s)})"
        else:
            badge = f"🟢 قادمة ({fmt_date_ar(d_s)})"

        with st.container(border=True):
            st.markdown(f"### {badge} — قضية {c.get('رقم_القضية', '')} لسنة {c.get('سنة_القضية', '')}")
            st.write(f"**المحكمة:** {c.get('المحكمة', '')} — **الدائرة:** {c.get('الدائرة', '')}")
            st.write(f"**تاريخ الجلسة:** {fmt_date_ar(d_s)} | **القرار:** {c.get('قرار_الجلسة', '') or '—'}")


def archive_judgments_page():
    st.markdown("<h2 style='text-align:center;color:#D4AF37;'>📁 الأرشيف والأحكام</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True, key="back_home_archive"):
        st.session_state.page = "الرئيسية"
        st.rerun()

    data = _refresh_data_cache()
    user = st.session_state.get("user") or {}
    my_cases = get_user_cases(data, user)
    archived = [
        c for c in my_cases
        if c.get("حكم_صدر") or c.get("حالة_الدعوى") in ("حكم صدر", "منتهية / أرشيف")
    ]
    if not archived:
        st.info("لا توجد أحكام أو ملفات محفوظة في الأرشيف حالياً.")
        return

    for c in archived:
        with st.container(border=True):
            st.markdown(f"### 📜 قضية {c.get('رقم_القضية', '')} لسنة {c.get('سنة_القضية', '')} — {c.get('المحكمة', '')}")
            st.write(f"**تاريخ الحكم:** {fmt_date_ar(c.get('تاريخ_حكم'), 'غير محدد')}")
            st.write(f"**منطوق الحكم / القرار:** {c.get('حكم_صدر', '') or '—'}")
            _render_case_files_section(c, key_prefix="arch_file")


def legal_library_page():
    st.markdown("<h2 style='text-align:center;color:#D4AF37;'>📚 المكتبة القانونية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True, key="back_home_library"):
        st.session_state.page = "الرئيسية"
        st.rerun()

    data = _refresh_data_cache()
    user = st.session_state.get("user") or {}
    lib_items = list(data.get("library") or [])

    with st.form("add_library_item_form"):
        title = st.text_input("عنوان المرجع / القانون / المذكرة / الحكم", key="lib_title")
        category = st.selectbox("التصنيف", ["قانون 148 لسنة 2019", "قانون 112 لسنة 1980", "قانون المرافعات", "أحكام ومبادئ قضائية", "مذكرات ودفوع", "أخرى"], key="lib_cat")
        notes = st.text_area("ملخص أو نص أو ملاحظات", height=110, key="lib_notes")
        doc_file = st.file_uploader("📎 إرفاق ملف", type=None, key="lib_file")
        if st.form_submit_button("💾 إضافة للمكتبة", type="primary", use_container_width=True):
            if not title.strip() and not notes.strip() and not doc_file:
                st.error("أدخل عنواناً أو نصاً أو ملفاً")
            else:
                saved_f = _save_uploaded_file_bytes(doc_file) if doc_file else None
                lib_items.insert(0, {
                    "id": secrets.token_hex(8),
                    "title": title.strip() or (doc_file.name if doc_file else "مرجع قانوني"),
                    "category": category,
                    "notes": notes.strip(),
                    "file": saved_f,
                    "user_id": user.get("id"),
                    "created_at": fmt_date_ar(datetime.now().date()),
                })
                data["library"] = lib_items
                save_data(data)
                st.success("✅ تمت الإضافة للمكتبة القانونية")
                st.rerun()

    if not lib_items:
        st.info("لا توجد مواد مضافة في المكتبة حالياً.")
        return

    for item in lib_items:
        with st.container(border=True):
            st.markdown(f"### 📖 {item.get('title', '')} — *{item.get('category', '')}*")
            if item.get("notes"):
                st.write(item.get("notes", ""))
            f_obj = item.get("file")
            if isinstance(f_obj, dict) and f_obj.get("data"):
                try:
                    st.download_button(
                        f"⬇️ تنزيل: {f_obj.get('name', 'ملف')}",
                        data=base64.b64decode(f_obj["data"]),
                        file_name=f_obj.get("name", "file"),
                        mime=f_obj.get("type", "application/octet-stream"),
                        key=f"lib_dl_{item.get('id', 'x')}",
                        use_container_width=True,
                    )
                except Exception:
                    pass


def advanced_search_page():
    st.markdown("<h2 style='text-align:center;color:#D4AF37;'>🔍 البحث المتقدم</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True, key="back_home_search"):
        st.session_state.page = "الرئيسية"
        st.rerun()

    data = _refresh_data_cache()
    user = st.session_state.get("user") or {}
    my_cases = get_user_cases(data, user)

    q_num = st.text_input("رقم القضية", key="adv_q_num")
    q_year = st.text_input("السنة", key="adv_q_year")
    q_court = st.text_input("المحكمة / الدائرة", key="adv_q_court")
    q_subj = st.text_input("الموضوع / الخصوم / المحامي / الحكم", key="adv_q_subj")

    results = []
    for c in my_cases:
        if q_num.strip() and _norm_ar(q_num) not in _norm_ar(c.get("رقم_القضية", "")):
            continue
        if q_year.strip() and _norm_ar(q_year) not in _norm_ar(c.get("سنة_القضية", "")):
            continue
        if q_court.strip():
            c_blob = f"{c.get('المحكمة', '')} {c.get('الدائرة', '')}"
            if _norm_ar(q_court) not in _norm_ar(c_blob):
                continue
        if q_subj.strip():
            s_blob = f"{c.get('موضوع_الدعوى', '')} {c.get('المدعي', '')} {c.get('المدعى_عليه', '')} {c.get('طرف_المحامي', '')} {c.get('حكم_صدر', '')}"
            if _norm_ar(q_subj) not in _norm_ar(s_blob):
                continue
        results.append(c)

    st.markdown(f"**نتائج البحث:** {len(results)}")
    for c in results:
        with st.container(border=True):
            st.markdown(f"### ⚖️ {c.get('رقم_القضية', '')} لسنة {c.get('سنة_القضية', '')} — {c.get('المحكمة', '')}")
            st.write(f"**الدائرة:** {c.get('الدائرة', '')} | **الحالة:** {c.get('حالة_الدعوى', 'متداولة')}")
            st.write(f"**الموضوع:** {c.get('موضوع_الدعوى', '') or '—'}")
            st.write(f"**تاريخ الجلسة:** {fmt_date_ar(c.get('تاريخ_جلسة'), 'غير محدد')} | **القرار:** {c.get('قرار_الجلسة', '') or '—'}")


# ============================================
# ===== التقارير الحكومية (بدون أخطاء واختيارات عمودية واضحة) =====
# ============================================

def _filter_cases_for_report(
    cases,
    kind="all",
    use_date_range=False,
    d_from=None,
    d_to=None,
    lawyer_filter="",
    subject_filter="",
    court_filter="",
    circuit_filter="",
    type_filter="",
):
    df_from = parse_date_safe(d_from) if use_date_range else None
    df_to = parse_date_safe(d_to) if use_date_range else None
    if df_from and df_to and df_from > df_to:
        df_from, df_to = df_to, df_from

    out = []
    for c in cases:
        status_val = str(c.get("حالة_الدعوى", "متداولة"))
        has_judgment = bool(str(c.get("حكم_صدر", "")).strip()) or (status_val == "حكم صدر")

        if kind == "judgments":
            if not has_judgment:
                continue
        elif kind == "active":
            if has_judgment or status_val == "منتهية / أرشيف":
                continue

        if use_date_range and (df_from or df_to):
            c_dates = _case_session_dates(c)
            d_j = parse_date_safe(c.get("تاريخ_حكم"))
            if d_j is not None:
                c_dates.append(d_j)
            matched_dt = False
            for dt_item in c_dates:
                ok_start = True if df_from is None else (dt_item >= df_from)
                ok_end = True if df_to is None else (dt_item <= df_to)
                if ok_start and ok_end:
                    matched_dt = True
                    break
            if not matched_dt:
                continue

        if lawyer_filter and str(lawyer_filter).strip() not in ("الكل", ""):
            if _norm_ar(lawyer_filter) not in _norm_ar(c.get("طرف_المحامي", "")):
                continue

        if subject_filter and str(subject_filter).strip():
            if _norm_ar(subject_filter) not in _norm_ar(c.get("موضوع_الدعوى", "")):
                continue

        if court_filter and str(court_filter).strip() not in ("الكل", ""):
            if _norm_ar(court_filter) not in _norm_ar(c.get("المحكمة", "")):
                continue

        if circuit_filter and str(circuit_filter).strip() not in ("الكل", ""):
            if _norm_ar(circuit_filter) not in _norm_ar(c.get("الدائرة", "")):
                continue

        if type_filter and str(type_filter).strip() not in ("الكل", ""):
            if _norm_ar(type_filter) not in _norm_ar(c.get("نوع_الدعوى", "")):
                continue

        out.append(c)

    return out


def _cases_to_report_df(cases_list, kind="all"):
    rows = []
    for idx, c in enumerate(cases_list, start=1):
        parties = f"{c.get('المدعي', '')} × {c.get('المدعى_عليه', '')}".strip(" ×")
        base_row = {
            "م": idx,
            "رقم القضية": str(c.get("رقم_القضية", "")),
            "السنة": str(c.get("سنة_القضية", "")),
            "المحكمة": str(c.get("المحكمة", "")),
            "الدائرة": str(c.get("الدائرة", "")),
            "الخصوم": parties,
            "الموضوع": str(c.get("موضوع_الدعوى", "")),
            "طرف المحامي": str(c.get("طرف_المحامي", "")),
            "تاريخ الجلسة": fmt_date_report(c.get("تاريخ_جلسة"), ""),
            "القرار / الحكم": str(c.get("حكم_صدر", "") or c.get("قرار_الجلسة", "")),
        }
        if kind == "judgments":
            base_row["تاريخ الحكم"] = fmt_date_report(c.get("تاريخ_حكم"), "")
            base_row["الحكم الصادر"] = str(c.get("حكم_صدر", "") or c.get("قرار_الجلسة", ""))
        rows.append(base_row)

    if not rows:
        return pd.DataFrame([{"م": 1, "البيان": "لا توجد دعاوى مطابقة للمعايير المحددة"}])
    return pd.DataFrame(rows)


def _render_html_report_table(df_rep):
    if df_rep is None or df_rep.empty:
        return "<div style='color:#fff;text-align:center;'>لا توجد بيانات</div>"
    headers_html = "".join(f"<th>{html.escape(str(col))}</th>" for col in df_rep.columns)
    body_rows = []
    for _, r in df_rep.iterrows():
        cells_html = "".join(f"<td>{html.escape(str(r[col] or ''))}</td>" for col in df_rep.columns)
        body_rows.append(f"<tr>{cells_html}</tr>")
    return (
        "<div class='table-scroll-wrap'>"
        "<table class='case-table'>"
        f"<thead><tr>{headers_html}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table>"
        "</div>"
    )


def reports_page():
    st.markdown("<h2 style='text-align:center;color:#D4AF37;'>📑 مركز التقارير الحكومية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True, key="back_home_reports"):
        st.session_state.page = "الرئيسية"
        st.rerun()

    data = _refresh_data_cache()
    user = st.session_state.get("user") or {}
    my_cases = get_user_cases(data, user)

    tab_statement, tab_stats, tab_numeric = st.tabs([
        "📋 بيان الدعاوى والأحكام",
        "📈 الإحصائيات",
        "📊 بيان عددي",
    ])

    with tab_statement:
        st.markdown("### نوع التقرير (اختيار عمودي واضح)")
        report_labels = [
            "بيان بجميع الدعاوى المتداولة",
            "بيان بالدعاوى المتداولة فقط",
            "بيان بالأحكام الصادرة",
        ]
        report_kinds = ["all", "active", "judgments"]
        selected_label, sel_idx = _selection_checkboxes(
            report_labels,
            key_prefix="rep_kind_select",
            active_idx_key="rep_kind_active_idx",
            horizontal=False,
        )
        active_kind = report_kinds[sel_idx]
        st.info(f"النوع المختار حالياً: **{selected_label}**")

        use_date_range = st.checkbox("📅 تصفية بفترة زمنية (من تاريخ / إلى تاريخ)", value=False, key="rep_use_dates")
        d_from = None
        d_to = None
        if use_date_range:
            d_from = date_input_dmy("من تاريخ", value=datetime.now().date().replace(day=1), key="rep_date_from")
            d_to = date_input_dmy("إلى تاريخ", value=datetime.now().date(), key="rep_date_to")

        courts_list = ["الكل"] + sorted({str(c.get("المحكمة", "")).strip() for c in my_cases if str(c.get("المحكمة", "")).strip()})
        circuits_list = ["الكل"] + sorted({str(c.get("الدائرة", "")).strip() for c in my_cases if str(c.get("الدائرة", "")).strip()})
        types_list = ["الكل"] + sorted({str(c.get("نوع_الدعوى", "")).strip() for c in my_cases if str(c.get("نوع_الدعوى", "")).strip()})
        lawyers_list = ["الكل"] + sorted({str(c.get("طرف_المحامي", "")).strip() for c in my_cases if str(c.get("طرف_المحامي", "")).strip()})

        court_sel = st.selectbox("المحكمة", courts_list, key="rep_court_sel")
        circuit_sel = st.selectbox("الدائرة", circuits_list, key="rep_circuit_sel")
        type_sel = st.selectbox("نوع الدعوى", types_list, key="rep_type_sel")
        lawyer_sel = st.selectbox("طرف الأستاذ / المحامي", lawyers_list, key="rep_lawyer_sel")
        subject_txt = st.text_input("بحث في موضوع الدعوى", key="rep_subject_txt")

        region_val = st.text_input("اسم المنطقة", value=user.get("insurance_area") or "البحيرة", key="rep_region_val")
        member_val = st.text_input("العضو القانوني", value=user.get("display_name") or user.get("username") or "", key="rep_member_val")
        manager_val = st.text_input("مدير إدارة القضايا", value="", key="rep_manager_val")
        gm_val = st.text_input("مدير عام الإدارات القانونية", value="", key="rep_gm_val")

        if st.button("📊 عرض وتجهيز البيان للطباعة", type="primary", use_container_width=True, key="rep_prepare_btn"):
            st.session_state["rep_prepared_flag"] = True

        if st.session_state.get("rep_prepared_flag"):
            filtered_cases = _filter_cases_for_report(
                my_cases,
                kind=active_kind,
                use_date_range=use_date_range,
                d_from=d_from,
                d_to=d_to,
                lawyer_filter=lawyer_sel,
                subject_filter=subject_txt,
                court_filter=court_sel,
                circuit_filter=circuit_sel,
                type_filter=type_sel,
            )
            df_out = _cases_to_report_df(filtered_cases, kind=active_kind)
            full_title = f"{selected_label}\nمنطقة {region_val}" if region_val else selected_label

            st.markdown(f"<h4 style='text-align:center;color:#FFD700;'>{html.escape(selected_label)} ({len(filtered_cases)} دعوى)</h4>", unsafe_allow_html=True)
            st.markdown(_render_html_report_table(df_out), unsafe_allow_html=True)
            st.markdown(
                _report_signatures_html(member_val, manager_val, gm_val, تحرير_تاريخ=datetime.now().date(), dark_theme=True),
                unsafe_allow_html=True,
            )

            st.markdown("### 📥 تحميل وطباعة التقرير")
            col_w, col_p, col_e = st.columns(3)
            with col_w:
                docx_bytes = to_word(df_out, full_title, region_val, member_val, selected_label, manager_val, gm_val)
                st.download_button(
                    "📄 تحميل Word",
                    data=docx_bytes,
                    file_name=report_download_name(selected_label, member_val, "", region_val, "docx"),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="dl_rep_word",
                )
            with col_p:
                pdf_bytes = to_pdf(df_out, full_title, region_val, member_val, selected_label, manager_val, gm_val)
                st.download_button(
                    "📕 تحميل PDF",
                    data=pdf_bytes,
                    file_name=report_download_name(selected_label, member_val, "", region_val, "pdf"),
                    mime="application/pdf",
                    use_container_width=True,
                    key="dl_rep_pdf",
                )
            with col_e:
                excel_bytes = to_excel(df_out, full_title)
                st.download_button(
                    "📊 تحميل Excel",
                    data=excel_bytes,
                    file_name=report_download_name(selected_label, member_val, "", region_val, "xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dl_rep_excel",
                )

    with tab_stats:
        st.markdown("### 📈 إحصائيات القضايا")
        total_c = len(my_cases)
        judg_c = len([c for c in my_cases if str(c.get("حكم_صدر", "")).strip() or c.get("حالة_الدعوى") == "حكم صدر"])
        active_c = max(0, total_c - judg_c)
        stats_df = pd.DataFrame([
            {"البند": "إجمالي الدعاوى", "العدد": total_c},
            {"البند": "الدعاوى المتداولة", "العدد": active_c},
            {"البند": "الأحكام الصادرة", "العدد": judg_c},
        ])
        st.dataframe(stats_df, use_container_width=True, hide_index=True)

    with tab_numeric:
        st.markdown("### 📊 بيان عددي حسب المحكمة")
        if not my_cases:
            st.info("لا توجد بيانات لإنشاء بيان عددي.")
        else:
            court_counts = {}
            for c in my_cases:
                crt = str(c.get("المحكمة", "")).strip() or "غير محدد"
                court_counts[crt] = court_counts.get(crt, 0) + 1
            num_df = pd.DataFrame([
                {"م": i + 1, "المحكمة": k, "عدد الدعاوى": v}
                for i, (k, v) in enumerate(sorted(court_counts.items(), key=lambda x: x[1], reverse=True))
            ])
            st.dataframe(num_df, use_container_width=True, hide_index=True)


# ============================================
# ===== التوجيه الرئيسي للتطبيق =====
# ============================================

def main():
    show_banners()
    page = st.session_state.get("page", "login")
    user = st.session_state.get("user")

    if page == "set_password":
        set_password_page()
        return

    if not user or page == "login":
        login_page()
        return

    if page == "الرئيسية":
        home_page()
    elif page == "إضافة_دعوى":
        add_case_page()
    elif page == "سجل_الدعاوى":
        cases_register_page()
    elif page == "تنبيهات_الجلسات":
        sessions_alerts_page()
    elif page == "الأرشيف":
        archive_judgments_page()
    elif page == "المكتبة":
        legal_library_page()
    elif page == "البحث":
        advanced_search_page()
    elif page == "التقارير":
        reports_page()
    elif page == "الرسائل":
        messaging_page()
    elif page == "لوحة التحكم":
        admin_control_page()
    elif page == "لوحة تحكمى":
        personal_control_page()
    elif page == "extract_member":
        extract_member_page()
    else:
        home_page()


if __name__ == "__main__":
    main()
