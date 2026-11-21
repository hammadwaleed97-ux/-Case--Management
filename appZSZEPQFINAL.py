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


# ===== دوال توحيد وتنسيق التواريخ (يوم . شهر . سنة) بقائمة منسدلة =====
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
    s_clean = re.sub(r"\s*\.\s*", "/", s_date)
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(s_clean, fmt).date()
        except Exception:
            continue
    try:
        return datetime.fromisoformat(s_clean).date()
    except Exception:
        return None


def fmt_date_ar(val, default=""):
    d = parse_date_safe(val)
    if d is None:
        return str(val) if val not in (None, "") else default
    return f"{d.day:02d} . {d.month:02d} . {d.year:04d}"


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

    c_d, c_m, c_y = st.columns(3, gap="small")
    with c_d:
        days = list(range(1, 32))
        safe_day_idx = min(max(d_val.day - 1, 0), 30)
        sel_d = st.selectbox(
            "اليوم",
            days,
            format_func=lambda d: f"{d:02d}",
            index=safe_day_idx,
            key=f"{base_key}_d",
            label_visibility="collapsed",
        )
    with c_m:
        months = list(range(1, 13))
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
/* صفحة الدخول */
.login-shell{
    max-width:760px;margin:45px auto;padding:28px;
    background:linear-gradient(145deg,#111c33,#1f3154);
    border:2px solid #D4AF37;border-radius:24px;
    box-shadow:0 20px 55px rgba(0,0,0,.45);
}
.login-shell h1,.login-shell h2,.login-shell p,.login-shell label{color:#fff!important;}
div[data-testid="stTabs"] button{
    color:#FFD700!important;font-weight:900!important;font-size:18px!important;
}
div[data-testid="stTabs"] button[aria-selected="true"]{
    color:#FFFFFF!important;background:rgba(212,175,55,0.22)!important;border-bottom:3px solid #FFD700!important;
}
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
[data-testid="stTabs"] button { color:#FFD700 !important; opacity:1 !important; visibility:visible !important; }
[data-testid="stFileUploader"] section * { color:#fff !important; opacity:1 !important; visibility:visible !important; }
input, textarea, [data-baseweb="select"], [data-testid="stFileUploader"] {color:#111!important;}
label, [data-testid="stCheckbox"] label, [data-testid="stRadio"] label, [data-testid="stMultiSelect"] label, [data-testid="stSelectbox"] label {color:#FFD700!important;font-weight:900!important;}
[data-testid="stCheckbox"] p, [data-testid="stRadio"] p {color:#FFFFFF!important;font-weight:900!important;}
[data-testid="stFileUploader"] section {background:#142038!important;border:1px solid #D4AF37!important;}
[data-testid="stButton"] button svg, [data-testid="stDownloadButton"] button svg, [data-testid="stTabs"] button svg, [data-testid="stExpander"] svg, [data-testid="stFileUploader"] svg { display:inline-block !important; visibility:visible !important; opacity:1 !important; color:#fff !important; fill:#fff !important; stroke:#fff !important; }
[data-testid="stButton"] button > div, [data-testid="stDownloadButton"] button > div { display:flex !important; align-items:center !important; justify-content:center !important; gap:.45rem !important; }

[data-testid="stFileUploader"] small, [data-testid="stFileUploader"] span {color:#fff!important;}
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
    div[data-testid="stSelectbox"] label {color:#FFD700!important;font-weight:900!important;}
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
                user["age_action_at"] = datetime.now().strftime("%d . %m . %Y %H:%M")
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
                    user["age_action_at"] = datetime.now().strftime("%d . %m . %Y %H:%M")
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
                user["age_action_at"] = datetime.now().strftime("%d . %m . %Y %H:%M")
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
            "created_at": datetime.now().strftime("%d . %m . %Y %H:%M"),
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
                        "created_at": datetime.now().strftime("%d . %m . %Y %H:%M"),
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
        cols = st.columns(6)
        for i,u in enumerate(area_users):
            with cols[i % 6]:
                name = html.escape(str(u.get("display_name") or u.get("username") or ""))
                job_value = membership_job_title(u) or role_label(u.get("role"))
                job = html.escape(job_value)
                photo = _photo_html(u.get("profile_photo"), u.get("profile_photo_type") or "image/png", 42)
                is_admin = u.get("role") == "admin"
                border = "#00E676" if is_admin else "#D4AF37"
                name_color = "#00E676" if is_admin else "#fff"
                admin_badge = "<div style='color:#00E676;font-size:10px;font-weight:900;'>ADMIN</div>" if is_admin else ""
                card = f"<div style='background:#142038;border:2px solid {border};border-radius:12px;padding:6px;text-align:center;min-height:112px;margin-bottom:5px;'>{photo}<div style='color:{name_color};font-size:12px;font-weight:900;line-height:1.2;'>{name}</div>{admin_badge}<div style='color:#D4AF37;font-size:10px;font-weight:800;line-height:1.2;'>{job}</div><div style='color:#cbd5e1;font-size:9px;line-height:1.2;'>{html.escape(area)}</div></div>"
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
                        "created_at": created.strftime("%d . %m . %Y %H:%M"),
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
def edit_memberships_page():
    if not st.session_state.user or st.session_state.user.get("role") != "admin":
        st.error("هذه الصفحة متاحة للادمن فقط")
        return

    st.markdown("<h2 style='color:#D4AF37;text-align:center;'>✏️ تعديل بيانات العضويات</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة إلى إدارة الأعضاء", key="back_to_manage_users_from_editor", use_container_width=True):
        st.session_state.page = "إدارة الأعضاء"
        st.rerun()

    users = _cached_users_snapshot()
    members = [u for u in users if u.get("role") != "admin"]
    if not members:
        st.info("لا يوجد أعضاء لتعديل بياناتهم.")
        return

    job_options = ["عضو قانونى", "مدير إدارة", "مدير عام", "قيادي مركز رئيسى 1", "قيادي مركز رئيسى 2", "قيادي مركز رئيسى 3"]
    role_map = {
        "عضو قانونى": "member",
        "مدير إدارة": "manager",
        "مدير عام": "general_manager",
        "قيادي مركز رئيسى 1": "leader1",
        "قيادي مركز رئيسى 2": "leader2",
        "قيادي مركز رئيسى 3": "leader3",
    }

    for user in members:
        uid = user.get("id")
        current_name = user.get("username", "")
        current_birth = parse_date_safe(user.get("birth_date"))
        current_area = user.get("insurance_area", "")
        current_job = user.get("job_title") or role_label(user.get("role"))
        if current_job not in job_options:
            current_job = "عضو قانونى"

        with st.container(border=True):
            st.markdown(f"### 👤 {current_name}")
            if user.get("profile_photo"):
                st.markdown(_photo_html(user.get("profile_photo"), user.get("profile_photo_type") or "image/png", 76), unsafe_allow_html=True)
            new_name = st.text_input("الاسم", value=current_name, key=f"edit_mem_name_{uid}")
            new_birth = date_input_dmy(
                "تاريخ الميلاد",
                value=current_birth or datetime.now().date().replace(year=max(1900, datetime.now().year - 30)),
                min_value=date(1900, 1, 1),
                max_value=date(2100, 12, 31),
                key=f"edit_mem_birth_{uid}"
            )
            new_area = st.text_input("المنطقة التأمينية", value=current_area, key=f"edit_mem_area_{uid}")
            new_job = st.selectbox("الوظيفة", job_options, index=job_options.index(current_job), key=f"edit_mem_job_{uid}")
            new_photo = st.file_uploader("تغيير الصورة الشخصية", type=["jpg","jpeg","png"], key=f"edit_mem_photo_{uid}")

            if st.button("💾 حفظ التعديلات", key=f"save_mem_edit_{uid}", type="primary", use_container_width=True):
                clean_name = new_name.strip()
                clean_area = new_area.strip()
                if not clean_name:
                    st.error("لا يمكن أن يكون الاسم فارغاً.")
                elif not clean_area:
                    st.error("لا يمكن أن تكون المنطقة التأمينية فارغة.")
                elif any(u.get("id") != uid and u.get("username") == clean_name for u in users):
                    st.error("يوجد عضو آخر مسجل بنفس الاسم.")
                else:
                    old_name = user.get("username", "")
                    user["username"] = clean_name
                    user["birth_date"] = fmt_date_ar(new_birth)
                    user["insurance_area"] = clean_area
                    user["job_title"] = new_job
                    user["role"] = role_map[new_job]
                    if new_photo:
                        user["profile_photo"] = base64.b64encode(new_photo.getvalue()).decode("utf-8")
                        user["profile_photo_type"] = new_photo.type
                    save_users(users)

                    if old_name and old_name != clean_name:
                        data = load_data()
                        changed = False
                        for c in data.get("cases", []):
                            if c.get("assigned_to") == old_name:
                                c["assigned_to"] = clean_name
                                changed = True
                        if changed:
                            save_data(data)

                    st.success(f"تم حفظ تعديلات العضوية بنجاح: {clean_name}")
                    st.rerun()

def recovery_settings_page():
    users = _cached_users_snapshot()
    user = next((u for u in users if u.get("id") == st.session_state.user.get("id")), None)
    if not user:
        return

    st.markdown("<h3 style='color:#D4AF37;text-align:center;'>📧 إعدادات الإيميل والاسترجاع</h3>", unsafe_allow_html=True)
    with st.container(border=True):
        primary_email = st.text_input("الإيميل الأساسي", value=user.get("email", ""), key="rec_primary_email")
        recovery_email = st.text_input("الإيميل الاحتياطي (اختياري)", value=user.get("recovery_email", ""), key="rec_backup_email")

        if st.button("💾 حفظ الإيميل", type="primary", use_container_width=True, key="save_recovery_emails"):
            if not primary_email.strip():
                st.error("لازم تكتب الإيميل الأساسي")
            else:
                user["email"] = primary_email.strip()
                user["recovery_email"] = recovery_email.strip()
                save_users(users)
                st.session_state.user = user
                st.success("تم حفظ الإيميل بنجاح")
                st.rerun()

def change_password_page():
    users = _cached_users_snapshot()
    user = next((u for u in users if u.get("id") == st.session_state.user.get("id")), None)
    if not user:
        return

    st.markdown("<h3 style='color:#D4AF37;text-align:center;'>🔑 تغيير كلمة السر</h3>", unsafe_allow_html=True)
    with st.container(border=True):
        current_pass = st.text_input("كلمة السر الحالية", type="password", key="chg_curr_pass")
        new_pass = st.text_input("كلمة السر الجديدة", type="password", key="chg_new_pass")
        confirm_pass = st.text_input("تأكيد كلمة السر الجديدة", type="password", key="chg_conf_pass")

        if st.button("تحديث كلمة السر", type="primary", use_container_width=True, key="chg_pass_btn"):
            valid_current = False
            try:
                valid_current = bool(user.get("password")) and bcrypt.checkpw(current_pass.encode(), user["password"].encode())
            except Exception:
                valid_current = False

            if not valid_current:
                st.error("كلمة السر الحالية غير صحيحة")
            elif not new_pass:
                st.error("اكتب كلمة السر الجديدة")
            elif new_pass != confirm_pass:
                st.error("كلمتا السر الجديدتان غير متطابقتين")
            else:
                user["password"] = bcrypt.hashpw(new_pass.encode(), bcrypt.gensalt()).decode()
                user["password_set"] = True
                save_users(users)
                st.session_state.user = user
                st.success("تم تحديث كلمة السر بنجاح")

def set_password_page():
    st.markdown("<h2 style='text-align:center; color:#C9A961'>تعيين كلمة السر الجديدة</h2>", unsafe_allow_html=True)
    pass1 = st.text_input("كلمة السر الجديدة", type="password", key="set_p1")
    pass2 = st.text_input("تأكيد كلمة السر", type="password", key="set_p2")
    if st.button("حفظ كلمة السر", type="primary", use_container_width=True, key="set_p_btn"):
        if not pass1 or pass1 != pass2:
            st.error("كلمة السر غير متطابقة أو فارغة")
        else:
            users = _cached_users_snapshot()
            for u in users:
                if u.get("id") == st.session_state.get("temp_user_id") or u.get("username") == st.session_state.get("temp_user"):
                    u["password"] = bcrypt.hashpw(pass1.encode(), bcrypt.gensalt()).decode()
                    u["password_set"] = True
                    st.session_state.user = u
                    break
            save_users(users)
            st.session_state.page = "الرئيسية"
            st.success("تم حفظ كلمة السر بنجاح")
            st.rerun()

# ===== البيانات والمستندات =====
DATA_FILE = os.path.join(LOCAL_DATA_DIR, "data.json")
DOCS_DIR = os.path.join(LOCAL_DATA_DIR, "docs")
os.makedirs(DOCS_DIR, exist_ok=True)
ALLOWED_DOC_EXTENSIONS = {"pdf", "doc", "docx", "xls", "xlsx", "csv", "txt", "rtf", "odt", "ods", "jpg", "jpeg", "png", "webp", "gif", "bmp"}
MAX_DOC_BYTES = 25 * 1024 * 1024

def _default_data():
    return {
        "cases": [],
        "library": [],
        "settings": {
            "app_title": "إدارة القضايا",
            "subtitle": "منظومة متكاملة لإدارة الدعاوى والقرارات والأحكام والمذكرات",
            "member_name": "",
            "manager_name": "",
            "member_title": "العضو القانوني",
            "manager_title": "مدير إدارة القضايا",
            "theme": "الأزرق الملكي والذهبي"
        }
    }

def load_data():
    d = load_json_persistent("data", DATA_FILE, _default_data())
    if not isinstance(d, dict):
        d = _default_data()
    d.setdefault("cases", [])
    d.setdefault("library", [])
    d.setdefault("settings", {})
    for k, v in _default_data()["settings"].items():
        d["settings"].setdefault(k, v)
    return d

def save_data(data):
    return save_json_persistent("data", DATA_FILE, data)

def _is_valid_uploaded_doc(uploaded_file):
    if not uploaded_file:
        return False, "لم يتم اختيار ملف."
    file_name = uploaded_file.name or ""
    ext = file_name.rsplit(".", 1)[-1].lower() if "." in file_name else ""
    if ext not in ALLOWED_DOC_EXTENSIONS:
        return False, f"نوع الملف (.{ext or 'غير معروف'}) غير مسموح به."
    raw = uploaded_file.getvalue()
    if not raw:
        return False, "الملف المرفوع فارغ."
    if len(raw) > MAX_DOC_BYTES:
        return False, "حجم الملف يتجاوز الحد الأقصى المسموح به (25 ميجابايت)."
    return True, ""

def save_uploaded_file(uploaded_file, prefix="doc"):
    ok, err = _is_valid_uploaded_doc(uploaded_file)
    if not ok:
        raise ValueError(err)
    orig_name = Path(uploaded_file.name).name
    ext = orig_name.rsplit(".", 1)[-1].lower() if "." in orig_name else "bin"
    safe_stem = re.sub(r"[^A-Za-z0-9_-]+", "_", Path(orig_name).stem).strip("_") or "file"
    fname = f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{secrets.token_hex(4)}_{safe_stem}.{ext}"
    fpath = os.path.join(DOCS_DIR, fname)
    raw = uploaded_file.getvalue()
    with open(fpath, "wb") as f:
        f.write(raw)
    storage_path = None
    if CLOUD_ENABLED:
        sp, _ = cloud_upload_file(raw, f"docs/{fname}", uploaded_file.type or "application/octet-stream")
        storage_path = sp
    return fpath, orig_name, storage_path

def _get_doc_bytes(doc_obj):
    if not isinstance(doc_obj, dict):
        return None
    fpath = doc_obj.get("path")
    if fpath and os.path.exists(fpath):
        try:
            with open(fpath, "rb") as f:
                return f.read()
        except Exception:
            pass
    sp = doc_obj.get("storage_path")
    if sp and CLOUD_ENABLED:
        try:
            r = requests.get(
                f"{SUPABASE_URL}/storage/v1/object/{SUPABASE_STORAGE_BUCKET}/{sp}",
                headers={"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}"},
                timeout=15
            )
            if r.ok:
                return r.content
        except Exception:
            pass
    return None

def _delete_doc_obj(doc_obj):
    if not isinstance(doc_obj, dict):
        return
    fpath = doc_obj.get("path")
    if fpath and os.path.exists(fpath):
        try:
            os.remove(fpath)
        except Exception:
            pass
    sp = doc_obj.get("storage_path")
    if sp:
        cloud_delete_file(sp)

def _render_doc_viewer(doc_obj, key_prefix="doc"):
    if not isinstance(doc_obj, dict):
        return
    name = doc_obj.get("name") or "مستند"
    raw = _get_doc_bytes(doc_obj)
    if not raw:
        st.warning(f"⚠️ الملف غير متوفر حالياً: {name}")
        return
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    mime_map = {
        "pdf": "application/pdf",
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "webp": "image/webp",
        "txt": "text/plain; charset=utf-8",
        "csv": "text/csv; charset=utf-8",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "doc": "application/msword",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "xls": "application/vnd.msword",
    }
    mime = mime_map.get(ext, "application/octet-stream")
    st.download_button(
        f"📥 تنزيل المستند: {name}",
        data=raw,
        file_name=name,
        mime=mime,
        use_container_width=True,
        key=f"{key_prefix}_dl_{secrets.token_hex(4)}"
    )
    if ext in ("png", "jpg", "jpeg", "webp", "gif", "bmp"):
        st.image(raw, caption=name, use_container_width=True)
    elif ext == "pdf":
        b64 = base64.b64encode(raw).decode("utf-8")
        st.markdown(
            f"<iframe src='data:application/pdf;base64,{b64}' width='100%' height='520' style='border:1px solid #C9A961;border-radius:10px;'></iframe>",
            unsafe_allow_html=True
        )
    elif ext in ("txt", "csv"):
        try:
            txt_preview = raw.decode("utf-8", errors="replace")
            st.text_area("معاينة محتوى الملف", value=txt_preview[:8000], height=220, disabled=True, key=f"{key_prefix}_txt_{secrets.token_hex(4)}")
        except Exception:
            pass

# ===== إعدادات الخطوط العربية للـ PDF =====
FONT_DIR = os.path.join(LOCAL_DATA_DIR, "fonts")
os.makedirs(FONT_DIR, exist_ok=True)
FONT_PATH = os.path.join(FONT_DIR, "Amiri-Regular.ttf")

def ensure_arabic_font():
    if os.path.exists(FONT_PATH) and os.path.getsize(FONT_PATH) > 10000:
        return True
    urls = [
        "https://github.com/google/fonts/raw/main/ofl/amiri/Amiri-Regular.ttf",
        "https://github.com/google/fonts/raw/main/ofl/cairo/Cairo%5Bslnt%2Cwght%5D.ttf",
        "https://github.com/google/fonts/raw/main/ofl/tajawal/Tajawal-Regular.ttf",
    ]
    for u in urls:
        try:
            r = requests.get(u, timeout=15)
            if r.status_code == 200 and len(r.content) > 10000:
                with open(FONT_PATH, "wb") as f:
                    f.write(r.content)
                return True
        except Exception:
            pass
    return False

def ar(text):
    if text is None:
        return ""
    s = str(text)
    if not s.strip():
        return s
    try:
        reshaped = arabic_reshaper.reshape(s)
        return get_display(reshaped)
    except Exception:
        return s

# ===== دوال الثيم والألوان والترقيم الآمن =====
THEMES = {
    "الأزرق الملكي والذهبي": {
        "bg": "#0B132B", "card": "#1C2541", "card2": "#131B2E",
        "gold": "#C9A961", "gold2": "#E5C158", "text": "#F8F9FA", "muted": "#CBD5E1",
        "border": "rgba(201,169,97,0.32)", "header_bg": "linear-gradient(135deg, #0B132B 0%, #1C2541 60%, #273469 100%)",
        "nav_bg": "#111936", "accent": "#3A506B"
    },
    "الأخضر الزمردي والذهبي": {
        "bg": "#071E18", "card": "#10372C", "card2": "#0B2820",
        "gold": "#D4AF37", "gold2": "#F0D06B", "text": "#F4F7F5", "muted": "#C7D5CF",
        "border": "rgba(212,175,55,0.35)", "header_bg": "linear-gradient(135deg, #071E18 0%, #10372C 65%, #195242 100%)",
        "nav_bg": "#0B2920", "accent": "#1E5646"
    },
    "الرمادي الفحمي والذهبي": {
        "bg": "#111215", "card": "#1C1E24", "card2": "#16181D",
        "gold": "#C9A961", "gold2": "#E2C376", "text": "#F3F4F6", "muted": "#C4C9D4",
        "border": "rgba(201,169,97,0.30)", "header_bg": "linear-gradient(135deg, #111215 0%, #1E2129 65%, #2B2F3A 100%)",
        "nav_bg": "#15171C", "accent": "#2E3340"
    },
    "العنابي الملكي والذهبي": {
        "bg": "#1E0B10", "card": "#33131C", "card2": "#260E15",
        "gold": "#D4AF37", "gold2": "#F1D27A", "text": "#FAF4F5", "muted": "#DEC7CD",
        "border": "rgba(212,175,55,0.35)", "header_bg": "linear-gradient(135deg, #1E0B10 0%, #33131C 65%, #4D1C29 100%)",
        "nav_bg": "#240D14", "accent": "#5C2232"
    }
}

def apply_theme(theme_name):
    t = THEMES.get(theme_name, THEMES["الأزرق الملكي والذهبي"])
    st.markdown(f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Amiri:wght@400;700&family=Cairo:wght@400;600;700;800&display=swap');

        .stApp {{
            background-color: {t["bg"]};
            color: {t["text"]};
            font-family: 'Cairo', 'Amiri', sans-serif;
            direction: rtl;
        }}
        .block-container {{
            max-width: 1220px;
            padding-top: 1.2rem;
            padding-bottom: 3rem;
        }}
        h1, h2, h3, h4, label, p, span, div {{
            font-family: 'Cairo', 'Amiri', sans-serif;
        }}
        .main-header {{
            background: {t["header_bg"]};
            border: 1px solid {t["border"]};
            border-bottom: 3px solid {t["gold"]};
            border-radius: 18px;
            padding: 22px 26px;
            margin-bottom: 18px;
            box-shadow: 0 10px 28px rgba(0,0,0,0.32);
            text-align: center;
        }}
        .main-title {{
            color: {t["gold2"]};
            font-size: 32px;
            font-weight: 800;
            margin: 0;
            letter-spacing: 0.2px;
        }}
        .sub-title {{
            color: {t["muted"]};
            font-size: 15px;
            margin-top: 6px;
        }}
        .gold-line {{
            height: 2px;
            background: linear-gradient(90deg, transparent, {t["gold"]}, transparent);
            margin: 12px auto 6px auto;
            width: 65%;
        }}
        .section-card {{
            background: {t["card"]};
            border: 1px solid {t["border"]};
            border-radius: 15px;
            padding: 18px 20px;
            margin-bottom: 16px;
            box-shadow: 0 6px 18px rgba(0,0,0,0.20);
        }}
        .stat-box {{
            background: {t["card2"]};
            border: 1px solid {t["border"]};
            border-top: 3px solid {t["gold"]};
            border-radius: 14px;
            padding: 16px 12px;
            text-align: center;
            box-shadow: 0 4px 14px rgba(0,0,0,0.18);
        }}
        .stat-num {{
            font-size: 28px;
            font-weight: 800;
            color: {t["gold2"]};
            line-height: 1.2;
        }}
        .stat-label {{
            font-size: 13px;
            color: {t["muted"]};
            margin-top: 4px;
            font-weight: 700;
        }}
        .case-card {{
            background: {t["card"]};
            border: 1px solid {t["border"]};
            border-right: 5px solid {t["gold"]};
            border-radius: 14px;
            padding: 16px 20px;
            margin-bottom: 14px;
            box-shadow: 0 5px 16px rgba(0,0,0,0.22);
        }}
        .badge-gold {{
            display: inline-block;
            background: rgba(201,169,97,0.16);
            color: {t["gold2"]};
            border: 1px solid {t["border"]};
            padding: 3px 10px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 700;
            margin-left: 6px;
            margin-bottom: 4px;
        }}
        .badge-info {{
            display: inline-block;
            background: rgba(255,255,255,0.08);
            color: {t["text"]};
            border: 1px solid rgba(255,255,255,0.16);
            padding: 3px 10px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            margin-left: 6px;
            margin-bottom: 4px;
        }}
        .stButton > button {{
            border-radius: 10px !important;
            font-weight: 700 !important;
            border: 1px solid {t["border"]} !important;
            transition: all 0.2s ease !important;
        }}
        .stButton > button:hover {{
            border-color: {t["gold2"]} !important;
            transform: translateY(-1px);
        }}
        div[data-testid="stExpander"] {{
            background: {t["card2"]};
            border: 1px solid {t["border"]};
            border-radius: 12px;
        }}
        .report-paper {{
            background: #FFFFFF;
            color: #111827;
            border-radius: 12px;
            padding: 28px 32px;
            border: 2px solid #C9A961;
            box-shadow: 0 8px 24px rgba(0,0,0,0.25);
            direction: rtl;
            text-align: right;
            font-family: 'Amiri', 'Cairo', serif;
        }}
        .report-paper h2, .report-paper h3, .report-paper h4, .report-paper p, .report-paper div, .report-paper td, .report-paper th {{
            color: #111827;
            font-family: 'Amiri', 'Cairo', serif;
        }}
        .report-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 14px;
            margin-bottom: 18px;
            font-size: 14px;
        }}
        .report-table th {{
            background: #1C2541;
            color: #F5D77F !important;
            border: 1px solid #1C2541;
            padding: 8px 10px;
            text-align: center;
            font-weight: 700;
        }}
        .report-table td {{
            border: 1px solid #CBD5E1;
            padding: 7px 9px;
            text-align: center;
            vertical-align: middle;
            color: #111827 !important;
        }}
        .report-table tr:nth-child(even) td {{
            background: #F8FAFC;
        }}
        .party-highlight {{
            color: #C0392B !important;
            font-weight: 800 !important;
        }}
        .report-selector-strip {{
            background: linear-gradient(135deg, #101b31, #1e2f50);
            border: 2px solid #D4AF37;
            border-radius: 14px;
            padding: 12px 14px;
            margin-bottom: 14px;
        }}
        .report-selector-strip [data-testid="stRadio"] > div {{
            gap: 10px;
            flex-wrap: wrap;
        }}
        .report-selector-strip [data-testid="stRadio"] label {{
            background: #142038 !important;
            border: 1px solid #D4AF37 !important;
            border-radius: 10px !important;
            padding: 6px 14px !important;
            color: #FFD700 !important;
            font-weight: 900 !important;
        }}
        .report-selector-strip [data-testid="stRadio"] label[data-checked="true"],
        .report-selector-strip [data-testid="stRadio"] div[role="radiogroup"] > label:has(input:checked) {{
            background: #D4AF37 !important;
            border-color: #FFFFFF !important;
            color: #101b31 !important;
        }}
        .report-selector-strip [data-testid="stRadio"] div[role="radiogroup"] > label:has(input:checked) * {{
            color: #101b31 !important;
            font-weight: 900 !important;
        }}
        .report-selector-strip [data-testid="stRadio"] label * {{
            color: #FFD700 !important;
            font-weight: 900 !important;
        }}
    </style>
    """, unsafe_allow_html=True)

def safe_int(val, default=0):
    try:
        if val is None or str(val).strip() == "":
            return default
        return int(float(str(val).strip()))
    except Exception:
        return default

def normalize_case_numbers(c):
    """الحفاظ على رقم الدعوى والحصر كما هما دون إعادة ترقيم أو تغيير."""
    c["رقم_الدعوى"] = str(c.get("رقم_الدعوى", "") or "").strip()
    c["رقم_الحصر"] = str(c.get("رقم_الحصر", "") or "").strip()
    return c

def case_sort_key(c):
    """ترتيب القضايا بدون تعديل أرقام الدعوى أو الحصر."""
    d_num = safe_int(c.get("رقم_الدعوى"), 10**9)
    h_num = safe_int(c.get("رقم_الحصر"), 10**9)
    return (d_num, str(c.get("رقم_الدعوى", "")), h_num, str(c.get("رقم_الحصر", "")), str(c.get("id", "")))

def get_sorted_cases(cases):
    return sorted([normalize_case_numbers(dict(x)) for x in cases], key=case_sort_key)

def _format_party_display(party_val, html_mode=False):
    s = str(party_val or "").strip()
    if not s:
        return ""
    if s.startswith("طرف"):
        formatted = s
    else:
        formatted = f"طرف الأستاذ / {s} المحامي"
    if html_mode:
        return f"<span class='party-highlight'>{html.escape(formatted)}</span>"
    return formatted

def _strip_party_prefix(party_val):
    s = str(party_val or "").strip()
    s = re.sub(r"^طرف\s+الأستاذ\s*/\s*", "", s)
    s = re.sub(r"\s+المحامي$", "", s)
    return s.strip()

def format_case_title(c, html_mode=False):
    d_no = str(c.get("رقم_الدعوى", "") or "").strip()
    year = str(c.get("سنة_الدعوى", "") or "").strip()
    court = str(c.get("المحكمة", "") or "").strip()
    ctype = str(c.get("نوع_الدعوى", "") or "").strip()
    party = str(c.get("الطرف", "") or "").strip()

    parts = []
    if d_no and year:
        parts.append(f"الدعوى رقم {d_no} لسنة {year} ق")
    elif d_no:
        parts.append(f"الدعوى رقم {d_no}")
    elif year:
        parts.append(f"لسنة {year} ق")

    if ctype:
        parts.append(ctype)
    if court:
        parts.append(court)

    base = " - ".join(parts) if parts else "دعوى بدون رقم"
    if party:
        return f"{base} ({_format_party_display(party, html_mode=html_mode)})"
    return base

def get_latest_session_info(c):
    sessions = c.get("جلسات", [])
    if sessions:
        last_s = sessions[-1]
        return (
            fmt_date_ar(last_s.get("تاريخ", c.get("تاريخ_جلسة", "")), ""),
            last_s.get("قرار", c.get("قرار_الجلسة", "")),
            last_s.get("ملاحظات", c.get("ملاحظات", "")),
        )
    return (
        fmt_date_ar(c.get("تاريخ_جلسة", ""), ""),
        c.get("قرار_الجلسة", ""),
        c.get("ملاحظات", ""),
    )

# ===== دوال التقارير (HTML / Excel / Word / PDF) =====
def build_report_data(cases, report_type, settings, filters=None):
    filters = filters or {}
    sorted_cases = get_sorted_cases(cases)
    filtered = []
    for c in sorted_cases:
        if filters.get("court") and filters["court"] != "الكل" and c.get("المحكمة") != filters["court"]:
            continue
        if filters.get("ctype") and filters["ctype"] != "الكل" and c.get("نوع_الدعوى") != filters["ctype"]:
            continue
        if filters.get("circle") and filters["circle"] != "الكل" and c.get("الدائرة") != filters["circle"]:
            continue
        if filters.get("party") and filters["party"] != "الكل" and c.get("الطرف") != filters["party"]:
            continue
        if filters.get("year") and filters["year"] != "الكل" and str(c.get("سنة_الدعوى", "")) != str(filters["year"]):
            continue
        if filters.get("status") and filters["status"] != "الكل" and c.get("حالة_الدعوى") != filters["status"]:
            continue
        if filters.get("date_from") or filters.get("date_to"):
            d_case = parse_date_safe(c.get("تاريخ_جلسة"))
            d_from = parse_date_safe(filters.get("date_from"))
            d_to = parse_date_safe(filters.get("date_to"))
            if d_case is None:
                continue
            if d_from and d_case < d_from:
                continue
            if d_to and d_case > d_to:
                continue
        if filters.get("search"):
            q = filters["search"].strip().lower()
            blob = " ".join([
                str(c.get("رقم_الدعوى", "")),
                str(c.get("سنة_الدعوى", "")),
                str(c.get("رقم_الحصر", "")),
                str(c.get("المدعي", "")),
                str(c.get("المدعى_عليه", "")),
                str(c.get("الطرف", "")),
                str(c.get("المحكمة", "")),
                str(c.get("الدائرة", "")),
                str(c.get("نوع_الدعوى", "")),
                str(c.get("قرار_الجلسة", "")),
                str(c.get("الحكم", "")),
                str(c.get("ملاحظات", "")),
            ]).lower()
            if q not in blob:
                continue
        filtered.append(c)

    if report_type == "بيان الأحكام":
        judgments = []
        for c in filtered:
            has_j = bool(str(c.get("الحكم", "")).strip()) or (c.get("حالة_الدعوى") == "محكوم فيها") or bool(c.get("تاريخ_الحكم"))
            if has_j:
                judgments.append(c)
        filtered = judgments

    return filtered

def _get_signers(settings):
    member_name = settings.get("member_name", "") or ""
    manager_name = settings.get("manager_name", "") or ""
    member_title = settings.get("member_title", "العضو القانوني") or "العضو القانوني"
    manager_title = settings.get("manager_title", "مدير إدارة القضايا") or "مدير إدارة القضايا"
    return member_title, member_name, manager_title, manager_name

def render_html_report(cases, report_type, settings, filters=None):
    rows_cases = build_report_data(cases, report_type, settings, filters)
    today_str = fmt_date_ar(datetime.now().date())
    member_title, member_name, manager_title, manager_name = _get_signers(settings)

    filter_parts = []
    if filters:
        if filters.get("court") and filters["court"] != "الكل":
            filter_parts.append(f"المحكمة: {filters['court']}")
        if filters.get("ctype") and filters["ctype"] != "الكل":
            filter_parts.append(f"النوع: {filters['ctype']}")
        if filters.get("circle") and filters["circle"] != "الكل":
            filter_parts.append(f"الدائرة: {filters['circle']}")
        if filters.get("party") and filters["party"] != "الكل":
            filter_parts.append(f"الطرف: {_format_party_display(filters['party'], html_mode=False)}")
        if filters.get("year") and filters["year"] != "الكل":
            filter_parts.append(f"السنة: {filters['year']}")
        if filters.get("date_from") or filters.get("date_to"):
            filter_parts.append(f"الفترة: {fmt_date_ar(filters.get('date_from'), 'البداية')} إلى {fmt_date_ar(filters.get('date_to'), 'النهاية')}")
    filter_summary = " | ".join(filter_parts) if filter_parts else "جميع الدعاوى المسجلة"

    html_parts = [
        "<div class='report-paper'>",
        f"<div style='text-align:center;border-bottom:2px solid #C9A961;padding-bottom:12px;margin-bottom:16px;'>",
        f"<div style='font-size:24px;font-weight:700;color:#1C2541;'>{html.escape(report_type)}</div>",
        f"<div style='font-size:14px;color:#475569;margin-top:4px;'>{html.escape(settings.get('app_title', 'إدارة القضايا'))} — {html.escape(filter_summary)}</div>",
        f"<div style='font-size:13px;color:#64748B;margin-top:3px;'>تحريرًا في: {html.escape(today_str)} | إجمالي السجلات: {len(rows_cases)}</div>",
        "</div>",
    ]

    if report_type == "بيان الدعاوى":
        html_parts.append("""
        <table class='report-table'>
            <thead>
                <tr>
                    <th style='width:6%;'>م</th>
                    <th style='width:12%;'>رقم الحصر</th>
                    <th style='width:15%;'>الدعوى والسنة</th>
                    <th style='width:17%;'>الخصوم</th>
                    <th style='width:14%;'>المحكمة والدائرة</th>
                    <th style='width:12%;'>تاريخ الجلسة</th>
                    <th style='width:14%;'>القرار</th>
                    <th style='width:10%;'>ملاحظات</th>
                </tr>
            </thead>
            <tbody>
        """)
        for idx, c in enumerate(rows_cases, 1):
            s_date, s_dec, s_note = get_latest_session_info(c)
            d_str = f"{c.get('رقم_الدعوى','')} لسنة {c.get('سنة_الدعوى','')}".strip()
            court_str = " - ".join([x for x in [c.get("المحكمة", ""), c.get("الدائرة", "")] if x])
            opponents = f"{c.get('المدعي','')} ضد {c.get('المدعى_عليه','')}".strip(" ضد")
            party_html = _format_party_display(c.get("الطرف", ""), html_mode=True)
            opp_html = html.escape(opponents) + (f"<br>{party_html}" if party_html else "")
            row_num_display = str(c.get("رقم_الدعوى") or c.get("رقم_الحصر") or idx)
            html_parts.append(
                f"<tr>"
                f"<td>{html.escape(row_num_display)}</td>"
                f"<td>{html.escape(str(c.get('رقم_الحصر','')))}</td>"
                f"<td>{html.escape(d_str)}</td>"
                f"<td>{opp_html}</td>"
                f"<td>{html.escape(court_str)}</td>"
                f"<td>{html.escape(fmt_date_ar(s_date))}</td>"
                f"<td>{html.escape(str(s_dec or ''))}</td>"
                f"<td>{html.escape(str(s_note or ''))}</td>"
                f"</tr>"
            )
        html_parts.append("</tbody></table>")

    elif report_type == "بيان الأحكام":
        html_parts.append("""
        <table class='report-table'>
            <thead>
                <tr>
                    <th style='width:6%;'>م</th>
                    <th style='width:16%;'>الدعوى والسنة</th>
                    <th style='width:18%;'>الخصوم</th>
                    <th style='width:15%;'>المحكمة والدائرة</th>
                    <th style='width:12%;'>تاريخ الحكم</th>
                    <th style='width:23%;'>منطوق الحكم</th>
                    <th style='width:10%;'>ملاحظات</th>
                </tr>
            </thead>
            <tbody>
        """)
        for idx, c in enumerate(rows_cases, 1):
            d_str = f"{c.get('رقم_الدعوى','')} لسنة {c.get('سنة_الدعوى','')}".strip()
            court_str = " - ".join([x for x in [c.get("المحكمة", ""), c.get("الدائرة", "")] if x])
            opponents = f"{c.get('المدعي','')} ضد {c.get('المدعى_عليه','')}".strip(" ضد")
            party_html = _format_party_display(c.get("الطرف", ""), html_mode=True)
            opp_html = html.escape(opponents) + (f"<br>{party_html}" if party_html else "")
            row_num_display = str(c.get("رقم_الدعوى") or c.get("رقم_الحصر") or idx)
            html_parts.append(
                f"<tr>"
                f"<td>{html.escape(row_num_display)}</td>"
                f"<td>{html.escape(d_str)}</td>"
                f"<td>{opp_html}</td>"
                f"<td>{html.escape(court_str)}</td>"
                f"<td>{html.escape(fmt_date_ar(c.get('تاريخ_الحكم','')))}</td>"
                f"<td>{html.escape(str(c.get('الحكم','')))}</td>"
                f"<td>{html.escape(str(c.get('ملاحظات','')))}</td>"
                f"</tr>"
            )
        html_parts.append("</tbody></table>")

    elif report_type == "الإحصائيات":
        total_c = len(rows_cases)
        by_status = {}
        by_court = {}
        by_type = {}
        by_circle = {}
        for c in rows_cases:
            st_val = c.get("حالة_الدعوى") or "متداولة"
            ct_val = c.get("المحكمة") or "غير محدد"
            tp_val = c.get("نوع_الدعوى") or "غير محدد"
            cr_val = c.get("الدائرة") or "غير محددة"
            by_status[st_val] = by_status.get(st_val, 0) + 1
            by_court[ct_val] = by_court.get(ct_val, 0) + 1
            by_type[tp_val] = by_type.get(tp_val, 0) + 1
            by_circle[cr_val] = by_circle.get(cr_val, 0) + 1

        html_parts.append(f"<p style='font-size:16px;font-weight:700;'>إجمالي الدعاوى المشمولة بالتقرير: {total_c}</p>")

        html_parts.append("<h4>أولًا: التوزيع حسب حالة الدعوى</h4><table class='report-table'><thead><tr><th>الحالة</th><th>العدد</th></tr></thead><tbody>")
        for k, v in by_status.items():
            html_parts.append(f"<tr><td>{html.escape(str(k))}</td><td>{v}</td></tr>")
        html_parts.append("</tbody></table>")

        html_parts.append("<h4>ثانيًا: التوزيع حسب المحكمة</h4><table class='report-table'><thead><tr><th>المحكمة</th><th>العدد</th></tr></thead><tbody>")
        for k, v in by_court.items():
            html_parts.append(f"<tr><td>{html.escape(str(k))}</td><td>{v}</td></tr>")
        html_parts.append("</tbody></table>")

        html_parts.append("<h4>ثالثًا: التوزيع حسب نوع الدعوى</h4><table class='report-table'><thead><tr><th>نوع الدعوى</th><th>العدد</th></tr></thead><tbody>")
        for k, v in by_type.items():
            html_parts.append(f"<tr><td>{html.escape(str(k))}</td><td>{v}</td></tr>")
        html_parts.append("</tbody></table>")

        html_parts.append("<h4>رابعًا: التوزيع حسب الدائرة</h4><table class='report-table'><thead><tr><th>الدائرة</th><th>العدد</th></tr></thead><tbody>")
        for k, v in by_circle.items():
            html_parts.append(f"<tr><td>{html.escape(str(k))}</td><td>{v}</td></tr>")
        html_parts.append("</tbody></table>")

    elif report_type == "بيان عددي":
        groups = {}
        for c in rows_cases:
            key = (
                c.get("المحكمة") or "غير محدد",
                c.get("نوع_الدعوى") or "غير محدد",
                c.get("الدائرة") or "غير محددة",
            )
            groups[key] = groups.get(key, 0) + 1

        html_parts.append("""
        <table class='report-table'>
            <thead>
                <tr>
                    <th style='width:8%;'>م</th>
                    <th style='width:28%;'>المحكمة</th>
                    <th style='width:26%;'>نوع الدعوى</th>
                    <th style='width:24%;'>الدائرة</th>
                    <th style='width:14%;'>العدد</th>
                </tr>
            </thead>
            <tbody>
        """)
        for idx, ((ct, tp, cr), cnt) in enumerate(groups.items(), 1):
            html_parts.append(
                f"<tr><td>{idx}</td><td>{html.escape(str(ct))}</td><td>{html.escape(str(tp))}</td><td>{html.escape(str(cr))}</td><td><strong>{cnt}</strong></td></tr>"
            )
        html_parts.append(
            f"<tr style='background:#FEF3C7;font-weight:700;'><td colspan='4'>الإجمالي الكلي</td><td>{len(rows_cases)}</td></tr>"
        )
        html_parts.append("</tbody></table>")

    html_parts.append(f"""
    <div style='display:flex;justify-content:space-between;align-items:flex-start;margin-top:34px;padding-top:16px;border-top:1px dashed #94A3B8;'>
        <div style='width:42%;text-align:center;'>
            <div style='font-weight:700;font-size:15px;'>{html.escape(member_title)}</div>
            <div style='margin-top:8px;font-size:14px;font-weight:700;'>{html.escape(member_name or '........................')}</div>
        </div>
        <div style='width:42%;text-align:center;'>
            <div style='font-weight:700;font-size:15px;'>{html.escape(manager_title)}</div>
            <div style='margin-top:8px;font-size:14px;font-weight:700;'>{html.escape(manager_name or '........................')}</div>
        </div>
    </div>
    """)
    html_parts.append("</div>")
    return "\n".join(html_parts)

def _set_cell_background(cell, hex_color):
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), hex_color.lstrip("#"))
    cell._tc.get_or_add_tcPr().append(shading)

def _set_table_borders(table, color="CBD5E1", sz="4"):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for b_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        b = OxmlElement(f"w:{b_name}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), sz)
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), color)
        borders.append(b)
    tblPr.append(borders)

def _set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)

def _add_formatted_cell_text(cell, plain_text, party_text="", font_size=10.5, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(plain_text or "")
    r.font.name = "Arial"
    r.font.size = Pt(font_size)
    r.bold = bold
    if party_text:
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(1)
        p2.paragraph_format.space_after = Pt(1)
        r2 = p2.add_run(_format_party_display(party_text, html_mode=False))
        r2.font.name = "Arial"
        r2.font.size = Pt(font_size)
        r2.bold = True

def to_word(cases, report_type, settings, filters=None):
    rows_cases = build_report_data(cases, report_type, settings, filters)
    today_str = fmt_date_ar(datetime.now().date())
    member_title, member_name, manager_title, manager_name = _get_signers(settings)

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title_p.add_run(f"{report_type} — {settings.get('app_title', 'إدارة القضايا')}")
    tr.font.name = "Arial"
    tr.font.size = Pt(17)
    tr.bold = True

    meta_p = doc.add_paragraph()
    meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mr = meta_p.add_run(f"تحريرًا في: {today_str}   |   عدد السجلات: {len(rows_cases)}")
    mr.font.name = "Arial"
    mr.font.size = Pt(10.5)

    if report_type == "بيان الدعاوى":
        headers = ["م", "رقم الحصر", "الدعوى والسنة", "الخصوم", "المحكمة والدائرة", "تاريخ الجلسة", "القرار", "ملاحظات"]
        tbl = doc.add_table(rows=1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_borders(tbl)
        for j, h in enumerate(headers):
            cell = tbl.rows[0].cells[j]
            _set_cell_background(cell, "1C2541")
            _set_cell_margins(cell, 90, 90, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(h)
            run.font.name = "Arial"
            run.font.size = Pt(10)
            run.bold = True

        for idx, c in enumerate(rows_cases, 1):
            s_date, s_dec, s_note = get_latest_session_info(c)
            d_str = f"{c.get('رقم_الدعوى','')} لسنة {c.get('سنة_الدعوى','')}".strip()
            court_str = " - ".join([x for x in [c.get("المحكمة", ""), c.get("الدائرة", "")] if x])
            opponents = f"{c.get('المدعي','')} ضد {c.get('المدعى_عليه','')}".strip(" ضد")
            row_num_display = str(c.get("رقم_الدعوى") or c.get("رقم_الحصر") or idx)
            vals = [row_num_display, str(c.get("رقم_الحصر", "")), d_str, opponents, court_str, fmt_date_ar(s_date), str(s_dec or ""), str(s_note or "")]
            row_cells = tbl.add_row().cells
            for j, val in enumerate(vals):
                _set_cell_margins(row_cells[j], 70, 70, 70, 70)
                if idx % 2 == 0:
                    _set_cell_background(row_cells[j], "F8FAFC")
                if j == 3:
                    _add_formatted_cell_text(row_cells[j], val, party_text=c.get("الطرف", ""), font_size=9.5)
                else:
                    _add_formatted_cell_text(row_cells[j], val, font_size=9.5)

    elif report_type == "بيان الأحكام":
        headers = ["م", "الدعوى والسنة", "الخصوم", "المحكمة والدائرة", "تاريخ الحكم", "منطوق الحكم", "ملاحظات"]
        tbl = doc.add_table(rows=1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_borders(tbl)
        for j, h in enumerate(headers):
            cell = tbl.rows[0].cells[j]
            _set_cell_background(cell, "1C2541")
            _set_cell_margins(cell, 90, 90, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(h)
            run.font.name = "Arial"
            run.font.size = Pt(10)
            run.bold = True

        for idx, c in enumerate(rows_cases, 1):
            d_str = f"{c.get('رقم_الدعوى','')} لسنة {c.get('سنة_الدعوى','')}".strip()
            court_str = " - ".join([x for x in [c.get("المحكمة", ""), c.get("الدائرة", "")] if x])
            opponents = f"{c.get('المدعي','')} ضد {c.get('المدعى_عليه','')}".strip(" ضد")
            row_num_display = str(c.get("رقم_الدعوى") or c.get("رقم_الحصر") or idx)
            vals = [row_num_display, d_str, opponents, court_str, fmt_date_ar(c.get("تاريخ_الحكم", "")), str(c.get("الحكم", "")), str(c.get("ملاحظات", ""))]
            row_cells = tbl.add_row().cells
            for j, val in enumerate(vals):
                _set_cell_margins(row_cells[j], 70, 70, 70, 70)
                if idx % 2 == 0:
                    _set_cell_background(row_cells[j], "F8FAFC")
                if j == 2:
                    _add_formatted_cell_text(row_cells[j], val, party_text=c.get("الطرف", ""), font_size=9.5)
                else:
                    _add_formatted_cell_text(row_cells[j], val, font_size=9.5)

    elif report_type == "بيان عددي":
        groups = {}
        for c in rows_cases:
            key = (
                c.get("المحكمة") or "غير محدد",
                c.get("نوع_الدعوى") or "غير محدد",
                c.get("الدائرة") or "غير محددة",
            )
            groups[key] = groups.get(key, 0) + 1
        headers = ["م", "المحكمة", "نوع الدعوى", "الدائرة", "العدد"]
        tbl = doc.add_table(rows=1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_borders(tbl)
        for j, h in enumerate(headers):
            cell = tbl.rows[0].cells[j]
            _set_cell_background(cell, "1C2541")
            _set_cell_margins(cell, 90, 90, 80, 80)
            _add_formatted_cell_text(cell, h, font_size=10, bold=True)
        for idx, ((ct, tp, cr), cnt) in enumerate(groups.items(), 1):
            row_cells = tbl.add_row().cells
            for j, val in enumerate([str(idx), str(ct), str(tp), str(cr), str(cnt)]):
                _set_cell_margins(row_cells[j], 70, 70, 70, 70)
                _add_formatted_cell_text(row_cells[j], val, font_size=10, bold=(j == 4))
        tot_cells = tbl.add_row().cells
        _add_formatted_cell_text(tot_cells[0], "الإجمالي الكلي", font_size=10.5, bold=True)
        tot_cells[0].merge(tot_cells[3])
        _add_formatted_cell_text(tot_cells[4], str(len(rows_cases)), font_size=10.5, bold=True)

    else:
        by_status = {}
        by_court = {}
        by_type = {}
        by_circle = {}
        for c in rows_cases:
            by_status[c.get("حالة_الدعوى") or "متداولة"] = by_status.get(c.get("حالة_الدعوى") or "متداولة", 0) + 1
            by_court[c.get("المحكمة") or "غير محدد"] = by_court.get(c.get("المحكمة") or "غير محدد", 0) + 1
            by_type[c.get("نوع_الدعوى") or "غير محدد"] = by_type.get(c.get("نوع_الدعوى") or "غير محدد", 0) + 1
            by_circle[c.get("الدائرة") or "غير محددة"] = by_circle.get(c.get("الدائرة") or "غير محددة", 0) + 1

        for sec_title, d_map, col1 in [
            ("التوزيع حسب حالة الدعوى", by_status, "الحالة"),
            ("التوزيع حسب المحكمة", by_court, "المحكمة"),
            ("التوزيع حسب نوع الدعوى", by_type, "نوع الدعوى"),
            ("التوزيع حسب الدائرة", by_circle, "الدائرة"),
        ]:
            sp = doc.add_paragraph()
            sp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            sr = sp.add_run(sec_title)
            sr.font.name = "Arial"
            sr.font.size = Pt(12)
            sr.bold = True

            tbl = doc.add_table(rows=1, cols=2)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            _set_table_borders(tbl)
            for j, h in enumerate([col1, "العدد"]):
                _set_cell_background(tbl.rows[0].cells[j], "1C2541")
                _add_formatted_cell_text(tbl.rows[0].cells[j], h, font_size=10, bold=True)
            for k, v in d_map.items():
                rc = tbl.add_row().cells
                _add_formatted_cell_text(rc[0], str(k), font_size=10)
                _add_formatted_cell_text(rc[1], str(v), font_size=10, bold=True)

    doc.add_paragraph().paragraph_format.space_before = Pt(14)
    sig_tbl = doc.add_table(rows=1, cols=3)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_tbl.autofit = False
    col_widths = [Inches(2.8), Inches(1.7), Inches(2.8)]
    for idx_c, w in enumerate(col_widths):
        sig_tbl.columns[idx_c].width = w
        sig_tbl.rows[0].cells[idx_c].width = w

    left_cell = sig_tbl.rows[0].cells[0]
    gap_cell = sig_tbl.rows[0].cells[1]
    right_cell = sig_tbl.rows[0].cells[2]

    gap_cell.text = ""
    _set_cell_margins(left_cell, 80, 80, 120, 120)
    _set_cell_margins(right_cell, 80, 80, 120, 120)

    p_left = left_cell.paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rl1 = p_left.add_run(member_title)
    rl1.font.name = "Arial"
    rl1.font.size = Pt(11.5)
    rl1.bold = True
    p_left_name = left_cell.add_paragraph()
    p_left_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_left_name.paragraph_format.space_before = Pt(6)
    rl2 = p_left_name.add_run(member_name or "........................")
    rl2.font.name = "Arial"
    rl2.font.size = Pt(11)

    p_right = right_cell.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr1 = p_right.add_run(manager_title)
    rr1.font.name = "Arial"
    rr1.font.size = Pt(11.5)
    rr1.bold = True
    p_right_name = right_cell.add_paragraph()
    p_right_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_right_name.paragraph_format.space_before = Pt(6)
    rr2 = p_right_name.add_run(manager_name or "........................")
    rr2.font.name = "Arial"
    rr2.font.size = Pt(11)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

def to_excel(cases, report_type, settings, filters=None):
    rows_cases = build_report_data(cases, report_type, settings, filters)
    today_str = fmt_date_ar(datetime.now().date())
    member_title, member_name, manager_title, manager_name = _get_signers(settings)

    records = []
    if report_type == "بيان الدعاوى":
        for idx, c in enumerate(rows_cases, 1):
            s_date, s_dec, s_note = get_latest_session_info(c)
            opp = f"{c.get('المدعي','')} ضد {c.get('المدعى_عليه','')}".strip(" ضد")
            if c.get("الطرف"):
                opp = f"{opp} ({_format_party_display(c.get('الطرف', ''), html_mode=False)})".strip()
            row_num_display = str(c.get("رقم_الدعوى") or c.get("رقم_الحصر") or idx)
            records.append({
                "م": row_num_display,
                "رقم الحصر": c.get("رقم_الحصر", ""),
                "رقم الدعوى": c.get("رقم_الدعوى", ""),
                "السنة": c.get("سنة_الدعوى", ""),
                "الخصوم": opp,
                "المحكمة": c.get("المحكمة", ""),
                "الدائرة": c.get("الدائرة", ""),
                "تاريخ الجلسة": fmt_date_ar(s_date),
                "القرار": s_dec,
                "ملاحظات": s_note,
            })
    elif report_type == "بيان الأحكام":
        for idx, c in enumerate(rows_cases, 1):
            opp = f"{c.get('المدعي','')} ضد {c.get('المدعى_عليه','')}".strip(" ضد")
            if c.get("الطرف"):
                opp = f"{opp} ({_format_party_display(c.get('الطرف', ''), html_mode=False)})".strip()
            row_num_display = str(c.get("رقم_الدعوى") or c.get("رقم_الحصر") or idx)
            records.append({
                "م": row_num_display,
                "رقم الدعوى": c.get("رقم_الدعوى", ""),
                "السنة": c.get("سنة_الدعوى", ""),
                "الخصوم": opp,
                "المحكمة": c.get("المحكمة", ""),
                "الدائرة": c.get("الدائرة", ""),
                "تاريخ الحكم": fmt_date_ar(c.get("تاريخ_الحكم", "")),
                "منطوق الحكم": c.get("الحكم", ""),
                "ملاحظات": c.get("ملاحظات", ""),
            })
    elif report_type == "بيان عددي":
        groups = {}
        for c in rows_cases:
            key = (
                c.get("المحكمة") or "غير محدد",
                c.get("نوع_الدعوى") or "غير محدد",
                c.get("الدائرة") or "غير محددة",
            )
            groups[key] = groups.get(key, 0) + 1
        for idx, ((ct, tp, cr), cnt) in enumerate(groups.items(), 1):
            records.append({"م": idx, "المحكمة": ct, "نوع الدعوى": tp, "الدائرة": cr, "العدد": cnt})
        records.append({"م": "الإجمالي", "المحكمة": "", "نوع الدعوى": "", "الدائرة": "", "العدد": len(rows_cases)})
    else:
        for c in rows_cases:
            opp = f"{c.get('المدعي','')} ضد {c.get('المدعى_عليه','')}".strip(" ضد")
            if c.get("الطرف"):
                opp = f"{opp} ({_format_party_display(c.get('الطرف', ''), html_mode=False)})".strip()
            records.append({
                "رقم الدعوى": c.get("رقم_الدعوى", ""),
                "السنة": c.get("سنة_الدعوى", ""),
                "نوع الدعوى": c.get("نوع_الدعوى", ""),
                "المحكمة": c.get("المحكمة", ""),
                "الدائرة": c.get("الدائرة", ""),
                "الحالة": c.get("حالة_الدعوى", ""),
                "الخصوم": opp,
                "تاريخ الجلسة": fmt_date_ar(c.get("تاريخ_جلسة", "")),
            })

    df = pd.DataFrame(records) if records else pd.DataFrame([{"ملاحظة": "لا توجد بيانات مطابقة"}])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="التقرير", startrow=3)
        ws = writer.sheets["التقرير"]
        ws.sheet_view.rightToLeft = True
        ws["A1"] = f"{report_type} — {settings.get('app_title', 'إدارة القضايا')}"
        ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1C2541")
        ws["A2"] = f"تحريرًا في: {today_str} | العدد: {len(rows_cases)}"
        ws["A2"].font = Font(name="Arial", size=10, bold=True, color="475569")

        header_fill = PatternFill(start_color="1C2541", end_color="1C2541", fill_type="solid")
        header_font = Font(name="Arial", size=10, bold=True, color="F5D77F")
        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=4, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for r_idx in range(5, 5 + len(df)):
            for c_idx in range(1, len(df.columns) + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = Font(name="Arial", size=10, color="111827")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        footer_row = 6 + len(df)
        ws.cell(row=footer_row, column=1, value=f"{member_title}: {member_name or '................'}").font = Font(name="Arial", size=10, bold=True)
        ws.cell(row=footer_row, column=max(2, len(df.columns)), value=f"{manager_title}: {manager_name or '................'}").font = Font(name="Arial", size=10, bold=True)

        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=12)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 5, 14), 42)

    buf.seek(0)
    return buf.getvalue()

def _pdf_write_table(pdf, headers, rows_data, col_widths, font_name, font_size=9.5):
    line_h = 5.2
    pdf.set_font(font_name, "", font_size)
    pdf.set_fill_color(28, 37, 65)
    pdf.set_text_color(245, 215, 127)
    pdf.set_draw_color(28, 37, 65)

    for j, h in enumerate(headers):
        pdf.cell(col_widths[j], 8.5, ar(h), border=1, align="C", fill=True)
    pdf.ln(8.5)

    pdf.set_text_color(17, 24, 39)
    pdf.set_draw_color(203, 213, 225)

    for r_idx, row_vals in enumerate(rows_data):
        cell_lines = []
        max_lines = 1
        for j, val in enumerate(row_vals):
            txt = str(val or "")
            w = max(col_widths[j] - 2.0, 8.0)
            wrapped = pdf.multi_cell(w, line_h, ar(txt), split_only=True) if txt.strip() else [""]
            lines = wrapped if wrapped else [""]
            cell_lines.append(lines)
            if len(lines) > max_lines:
                max_lines = len(lines)

        row_h = max(max_lines * line_h + 2.2, 7.5)
        if pdf.get_y() + row_h > pdf.page_break_trigger:
            pdf.add_page()
            pdf.set_font(font_name, "", font_size)
            pdf.set_fill_color(28, 37, 65)
            pdf.set_text_color(245, 215, 127)
            pdf.set_draw_color(28, 37, 65)
            for j, h in enumerate(headers):
                pdf.cell(col_widths[j], 8.5, ar(h), border=1, align="C", fill=True)
            pdf.ln(8.5)
            pdf.set_text_color(17, 24, 39)
            pdf.set_draw_color(203, 213, 225)

        x_start = pdf.get_x()
        y_start = pdf.get_y()
        if r_idx % 2 == 1:
            pdf.set_fill_color(248, 250, 252)
            pdf.rect(x_start, y_start, sum(col_widths), row_h, style="F")

        cur_x = x_start
        for j, lines in enumerate(cell_lines):
            w = col_widths[j]
            pdf.rect(cur_x, y_start, w, row_h)
            block_h = len(lines) * line_h
            y_text = y_start + (row_h - block_h) / 2.0
            for l_i, line_txt in enumerate(lines):
                pdf.set_xy(cur_x + 1.0, y_text + l_i * line_h)
                pdf.cell(w - 2.0, line_h, line_txt, align="C")
            cur_x += w

        pdf.set_xy(x_start, y_start + row_h)

def to_pdf(cases, report_type, settings, filters=None):
    rows_cases = build_report_data(cases, report_type, settings, filters)
    today_str = fmt_date_ar(datetime.now().date())
    member_title, member_name, manager_title, manager_name = _get_signers(settings)

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    has_ar = ensure_arabic_font()
    font_name = "ArabicFont" if has_ar else "Helvetica"
    if has_ar:
        pdf.add_font("ArabicFont", "", FONT_PATH, uni=True)

    pdf.set_font(font_name, "", 15)
    pdf.set_text_color(28, 37, 65)
    pdf.cell(0, 9, ar(f"{report_type} — {settings.get('app_title', 'إدارة القضايا')}"), ln=True, align="C")

    pdf.set_font(font_name, "", 10)
    pdf.set_text_color(71, 85, 105)
    pdf.cell(0, 6.5, ar(f"تحريرًا في: {today_str}   |   عدد السجلات: {len(rows_cases)}"), ln=True, align="C")
    pdf.ln(3)

    if report_type == "بيان الدعاوى":
        headers = ["م", "رقم الحصر", "الدعوى والسنة", "الخصوم", "المحكمة والدائرة", "تاريخ الجلسة", "القرار", "ملاحظات"]
        col_widths = [14.0, 22.0, 36.0, 56.0, 42.0, 25.0, 46.0, 28.0]
        rows_data = []
        for idx, c in enumerate(rows_cases, 1):
            s_date, s_dec, s_note = get_latest_session_info(c)
            d_str = f"{c.get('رقم_الدعوى','')} لسنة {c.get('سنة_الدعوى','')}".strip()
            court_str = " - ".join([x for x in [c.get("المحكمة", ""), c.get("الدائرة", "")] if x])
            opp = f"{c.get('المدعي','')} ضد {c.get('المدعى_عليه','')}".strip(" ضد")
            if c.get("الطرف"):
                opp = f"{opp}\n{_format_party_display(c.get('الطرف', ''), html_mode=False)}".strip()
            row_num_display = str(c.get("رقم_الدعوى") or c.get("رقم_الحصر") or idx)
            rows_data.append([row_num_display, str(c.get("رقم_الحصر", "")), d_str, opp, court_str, fmt_date_ar(s_date), str(s_dec or ""), str(s_note or "")])
        _pdf_write_table(pdf, headers, rows_data, col_widths, font_name, font_size=9.0)

    elif report_type == "بيان الأحكام":
        headers = ["م", "الدعوى والسنة", "الخصوم", "المحكمة والدائرة", "تاريخ الحكم", "منطوق الحكم", "ملاحظات"]
        col_widths = [15.0, 38.0, 58.0, 44.0, 26.0, 62.0, 26.0]
        rows_data = []
        for idx, c in enumerate(rows_cases, 1):
            d_str = f"{c.get('رقم_الدعوى','')} لسنة {c.get('سنة_الدعوى','')}".strip()
            court_str = " - ".join([x for x in [c.get("المحكمة", ""), c.get("الدائرة", "")] if x])
            opp = f"{c.get('المدعي','')} ضد {c.get('المدعى_عليه','')}".strip(" ضد")
            if c.get("الطرف"):
                opp = f"{opp}\n{_format_party_display(c.get('الطرف', ''), html_mode=False)}".strip()
            row_num_display = str(c.get("رقم_الدعوى") or c.get("رقم_الحصر") or idx)
            rows_data.append([row_num_display, d_str, opp, court_str, fmt_date_ar(c.get("تاريخ_الحكم", "")), str(c.get("الحكم", "")), str(c.get("ملاحظات", ""))])
        _pdf_write_table(pdf, headers, rows_data, col_widths, font_name, font_size=9.0)

    elif report_type == "بيان عددي":
        groups = {}
        for c in rows_cases:
            key = (
                c.get("المحكمة") or "غير محدد",
                c.get("نوع_الدعوى") or "غير محدد",
                c.get("الدائرة") or "غير محددة",
            )
            groups[key] = groups.get(key, 0) + 1
        headers = ["م", "المحكمة", "نوع الدعوى", "الدائرة", "العدد"]
        col_widths = [18.0, 82.0, 74.0, 68.0, 27.0]
        rows_data = []
        for idx, ((ct, tp, cr), cnt) in enumerate(groups.items(), 1):
            rows_data.append([str(idx), str(ct), str(tp), str(cr), str(cnt)])
        rows_data.append(["", ar("الإجمالي الكلي"), "", "", str(len(rows_cases))])
        _pdf_write_table(pdf, headers, rows_data, col_widths, font_name, font_size=9.5)

    else:
        by_status = {}
        by_court = {}
        by_type = {}
        by_circle = {}
        for c in rows_cases:
            by_status[c.get("حالة_الدعوى") or "متداولة"] = by_status.get(c.get("حالة_الدعوى") or "متداولة", 0) + 1
            by_court[c.get("المحكمة") or "غير محدد"] = by_court.get(c.get("المحكمة") or "غير محدد", 0) + 1
            by_type[c.get("نوع_الدعوى") or "غير محدد"] = by_type.get(c.get("نوع_الدعوى") or "غير محدد", 0) + 1
            by_circle[c.get("الدائرة") or "غير محددة"] = by_circle.get(c.get("الدائرة") or "غير محددة", 0) + 1

        for sec_title, d_map, col1 in [
            ("التوزيع حسب حالة الدعوى", by_status, "الحالة"),
            ("التوزيع حسب المحكمة", by_court, "المحكمة"),
            ("التوزيع حسب نوع الدعوى", by_type, "نوع الدعوى"),
            ("التوزيع حسب الدائرة", by_circle, "الدائرة"),
        ]:
            pdf.set_font(font_name, "", 11.5)
            pdf.set_text_color(28, 37, 65)
            pdf.cell(0, 7.5, ar(sec_title), ln=True, align="R")
            r_list = [[str(k), str(v)] for k, v in d_map.items()]
            _pdf_write_table(pdf, [col1, "العدد"], r_list, [120.0, 40.0], font_name, font_size=9.5)
            pdf.ln(3)

    if pdf.get_y() + 26 > pdf.page_break_trigger:
        pdf.add_page()

    pdf.ln(6)
    y_sig = pdf.get_y()
    pdf.set_font(font_name, "", 11)
    pdf.set_text_color(17, 24, 39)

    pdf.set_xy(20, y_sig)
    pdf.cell(105, 6.5, ar(member_title), align="C")
    pdf.set_xy(20, y_sig + 7.5)
    pdf.cell(105, 6.5, ar(member_name or "........................"), align="C")

    pdf.set_xy(172, y_sig)
    pdf.cell(105, 6.5, ar(manager_title), align="C")
    pdf.set_xy(172, y_sig + 7.5)
    pdf.cell(105, 6.5, ar(manager_name or "........................"), align="C")

    out = pdf.output(dest="S")
    if isinstance(out, str):
        return out.encode("latin1")
    return bytes(out)
# ===== مكونات الواجهة (الهيدر، الإحصائيات، عرض القضايا) =====
def render_header(settings):
    title = settings.get("app_title", "إدارة القضايا")
    subtitle = settings.get("subtitle", "منظومة متكاملة لإدارة الدعاوى والقرارات والأحكام والمذكرات")
    st.markdown(f"""
    <div class="main-header">
        <div class="main-title">⚖️ {html.escape(title)}</div>
        <div class="gold-line"></div>
        <div class="sub-title">{html.escape(subtitle)}</div>
    </div>
    """, unsafe_allow_html=True)

def render_stats_row(cases):
    total = len(cases)
    active = sum(1 for c in cases if (c.get("حالة_الدعوى") or "متداولة") == "متداولة")
    judged = sum(1 for c in cases if (c.get("حالة_الدعوى") == "محكوم فيها" or bool(str(c.get("الحكم", "")).strip())))
    today_str = fmt_date_ar(datetime.now().date())
    today_sessions = sum(1 for c in cases if fmt_date_ar(c.get("تاريخ_جلسة", "")) == today_str)

    c1, c2, c3, c4 = st.columns(4)
    for col, num, lbl in [
        (c1, total, "إجمالي الدعاوى"),
        (c2, active, "دعاوى متداولة"),
        (c3, judged, "أحكام صادرة"),
        (c4, today_sessions, "جلسات اليوم"),
    ]:
        with col:
            st.markdown(f"""
            <div class="stat-box">
                <div class="stat-num">{num}</div>
                <div class="stat-label">{lbl}</div>
            </div>
            """, unsafe_allow_html=True)

def render_case_card(c, index=0, allow_edit=True, allow_delete=True, data_ref=None):
    cid = c.get("id") or f"case_{index}"
    title_html = format_case_title(c, html_mode=True)
    status = c.get("حالة_الدعوى") or "متداولة"
    party_html = _format_party_display(c.get("الطرف", ""), html_mode=True)
    opponents_plain = f"{c.get('المدعي','')} ضد {c.get('المدعى_عليه','')}".strip(" ضد")
    opponents_html = html.escape(opponents_plain) + (f" — {party_html}" if party_html else "")
    s_date, s_dec, s_note = get_latest_session_info(c)

    st.markdown(f"""
    <div class="case-card">
        <div style="font-size:18px;font-weight:800;color:#F5D77F;margin-bottom:6px;">{title_html}</div>
        <div style="margin-bottom:8px;">
            <span class="badge-gold">الحالة: {html.escape(str(status))}</span>
            {f"<span class='badge-info'>الحصر: {html.escape(str(c.get('رقم_الحصر')))}</span>" if c.get('رقم_الحصر') else ""}
            {f"<span class='badge-info'>الدائرة: {html.escape(str(c.get('الدائرة')))}</span>" if c.get('الدائرة') else ""}
            {f"<span class='badge-info'>جلسة: {html.escape(fmt_date_ar(s_date))}</span>" if s_date else ""}
        </div>
        <div style="font-size:14px;margin-bottom:4px;"><strong>الخصوم:</strong> {opponents_html}</div>
        {f"<div style='font-size:14px;margin-bottom:4px;'><strong>آخر قرار:</strong> {html.escape(str(s_dec))}</div>" if s_dec else ""}
        {f"<div style='font-size:14px;margin-bottom:4px;color:#E5C158;'><strong>الحكم:</strong> {html.escape(str(c.get('الحكم')))}</div>" if c.get('الحكم') else ""}
        {f"<div style='font-size:13px;color:#CBD5E1;'><strong>ملاحظات:</strong> {html.escape(str(s_note or c.get('ملاحظات','')))}</div>" if (s_note or c.get('ملاحظات')) else ""}
    </div>
    """, unsafe_allow_html=True)

    with st.expander(f"📋 تفاصيل الدعوى والجلسات والمستندات ({c.get('رقم_الدعوى','')}/{c.get('سنة_الدعوى','')})", expanded=False):
        st.write(f"**المدعي:** {c.get('المدعي','') or '—'}  |  **المدعى عليه:** {c.get('المدعى_عليه','') or '—'}")
        if c.get("الطرف"):
            st.markdown(f"**الطرف:** {_format_party_display(c.get('الطرف', ''), html_mode=True)}", unsafe_allow_html=True)
        st.write(f"**المحكمة:** {c.get('المحكمة','') or '—'}  |  **الدائرة:** {c.get('الدائرة','') or '—'}  |  **نوع الدعوى:** {c.get('نوع_الدعوى','') or '—'}")
        st.write(f"**رقم الحصر:** {c.get('رقم_الحصر','') or '—'}  |  **تاريخ الحكم:** {fmt_date_ar(c.get('تاريخ_الحكم',''), '—')}")

        sessions = c.get("جلسات") or []
        if sessions:
            st.markdown("**🗓️ سجل الجلسات:**")
            for idx_s, s in enumerate(sessions, 1):
                st.markdown(f"- **جلسة {idx_s}** ({fmt_date_ar(s.get('تاريخ',''), '—')}): {s.get('قرار','') or '—'} {f'— *ملاحظات:* {s.get(\"ملاحظات\")}' if s.get('ملاحظات') else ''}")

        docs = c.get("مستندات") or []
        if docs:
            st.markdown("**📎 المستندات المرفقة:**")
            for d_idx, doc_obj in enumerate(docs):
                _render_doc_viewer(doc_obj, key_prefix=f"case_doc_{cid}_{d_idx}")

        if allow_edit or allow_delete:
            b1, b2 = st.columns(2)
            if allow_edit:
                with b1:
                    if st.button("✏️ تعديل الدعوى", key=f"edit_case_btn_{cid}_{index}", use_container_width=True):
                        st.session_state["editing_case_id"] = c.get("id")
                        st.session_state.page = "إضافة دعوى"
                        st.rerun()
            if allow_delete:
                with b2:
                    if st.button("🗑️ حذف الدعوى", key=f"del_case_btn_{cid}_{index}", use_container_width=True):
                        st.session_state[f"confirm_del_case_{cid}_{index}"] = True
                        st.rerun()

            if st.session_state.get(f"confirm_del_case_{cid}_{index}"):
                st.warning("تأكيد حذف هذه الدعوى نهائيًا؟")
                y, n = st.columns(2)
                with y:
                    if st.button("نعم، احذف", key=f"yes_del_case_{cid}_{index}", type="primary", use_container_width=True):
                        full_data = data_ref if isinstance(data_ref, dict) else load_data()
                        target_case = next((x for x in full_data.get("cases", []) if x.get("id") == c.get("id")), None)
                        if target_case:
                            for doc_obj in target_case.get("مستندات") or []:
                                _delete_doc_obj(doc_obj)
                        full_data["cases"] = [x for x in full_data.get("cases", []) if x.get("id") != c.get("id")]
                        save_data(full_data)
                        st.session_state.pop(f"confirm_del_case_{cid}_{index}", None)
                        st.success("تم حذف الدعوى")
                        st.rerun()
                with n:
                    if st.button("إلغاء", key=f"no_del_case_{cid}_{index}", use_container_width=True):
                        st.session_state.pop(f"confirm_del_case_{cid}_{index}", None)
                        st.rerun()

# ===== الصفحات الرئيسية للمنظومة =====
def home_page(data):
    settings = data["settings"]
    cases = data["cases"]
    render_header(settings)
    render_stats_row(cases)

    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.markdown("### ⚡ وصول سريع")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if st.button("➕ إضافة دعوى جديدة", use_container_width=True, type="primary"):
            st.session_state.pop("editing_case_id", None)
            st.session_state.page = "إضافة دعوى"
            st.rerun()
    with c2:
        if st.button("📅 جلسات اليوم", use_container_width=True):
            st.session_state.page = "الجلسات"
            st.rerun()
    with c3:
        if st.button("📄 التقارير والطباعة", use_container_width=True):
            st.session_state.page = "التقارير"
            st.rerun()
    with c4:
        if st.button("📂 كل الدعاوى", use_container_width=True):
            st.session_state.page = "الدعاوى"
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    today_str = fmt_date_ar(datetime.now().date())
    today_cases = [c for c in get_sorted_cases(cases) if fmt_date_ar(c.get("تاريخ_جلسة", "")) == today_str]
    if today_cases:
        st.markdown(f"### 📅 جلسات اليوم ({today_str}) — العدد: {len(today_cases)}")
        for i, c in enumerate(today_cases):
            render_case_card(c, index=i, data_ref=data)

    st.markdown("### 🕒 أحدث الدعاوى المسجلة")
    recent_cases = list(reversed(get_sorted_cases(cases)))[:5]
    if recent_cases:
        for i, c in enumerate(recent_cases):
            render_case_card(c, index=100 + i, data_ref=data)
    else:
        st.info("لا توجد دعاوى مسجلة بعد. يمكنك البدء بإضافة دعوى جديدة.")

def add_edit_case_page(data):
    cases = data["cases"]
    editing_id = st.session_state.get("editing_case_id")
    existing = next((c for c in cases if c.get("id") == editing_id), None) if editing_id else None

    st.markdown(f"## {'✏️ تعديل بيانات الدعوى' if existing else '➕ إضافة دعوى جديدة'}", unsafe_allow_html=True)

    with st.container(border=True):
        r1, r2, r3 = st.columns(3)
        with r1:
            case_no = st.text_input("رقم الدعوى", value=str(existing.get("رقم_الدعوى", "") if existing else ""), key="case_no_inp")
        with r2:
            case_year = st.text_input("سنة الدعوى", value=str(existing.get("سنة_الدعوى", datetime.now().year) if existing else datetime.now().year), key="case_year_inp")
        with r3:
            survey_no = st.text_input("رقم الحصر", value=str(existing.get("رقم_الحصر", "") if existing else ""), key="survey_no_inp")

        p1, p2, p3 = st.columns(3)
        with p1:
            plaintiff = st.text_input("المدعي", value=str(existing.get("المدعي", "") if existing else ""), key="plaintiff_inp")
        with p2:
            defendant = st.text_input("المدعى عليه", value=str(existing.get("المدعى_عليه", "") if existing else ""), key="defendant_inp")
        with p3:
            raw_party = existing.get("الطرف", "") if existing else ""
            party_clean = _strip_party_prefix(raw_party)
            party_name = st.text_input("الطرف (اسم المحامي / الصفة)", value=party_clean, placeholder="مثال: فلان الفلاني", key="party_inp")
            if party_name.strip():
                st.markdown(f"<div style='font-size:13px;margin-top:2px;'>الصياغة المعتمدة: {_format_party_display(party_name, html_mode=True)}</div>", unsafe_allow_html=True)

        m1, m2, m3 = st.columns(3)
        with m1:
            court = st.text_input("المحكمة", value=str(existing.get("المحكمة", "") if existing else ""), key="court_inp")
        with m2:
            circle = st.text_input("الدائرة", value=str(existing.get("الدائرة", "") if existing else ""), key="circle_inp")
        with m3:
            ctype = st.text_input("نوع الدعوى", value=str(existing.get("نوع_الدعوى", "") if existing else ""), key="ctype_inp")

        s1, s2 = st.columns(2)
        with s1:
            status_opts = ["متداولة", "محكوم فيها", "مؤجلة", "منتهية / مشطوبة"]
            cur_st = existing.get("حالة_الدعوى", "متداولة") if existing else "متداولة"
            if cur_st not in status_opts:
                status_opts.append(cur_st)
            status = st.selectbox("حالة الدعوى", status_opts, index=status_opts.index(cur_st), key="status_inp")
        with s2:
            session_date = date_input_dmy(
                "تاريخ الجلسة",
                value=existing.get("تاريخ_جلسة") if existing else datetime.now().date(),
                key="session_date_inp"
            )

        decision = st.text_area("قرار الجلسة", value=str(existing.get("قرار_الجلسة", "") if existing else ""), height=80, key="decision_inp")

        j1, j2 = st.columns([2, 1])
        with j1:
            judgment = st.text_area("منطوق الحكم (إن وجد)", value=str(existing.get("الحكم", "") if existing else ""), height=80, key="judgment_inp")
        with j2:
            has_j_date = bool(existing and existing.get("تاريخ_الحكم"))
            include_j_date = st.checkbox("تحديد تاريخ الحكم", value=has_j_date, key="inc_j_date")
            judgment_date = ""
            if include_j_date:
                judgment_date = date_input_dmy(
                    "تاريخ الحكم",
                    value=existing.get("تاريخ_الحكم") if (existing and existing.get("تاريخ_الحكم")) else datetime.now().date(),
                    key="judgment_date_inp"
                )

        notes = st.text_area("ملاحظات", value=str(existing.get("ملاحظات", "") if existing else ""), height=75, key="notes_inp")

        uploaded_docs = st.file_uploader(
            "📎 إرفاق مستندات أو صور أو ملفات (PDF, DOCX, XLSX, JPG, PNG)",
            type=list(ALLOWED_DOC_EXTENSIONS),
            accept_multiple_files=True,
            key="case_docs_uploader"
        )

        if st.button("💾 حفظ الدعوى", type="primary", use_container_width=True, key="save_case_btn"):
            if not case_no.strip() and not plaintiff.strip() and not defendant.strip():
                st.error("يرجى إدخال رقم الدعوى أو بيانات الخصوم.")
                return

            new_docs = list(existing.get("مستندات") or []) if existing else []
            for uf in (uploaded_docs or []):
                try:
                    fpath, orig_name, sp = save_uploaded_file(uf, prefix="case")
                    new_docs.append({
                        "name": orig_name,
                        "path": fpath,
                        "storage_path": sp,
                        "uploaded_at": fmt_date_ar(datetime.now().date())
                    })
                except ValueError as ve:
                    st.error(str(ve))
                    return

            sessions_list = list(existing.get("جلسات") or []) if existing else []
            if session_date or decision.strip() or notes.strip():
                prev_latest = get_latest_session_info(existing) if existing else ("", "", "")
                if (fmt_date_ar(session_date), decision.strip(), notes.strip()) != prev_latest:
                    sessions_list.append({
                        "تاريخ": fmt_date_ar(session_date),
                        "قرار": decision.strip(),
                        "ملاحظات": notes.strip()
                    })

            final_status = "محكوم فيها" if judgment.strip() and status == "متداولة" else status
            formatted_party = _format_party_display(party_name, html_mode=False)

            record = {
                "id": existing.get("id") if existing else secrets.token_hex(8),
                "رقم_الدعوى": case_no.strip(),
                "سنة_الدعوى": case_year.strip(),
                "رقم_الحصر": survey_no.strip(),
                "المدعي": plaintiff.strip(),
                "المدعى_عليه": defendant.strip(),
                "الطرف": formatted_party,
                "المحكمة": court.strip(),
                "الدائرة": circle.strip(),
                "نوع_الدعوى": ctype.strip(),
                "حالة_الدعوى": final_status,
                "تاريخ_جلسة": fmt_date_ar(session_date),
                "قرار_الجلسة": decision.strip(),
                "الحكم": judgment.strip(),
                "تاريخ_الحكم": fmt_date_ar(judgment_date) if include_j_date else "",
                "ملاحظات": notes.strip(),
                "جلسات": sessions_list,
                "مستندات": new_docs,
                "user_id": existing.get("user_id") if existing else (st.session_state.get("user") or {}).get("id"),
                "assigned_to": existing.get("assigned_to") if existing else (st.session_state.get("user") or {}).get("username", ""),
                "updated_at": fmt_date_ar(datetime.now().date())
            }

            if existing:
                data["cases"] = [record if x.get("id") == existing.get("id") else x for x in cases]
            else:
                data["cases"].append(record)

            save_data(data)
            st.session_state.pop("editing_case_id", None)
            st.success("تم حفظ بيانات الدعوى بنجاح ✓")
            st.session_state.page = "الدعاوى"
            st.rerun()

def cases_list_page(data):
    st.markdown("## 📂 سجل الدعاوى")
    cases = get_sorted_cases(data["cases"])

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        q = st.text_input("بحث شامل", placeholder="رقم الدعوى، الحصر، الخصوم، الدائرة...", key="cases_search_q")
    with f2:
        courts = ["الكل"] + sorted({c.get("المحكمة", "") for c in cases if c.get("المحكمة")})
        sel_court = st.selectbox("المحكمة", courts, key="cases_filter_court")
    with f3:
        statuses = ["الكل"] + sorted({c.get("حالة_الدعوى", "متداولة") for c in cases if c.get("حالة_الدعوى") or True})
        sel_status = st.selectbox("الحالة", statuses, key="cases_filter_status")
    with f4:
        parties = ["الكل"] + sorted({_format_party_display(c.get("الطرف", ""), html_mode=False) for c in cases if c.get("الطرف")})
        sel_party = st.selectbox("الطرف", parties, key="cases_filter_party")

    filtered = []
    for c in cases:
        if sel_court != "الكل" and c.get("المحكمة") != sel_court:
            continue
        if sel_status != "الكل" and (c.get("حالة_الدعوى") or "متداولة") != sel_status:
            continue
        if sel_party != "الكل" and _format_party_display(c.get("الطرف", ""), html_mode=False) != sel_party:
            continue
        if q.strip():
            ql = q.strip().lower()
            blob = " ".join([
                str(c.get("رقم_الدعوى", "")), str(c.get("سنة_الدعوى", "")), str(c.get("رقم_الحصر", "")),
                str(c.get("المدعي", "")), str(c.get("المدعى_عليه", "")), str(c.get("الطرف", "")),
                str(c.get("المحكمة", "")), str(c.get("الدائرة", "")), str(c.get("نوع_الدعوى", "")),
                str(c.get("قرار_الجلسة", "")), str(c.get("الحكم", "")), str(c.get("ملاحظات", ""))
            ]).lower()
            if ql not in blob:
                continue
        filtered.append(c)

    st.markdown(f"**عدد الدعاوى المطابقة:** {len(filtered)}")
    if not filtered:
        st.info("لا توجد دعاوى مطابقة لمعايير البحث.")
        return

    for idx, c in enumerate(filtered):
        render_case_card(c, index=idx, data_ref=data)

def sessions_page(data):
    st.markdown("## 📅 جدول الجلسات والقرارات")
    cases = get_sorted_cases(data["cases"])

    mode = st.radio("عرض الجلسات", ["جلسات اليوم", "جلسات حسب تاريخ محدد", "كل الدعاوى ذات الجلسات"], horizontal=True, key="sessions_mode")
    target_date = datetime.now().date()
    if mode == "جلسات حسب تاريخ محدد":
        target_date = parse_date_safe(date_input_dmy("اختر تاريخ الجلسة", value=datetime.now().date(), key="sessions_pick_date")) or datetime.now().date()

    target_str = fmt_date_ar(target_date)
    matched = []
    for c in cases:
        s_dates = [fmt_date_ar(d) for d in _case_session_dates(c)]
        if mode == "جلسات اليوم":
            if fmt_date_ar(datetime.now().date()) in s_dates:
                matched.append(c)
        elif mode == "جلسات حسب تاريخ محدد":
            if target_str in s_dates:
                matched.append(c)
        else:
            if s_dates or c.get("قرار_الجلسة"):
                matched.append(c)

    st.markdown(f"**عدد الدعاوى:** {len(matched)}")
    if not matched:
        st.info("لا توجد جلسات مسجلة في هذا النطاق.")
        return

    for idx, c in enumerate(matched):
        render_case_card(c, index=idx, data_ref=data)
        with st.expander(f"➕ تسجيل جلسة جديدة للدعوى {c.get('رقم_الدعوى','')}/{c.get('سنة_الدعوى','')}", expanded=False):
            new_s_date = date_input_dmy("تاريخ الجلسة الجديدة", value=datetime.now().date(), key=f"new_s_date_{c.get('id')}_{idx}")
            new_s_dec = st.text_input("قرار الجلسة", key=f"new_s_dec_{c.get('id')}_{idx}")
            new_s_note = st.text_input("ملاحظات", key=f"new_s_note_{c.get('id')}_{idx}")
            if st.button("حفظ الجلسة", key=f"save_new_s_{c.get('id')}_{idx}", type="primary"):
                for orig in data["cases"]:
                    if orig.get("id") == c.get("id"):
                        orig.setdefault("جلسات", []).append({
                            "تاريخ": fmt_date_ar(new_s_date),
                            "قرار": new_s_dec.strip(),
                            "ملاحظات": new_s_note.strip()
                        })
                        orig["تاريخ_جلسة"] = fmt_date_ar(new_s_date)
                        if new_s_dec.strip():
                            orig["قرار_الجلسة"] = new_s_dec.strip()
                        if new_s_note.strip():
                            orig["ملاحظات"] = new_s_note.strip()
                        break
                save_data(data)
                st.success("تم تسجيل الجلسة وتحديث الدعوى")
                st.rerun()

def reports_page(data):
    st.markdown("## 📄 التقارير والطباعة والتصدير")
    settings = data["settings"]
    cases = data["cases"]

    st.markdown("<div class='report-selector-strip'>", unsafe_allow_html=True)
    report_type = st.radio(
        "نوع التقرير",
        ["بيان الدعاوى", "بيان الأحكام", "الإحصائيات", "بيان عددي"],
        horizontal=True,
        key="report_type_radio_v2"
    )
    st.markdown("</div>", unsafe_allow_html=True)

    with st.expander("🎛️ تصفية بيانات التقرير والبحث", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            courts = ["الكل"] + sorted({c.get("المحكمة", "") for c in cases if c.get("المحكمة")})
            sel_court = st.selectbox("المحكمة", courts, key="rep_court")
        with c2:
            ctypes = ["الكل"] + sorted({c.get("نوع_الدعوى", "") for c in cases if c.get("نوع_الدعوى")})
            sel_ctype = st.selectbox("نوع الدعوى", ctypes, key="rep_ctype")
        with c3:
            circles = ["الكل"] + sorted({c.get("الدائرة", "") for c in cases if c.get("الدائرة")})
            sel_circle = st.selectbox("الدائرة", circles, key="rep_circle")
        with c4:
            parties = ["الكل"] + sorted({_format_party_display(c.get("الطرف", ""), html_mode=False) for c in cases if c.get("الطرف")})
            sel_party = st.selectbox("الطرف", parties, key="rep_party")

        d1, d2, d3, d4 = st.columns(4)
        with d1:
            years = ["الكل"] + sorted({str(c.get("سنة_الدعوى", "")) for c in cases if c.get("سنة_الدعوى")})
            sel_year = st.selectbox("سنة الدعوى", years, key="rep_year")
        with d2:
            use_from = st.checkbox("من تاريخ جلسة", value=False, key="rep_use_from")
            d_from = date_input_dmy("من تاريخ", value=datetime.now().date(), key="rep_d_from") if use_from else ""
        with d3:
            use_to = st.checkbox("إلى تاريخ جلسة", value=False, key="rep_use_to")
            d_to = date_input_dmy("إلى تاريخ", value=datetime.now().date(), key="rep_d_to") if use_to else ""
        with d4:
            q_search = st.text_input("بحث داخل التقرير", placeholder="رقم، خصوم، حكم، ملاحظات...", key="rep_search")

    filters = {
        "court": sel_court,
        "ctype": sel_ctype,
        "circle": sel_circle,
        "party": sel_party,
        "year": sel_year,
        "date_from": d_from,
        "date_to": d_to,
        "search": q_search,
    }

    html_report = render_html_report(cases, report_type, settings, filters)
    st.markdown(html_report, unsafe_allow_html=True)

    st.markdown("### 📤 تصدير التقرير")
    ex1, ex2, ex3 = st.columns(3)
    with ex1:
        word_bytes = to_word(cases, report_type, settings, filters)
        st.download_button(
            "📘 تحميل Word (.docx)",
            data=word_bytes,
            file_name=f"{report_type}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="dl_rep_word"
        )
    with ex2:
        excel_bytes = to_excel(cases, report_type, settings, filters)
        st.download_button(
            "📗 تحميل Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"{report_type}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_rep_excel"
        )
    with ex3:
        pdf_bytes = to_pdf(cases, report_type, settings, filters)
        st.download_button(
            "📕 تحميل PDF (.pdf)",
            data=pdf_bytes,
            file_name=f"{report_type}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="dl_rep_pdf"
        )

def active_cases_by_circle_page(data):
    st.markdown("## ⚖️ الدعاوى المتداولة حسب (المحكمة، النوع، الدائرة)")
    cases = [c for c in get_sorted_cases(data["cases"]) if (c.get("حالة_الدعوى") or "متداولة") == "متداولة"]

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        courts = ["الكل"] + sorted({c.get("المحكمة", "") for c in cases if c.get("المحكمة")})
        sel_court = st.selectbox("المحكمة", courts, key="ac_court")
    with c2:
        ctypes = ["الكل"] + sorted({c.get("نوع_الدعوى", "") for c in cases if c.get("نوع_الدعوى")})
        sel_ctype = st.selectbox("نوع الدعوى", ctypes, key="ac_ctype")
    with c3:
        circles = ["الكل"] + sorted({c.get("الدائرة", "") for c in cases if c.get("الدائرة")})
        sel_circle = st.selectbox("الدائرة", circles, key="ac_circle")
    with c4:
        q = st.text_input("بحث بالمحكمة أو النوع أو الدائرة أو الرقم", key="ac_search")

    filtered = []
    for c in cases:
        if sel_court != "الكل" and c.get("المحكمة") != sel_court:
            continue
        if sel_ctype != "الكل" and c.get("نوع_الدعوى") != sel_ctype:
            continue
        if sel_circle != "الكل" and c.get("الدائرة") != sel_circle:
            continue
        if q.strip():
            ql = q.strip().lower()
            blob = " ".join([
                str(c.get("المحكمة", "")), str(c.get("نوع_الدعوى", "")), str(c.get("الدائرة", "")),
                str(c.get("رقم_الدعوى", "")), str(c.get("سنة_الدعوى", "")), str(c.get("رقم_الحصر", "")),
                str(c.get("المدعي", "")), str(c.get("المدعى_عليه", "")), str(c.get("الطرف", ""))
            ]).lower()
            if ql not in blob:
                continue
        filtered.append(c)

    grouped = {}
    for c in filtered:
        grp_key = (
            c.get("المحكمة") or "غير محدد",
            c.get("نوع_الدعوى") or "غير محدد",
            c.get("الدائرة") or "غير محددة",
        )
        grouped.setdefault(grp_key, []).append(c)

    st.markdown(f"**إجمالي الدعاوى المتداولة المطابقة:** {len(filtered)}  |  **عدد المجموعات:** {len(grouped)}")
    if not grouped:
        st.info("لا توجد دعاوى متداولة مطابقة.")
        return

    for (ct, tp, cr), grp_cases in grouped.items():
        with st.expander(f"🏛️ المحكمة: {ct}  |  النوع: {tp}  |  الدائرة: {cr}  ({len(grp_cases)} دعوى)", expanded=True):
            for idx, c in enumerate(grp_cases):
                render_case_card(c, index=idx, data_ref=data)

def library_page(data):
    st.markdown("## 📚 المكتبة القانونية والمستندات")
    lib_items = data.setdefault("library", [])

    with st.expander("➕ إضافة مستند أو مرجع قانوني جديد", expanded=False):
        doc_title = st.text_input("عنوان المستند أو المرجع", key="lib_title")
        doc_cat = st.text_input("التصنيف", value="مذكرات وأحكام ومراجع", key="lib_cat")
        doc_notes = st.text_area("ملخص أو ملاحظات", key="lib_notes")
        doc_file = st.file_uploader("رفع ملف المستند", type=list(ALLOWED_DOC_EXTENSIONS), key="lib_file_uploader")

        if st.button("💾 حفظ في المكتبة", type="primary", key="save_lib_btn"):
            if not doc_title.strip() and not doc_file:
                st.error("يرجى إدخال عنوان المستند أو رفع ملف.")
            else:
                doc_obj = None
                if doc_file:
                    try:
                        fpath, orig_name, sp = save_uploaded_file(doc_file, prefix="lib")
                        doc_obj = {"name": orig_name, "path": fpath, "storage_path": sp}
                    except ValueError as ve:
                        st.error(str(ve))
                        return
                lib_items.append({
                    "id": secrets.token_hex(8),
                    "title": doc_title.strip() or (doc_obj["name"] if doc_obj else "مستند"),
                    "category": doc_cat.strip(),
                    "notes": doc_notes.strip(),
                    "doc": doc_obj,
                    "created_at": fmt_date_ar(datetime.now().date()),
                    "user_id": (st.session_state.get("user") or {}).get("id")
                })
                save_data(data)
                st.success("تم حفظ المستند في المكتبة")
                st.rerun()

    q = st.text_input("🔎 بحث في المكتبة", key="lib_search_q")
    shown = [
        item for item in lib_items
        if not q.strip() or q.strip().lower() in f"{item.get('title','')} {item.get('category','')} {item.get('notes','')}".lower()
    ]

    if not shown:
        st.info("لا توجد مستندات في المكتبة.")
        return

    for idx, item in enumerate(reversed(shown)):
        with st.container(border=True):
            st.markdown(f"### 📄 {item.get('title','')}")
            st.caption(f"التصنيف: {item.get('category','—')} | التاريخ: {fmt_date_ar(item.get('created_at',''), '—')}")
            if item.get("notes"):
                st.write(item.get("notes"))
            if item.get("doc"):
                _render_doc_viewer(item["doc"], key_prefix=f"lib_item_{item.get('id')}_{idx}")
            if st.button("🗑️ حذف المستند", key=f"del_lib_{item.get('id')}_{idx}"):
                if item.get("doc"):
                    _delete_doc_obj(item["doc"])
                data["library"] = [x for x in data["library"] if x.get("id") != item.get("id")]
                save_data(data)
                st.rerun()
def settings_page(data):
    st.markdown("## ⚙️ إعدادات المنظومة والتوقيعات")
    settings = data.setdefault("settings", {})

    with st.container(border=True):
        app_title = st.text_input("اسم المنظومة / عنوان التقارير", value=settings.get("app_title", "إدارة القضايا"), key="set_app_title")
        subtitle = st.text_input("العنوان الفرعي", value=settings.get("subtitle", "منظومة متكاملة لإدارة الدعاوى والقرارات والأحكام والمذكرات"), key="set_subtitle")
        theme = st.selectbox("ألوان الواجهة", list(THEMES.keys()), index=list(THEMES.keys()).index(settings.get("theme", "الأزرق الملكي والذهبي")) if settings.get("theme") in THEMES else 0, key="set_theme")

        st.markdown("### ✍️ بيانات التوقيعات في التقارير")
        s1, s2 = st.columns(2)
        with s1:
            member_title = st.text_input("مسمى وظيفة العضو القانوني", value=settings.get("member_title", "العضو القانوني"), key="set_member_title")
            member_name = st.text_input("اسم العضو القانوني", value=settings.get("member_name", ""), key="set_member_name")
        with s2:
            manager_title = st.text_input("مسمى وظيفة مدير إدارة القضايا", value=settings.get("manager_title", "مدير إدارة القضايا"), key="set_manager_title")
            manager_name = st.text_input("اسم مدير إدارة القضايا", value=settings.get("manager_name", ""), key="set_manager_name")

        if st.button("💾 حفظ الإعدادات", type="primary", use_container_width=True, key="save_settings_btn"):
            settings["app_title"] = app_title.strip() or "إدارة القضايا"
            settings["subtitle"] = subtitle.strip()
            settings["theme"] = theme
            settings["member_title"] = member_title.strip() or "العضو القانوني"
            settings["member_name"] = member_name.strip()
            settings["manager_title"] = manager_title.strip() or "مدير إدارة القضايا"
            settings["manager_name"] = manager_name.strip()
            save_data(data)
            st.success("تم حفظ الإعدادات والتوقيعات بنجاح ✓")
            st.rerun()

def main():
    if "user" not in st.session_state:
        st.session_state.user = None
    if "role" not in st.session_state:
        st.session_state.role = None
    if "page" not in st.session_state:
        st.session_state.page = "الرئيسية"

    data = load_data()
    apply_theme(data.get("settings", {}).get("theme", "الأزرق الملكي والذهبي"))

    if not st.session_state.user:
        login_page()
        return

    if st.session_state.page == "set_password":
        set_password_page()
        return

    user = st.session_state.user
    is_admin = (user.get("role") == "admin")

    with st.sidebar:
        if user.get("profile_photo"):
            st.markdown(_photo_html(user.get("profile_photo"), user.get("profile_photo_type") or "image/png", 84), unsafe_allow_html=True)
        st.markdown(f"<div style='text-align:center;color:#D4AF37;font-weight:900;font-size:18px;'>{html.escape(str(user.get('display_name') or user.get('username') or ''))}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='text-align:center;color:#CBD5E1;font-size:13px;margin-bottom:10px;'>{html.escape(str(user.get('job_title') or role_label(user.get('role'))))}</div>", unsafe_allow_html=True)

        nav_items = [
            ("الرئيسية", "🏠 الرئيسية"),
            ("إضافة دعوى", "➕ إضافة دعوى"),
            ("الدعاوى", "📂 سجل الدعاوى"),
            ("الجلسات", "📅 الجلسات والقرارات"),
            ("الدعاوى المتداولة حسب الدائرة", "⚖️ متداولة (محكمة/نوع/دائرة)"),
            ("التقارير", "📄 التقارير والطباعة"),
            ("المكتبة", "📚 المكتبة القانونية"),
            ("الرسائل", "📢 الرسائل والبيانات"),
            ("الإعدادات", "⚙️ الإعدادات"),
        ]
        if is_admin:
            nav_items.insert(1, ("لوحة التحكم", "🛡️ لوحة التحكم"))
            nav_items.insert(2, ("إدارة الأعضاء", "👥 إدارة الأعضاء"))
            nav_items.insert(3, ("extract_member", "🪪 استخراج عضوية"))

        for page_key, page_label in nav_items:
            btn_type = "primary" if st.session_state.page == page_key else "secondary"
            if st.button(page_label, key=f"nav_btn_{page_key}", use_container_width=True, type=btn_type):
                if page_key != "إضافة دعوى":
                    st.session_state.pop("editing_case_id", None)
                st.session_state.page = page_key
                st.rerun()

        st.markdown("---")
        if st.button("🚪 تسجيل الخروج", key="sidebar_logout_btn", use_container_width=True):
            st.session_state.user = None
            st.session_state.role = None
            st.session_state.page = "login"
            st.rerun()

    show_top_notices()

    cur_page = st.session_state.page
    if cur_page == "الرئيسية":
        home_page(data)
    elif cur_page == "لوحة التحكم":
        admin_control_page()
    elif cur_page == "إدارة الأعضاء":
        manage_users_page()
    elif cur_page == "تعديل_بيانات_العضويات":
        edit_memberships_page()
    elif cur_page == "extract_member":
        extract_member_page()
    elif cur_page == "إضافة دعوى":
        add_edit_case_page(data)
    elif cur_page == "الدعاوى":
        cases_list_page(data)
    elif cur_page == "الجلسات":
        sessions_page(data)
    elif cur_page == "الدعاوى المتداولة حسب الدائرة":
        active_cases_by_circle_page(data)
    elif cur_page == "التقارير":
        reports_page(data)
    elif cur_page == "المكتبة":
        library_page(data)
    elif cur_page == "الرسائل":
        messaging_page()
    elif cur_page == "الإعدادات":
        settings_page(data)
    else:
        home_page(data)

    show_member_directory()


if __name__ == "__main__":
    main()
