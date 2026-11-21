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


# ===== دوال توحيد وتنسيق التواريخ (يوم / شهر / سنة) بقائمة منسدلة =====
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


def fmt_date_ar(val, default=""):
    d = parse_date_safe(val)
    if d is None:
        return str(val) if val not in (None, "") else default
    return f"{d.day:02d} / {d.month:02d} / {d.year:04d}"


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


def format_lawyer_party_label(lawyer_name):
    raw = str(lawyer_name or "").strip()
    if not raw:
        return ""
    cleaned = re.sub(r"^(?:طرف\s+)?(?:الأستاذ|الاستاذ)\s*/?\s*", "", raw, flags=re.IGNORECASE).strip()
    cleaned = re.sub(r"\s+المحامي$", "", cleaned).strip() or raw
    return f"طرف الأستاذ / {cleaned} المحامي"


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

    # سطر واحد متصل للقوائم المنسدلة: اليوم ، الشهر ، السنة
    st.markdown(f"<div data-dmy-row='{html.escape(base_key)}' style='display:flex!important;flex-direction:row!important;align-items:center!important;gap:6px!important;width:100%!important;'>", unsafe_allow_html=True)
    c_d, c_m, c_y = st.columns([1, 1, 1.35], gap="small")
    with c_d:
        sel_d = st.selectbox(
            "اليوم",
            days,
            format_func=lambda d: f"{d:02d}",
            index=min(max(d_val.day - 1, 0), 30),
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
    st.markdown("</div>", unsafe_allow_html=True)

    final_d = date(int(sel_y), int(sel_m), _clamp_day_for_ym(sel_y, sel_m, sel_d))
    return fmt_date_ar(final_d)


# ===== التصميم القضائي الجديد =====
st.markdown("""
<style>
/* صفوف التاريخ القوائم المنسدلة في سطر واحد: يوم ، شهر ، سنة */
[data-testid="stHorizontalBlock"]:has([data-dmy-row]) {
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: nowrap !important;
    align-items: center !important;
    gap: 6px !important;
}
[data-testid="stHorizontalBlock"]:has([data-dmy-row]) > div[data-testid="column"] {
    flex: 1 1 0 !important;
    min-width: 0 !important;
    width: auto !important;
}

/* وضوح نصوص اختيارات التبويبات والتقارير */
[data-testid="stTabs"] button,
[data-testid="stTabs"] button * {
    color: #FFFFFF !important;
    font-weight: 900 !important;
    opacity: 1 !important;
}
[data-testid="stTabs"] button[aria-selected="true"],
[data-testid="stTabs"] button[aria-selected="true"] * {
    color: #FFD700 !important;
    border-bottom-color: #FFD700 !important;
}
[data-testid="stRadio"] label,
[data-testid="stRadio"] label *,
[data-testid="stRadio"] div[role="radiogroup"] label,
[data-testid="stRadio"] div[role="radiogroup"] label * {
    color: #FFFFFF !important;
    font-weight: 900 !important;
    opacity: 1 !important;
}

/* صفحة الدخول */
.login-shell{
    max-width:760px;margin:45px auto;padding:28px;
    background:linear-gradient(145deg,#111c33,#1f3154);
    border:2px solid #D4AF37;border-radius:24px;
    box-shadow:0 20px 55px rgba(0,0,0,.45);
}
.login-shell h1,.login-shell h2,.login-shell p,.login-shell label{color:#fff!important;}
div[data-testid="stTabs"] button{
    color:#fff!important;font-weight:900!important;font-size:18px!important;
}
div[data-testid="stTabs"] button[aria-selected="true"]{
    color:#D4AF37!important;border-bottom:3px solid #D4AF37!important;
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
[data-testid="stTabs"] button { color:#fff !important; opacity:1 !important; visibility:visible !important; }
[data-testid="stFileUploader"] section * { color:#fff !important; opacity:1 !important; visibility:visible !important; }
input, textarea, [data-baseweb="select"], [data-testid="stFileUploader"] {color:#111!important;}
label, [data-testid="stCheckbox"] label, [data-testid="stRadio"] label, [data-testid="stMultiSelect"] label, [data-testid="stSelectbox"] label {color:#fff!important;font-weight:900!important;}
[data-testid="stCheckbox"] p, [data-testid="stRadio"] p {color:#fff!important;font-weight:800!important;}
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
                user["age_action_at"] = datetime.now().strftime("%d / %m / %Y %H:%M")
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
                    user["age_action_at"] = datetime.now().strftime("%d / %m / %Y %H:%M")
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
                user["age_action_at"] = datetime.now().strftime("%d / %m / %Y %H:%M")
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
            "created_at": datetime.now().strftime("%d / %m / %Y %H:%M"),
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
                        "created_at": datetime.now().strftime("%d / %m / %Y %H:%M"),
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
                        "created_at": created.strftime("%d / %m / %Y %H:%M"),
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
    min-height: 52px !important;
    opacity: 1 !important;
    box-shadow: 0 6px 16px rgba(0,0,0,.22) !important;
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
}
div[data-testid="stTabs"] button[aria-selected="true"],
div[data-testid="stTabs"] button[aria-selected="true"] p,
div[data-testid="stTabs"] button[aria-selected="true"] span {
    color:#FFD700 !important;
}
div[data-testid="stTextInput"] label,
div[data-testid="stPasswordInput"] label,
div[data-testid="stDateInput"] label,
div[data-testid="stSelectbox"] label,
div[data-testid="stTextArea"] label {
    color:#ffffff !important;
    font-weight:900 !important;
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

def _report_signatures_html(عضو_قانوني, مدير_ادارة, مدير_عام, تحرير_تاريخ=None, dark_theme=True):
    dt_str = fmt_date_ar(تحرير_تاريخ or datetime.now().date())
    main_color = "#B8860B" if dark_theme else "#1E2A47"
    sub_color = "#FFFFFF" if dark_theme else "#000000"
    gm_color = "#ff6b6b" if dark_theme else "#c0392b"
    member_line = f"<div style='color:{sub_color}; font-weight:800; margin-top:4px;'>{html.escape(str(عضو_قانوني))}</div>" if str(عضو_قانوني or "").strip() else ""
    manager_line = f"<div style='color:{sub_color}; font-weight:800; margin-top:4px;'>{html.escape(str(مدير_ادارة))}</div>" if str(مدير_ادارة or "").strip() else ""
    gm_line = f"<div style='color:{sub_color}; font-weight:800; margin-top:4px;'>{html.escape(str(مدير_عام))}</div>" if str(مدير_عام or "").strip() else ""
    return f"""
    <div style="margin-top:25px; color:{main_color}; font-size:12px; width:100%; direction:rtl;">
        <p style="text-align:center; margin-bottom:18px; font-size:13px; font-weight:700;">وتفضلوا بقبول وافر الاحترام والتقدير،</p>
        <table style="width:100%; table-layout:fixed; border-collapse:collapse; border:none; direction:rtl;">
            <tr>
                <td style="width:40%; text-align:right; vertical-align:top; border:none; padding:6px 8px; overflow-wrap:break-word; word-break:normal;">
                    <div style="font-weight:900; font-size:13px; color:{main_color};">العضو القانوني</div>
                    {member_line}
                    <div style="margin-top:12px; color:{sub_color};">....................</div>
                    <div style="margin-top:14px; font-weight:800; color:{sub_color};">تحرر في: {dt_str}</div>
                </td>
                <td style="width:20%; border:none;"></td>
                <td style="width:40%; text-align:left; vertical-align:top; border:none; padding:6px 8px; overflow-wrap:break-word; word-break:normal;">
                    <div style="font-weight:900; font-size:13px; color:{main_color};">مدير إدارة القضايا</div>
                    {manager_line}
                    <div style="margin-top:12px; color:{sub_color};">....................</div>
                </td>
            </tr>
        </table>
        <div style="text-align:center; margin-top:20px; padding:6px 10px; overflow-wrap:break-word; word-break:normal;">
            <div style="font-weight:900; font-size:13px; color:{gm_color};">مدير عام الإدارات القانونية</div>
            {gm_line}
            <div style="margin-top:12px; color:{sub_color};">....................</div>
        </div>
    </div>
    """

def to_excel(df, title="التقرير"):
    if df is None or df.empty:
        df = pd.DataFrame([{"البيان": "لا توجد بيانات"}])
    df_out = df.fillna("").copy()
    for col in df_out.columns:
        if any(k in str(col) for k in ("تاريخ", "التاريخ")):
            df_out[col] = df_out[col].apply(lambda v: fmt_date_ar(v, str(v)))
        else:
            df_out[col] = df_out[col].astype(str)
    df_rtl = df_out.iloc[:, ::-1] if len(df_out.columns) > 1 else df_out
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

    rtl(f"تحريرًا في: {fmt_date_ar(datetime.now().date())}", 9, False, WD_ALIGN_PARAGRAPH.RIGHT, 4)
    rtl("وتفضلوا بقبول وافر الاحترام والتقدير،", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, 6)

    sig_tbl = doc.add_table(rows=1, cols=3)
    _set_table_fixed_layout(sig_tbl, usable_w)
    _set_row_cant_split(sig_tbl.rows[0])

    right_cell = sig_tbl.rows[0].cells[0]
    gap_cell = sig_tbl.rows[0].cells[1]
    left_cell = sig_tbl.rows[0].cells[2]

    right_cell.width = Inches(2.9)
    gap_cell.width = Inches(1.5)
    left_cell.width = Inches(2.9)

    _set_cell_margins(right_cell, top=40, bottom=40, left=60, right=60)
    _set_cell_margins(gap_cell, top=40, bottom=40, left=80, right=80)
    _set_cell_margins(left_cell, top=40, bottom=40, left=60, right=60)

    p_r = right_cell.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.space_before = Pt(0)
    p_r.paragraph_format.space_after = Pt(0)
    p_r.paragraph_format.line_spacing = 1.15
    p_r._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    r_text = "العضو القانوني\n" + (f"{member_name}\n" if str(member_name or "").strip() else "") + "...................."
    run_r = p_r.add_run(r_text)
    run_r.font.name = "Arial"
    run_r.font.size = Pt(9.5)
    run_r.bold = True

    p_l = left_cell.paragraphs[0]
    p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_l.paragraph_format.space_before = Pt(0)
    p_l.paragraph_format.space_after = Pt(0)
    p_l.paragraph_format.line_spacing = 1.15
    p_l._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    l_text = "مدير إدارة القضايا\n" + (f"{manager_name}\n" if str(manager_name or "").strip() else "") + "...................."
    run_l = p_l.add_run(l_text)
    run_l.font.name = "Arial"
    run_l.font.size = Pt(9.5)
    run_l.bold = True

    rtl("مدير عام الإدارات القانونية", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, 1)
    if str(general_manager or "").strip():
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
    """يحافظ على فواصل التاريخ (DD / MM / YYYY) بوضوح عند تحويل النص للـ PDF."""
    s = "" if text is None else str(text)
    return re.sub(r'(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})', r'\1 / \2 / \3', s)

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
    pdf.cell(0, 5.5, ar(f"تحريرًا في: {fmt_date_ar(datetime.now().date())}"), 0, 1, "R")
    pdf.ln(2)
    pdf.cell(0, 5.5, ar("وتفضلوا بقبول وافر الاحترام والتقدير،"), 0, 1, "C")
    pdf.ln(3)
    y_sig = pdf.get_y()
    pdf.set_xy(165, y_sig)
    pdf.cell(122, 5.5, ar("العضو القانوني"), 0, 0, "R")
    pdf.set_xy(10, y_sig)
    pdf.cell(122, 5.5, ar("مدير إدارة القضايا"), 0, 1, "L")

    cur_y = y_sig + 5.5
    if str(member_name or "").strip() or str(manager_name or "").strip():
        if str(member_name or "").strip():
            pdf.set_xy(165, cur_y)
            pdf.cell(122, 5.5, ar(member_name), 0, 0, "R")
        if str(manager_name or "").strip():
            pdf.set_xy(10, cur_y)
            pdf.cell(122, 5.5, ar(manager_name), 0, 0, "L")
        pdf.ln(5.5)
        cur_y = pdf.get_y()

    pdf.set_xy(165, cur_y)
    pdf.cell(122, 5.5, ar("...................."), 0, 0, "R")
    pdf.set_xy(10, cur_y)
    pdf.cell(122, 5.5, ar("...................."), 0, 1, "L")
    pdf.ln(5)
    pdf.cell(0, 5.5, ar("مدير عام الإدارات القانونية"), 0, 1, "C")
    if str(general_manager or "").strip():
        pdf.cell(0, 5.5, ar(general_manager), 0, 1, "C")
    pdf.cell(0, 5.5, ar("...................."), 0, 1, "C")
    raw = pdf.output()
    return raw.encode("latin-1", errors="replace") if isinstance(raw, str) else bytes(raw)
# =========== الجزء الثامن: التقارير ============
# ================= القسم 1 من 2 =================
def get_export_html(content_html, title):
    return f"""<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
<meta charset="UTF-8">
<title>{title}</title>
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');
    body {{ font-family: "Cairo", "Times New Roman", serif; direction: rtl; background: white; color: black; padding: 20px; }}
  .case-table {{width:100%; border-collapse:collapse; font-size:12px; margin-top:12px; direction:rtl;}}
  .case-table th {{ background: #B8860B; color:#000; padding:8px; border:1px solid #8B6914; font-weight:900; text-align:center; overflow-wrap:break-word; word-break:normal; white-space:normal; }}
  .case-table td {{padding:6px; border:1px solid #B8860B; text-align:center; background:#fff; color:#000; overflow-wrap:break-word; word-break:normal; white-space:normal;}}
  .table-container {{overflow-x:auto; width:100%;}}
    h1, h2, h3 {{ color: #8B6914!important; text-align: center; overflow-wrap:break-word; word-break:normal; }}
</style>
</head>
<body>
{content_html}
</body>
</html>"""

def _build_report_heading_text(base_title, lawyer_name=""):
    b = str(base_title or "").strip()
    l_label = format_lawyer_party_label(lawyer_name)
    if l_label:
        return f"{b}\n{l_label}"
    return b

def _date_in_range(val, d_from, d_to, fallback_dates=None):
    if d_from is None and d_to is None:
        return True
    candidates = []
    d_main = parse_date_safe(val)
    if d_main is not None:
        candidates.append(d_main)
    if fallback_dates:
        for fd in fallback_dates:
            if fd is not None:
                candidates.append(fd)
    if not candidates:
        return False
    for d in candidates:
        if d_from is not None and d < d_from:
            continue
        if d_to is not None and d > d_to:
            continue
        return True
    return False

def _clear_other_report_checkboxes(active_key, all_keys):
    for k in all_keys:
        if k != active_key:
            st.session_state[k] = False

if st.session_state.page == "تقارير":
    data = load_data()
    my_cases = get_my_cases(data)

    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');
  .case-table {width:100%; border-collapse:collapse; font-size:10px; margin-top:12px; direction:rtl; font-family: "Cairo", sans-serif;}
  .case-table th {background: linear-gradient(135deg, #8B6914 0%, #B8860B 100%); color:#000; padding:6px; border:1px solid #8B6914; font-weight:900; text-align:center; white-space: normal; overflow-wrap:break-word; word-break:normal; font-size:10px;}
  .case-table td {padding:5px; border:1px solid #B8860B; text-align:center; background:#1E2A47; color:#fff; font-size:10px; white-space: normal; overflow-wrap:break-word; word-break:normal;}
  .table-container {overflow-x:auto; width:100%;}
    h1, h2, h3, h4, label {color: #B8860B!important; font-family: "Cairo", sans-serif;}
  .stButton>button {border: 2px solid #B8860B!important; color: #B8860B!important; font-family: "Cairo", sans-serif; font-size:13px;}
  .stButton>button:hover {background: #B8860B!important; color: #000!important;}
    [data-testid="stTab"] button,
    [data-testid="stTab"] button p,
    [data-testid="stTab"] button span {
        color: #FFFFFF !important;
        font-family: "Cairo", sans-serif;
        font-size:13px;
        font-weight:900 !important;
        opacity:1 !important;
    }
    [data-testid="stTab"] button[aria-selected="true"],
    [data-testid="stTab"] button[aria-selected="true"] p,
    [data-testid="stTab"] button[aria-selected="true"] span {
        border-bottom: 3px solid #FFD700 !important;
        color: #FFD700 !important;
    }
    [data-testid="stRadio"] label,
    [data-testid="stRadio"] label *,
    [data-testid="stRadio"] div[role="radiogroup"] label,
    [data-testid="stRadio"] div[role="radiogroup"] label * {
        color: #FFFFFF !important;
        font-weight: 900 !important;
        opacity: 1 !important;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<h2 style='color:#B8860B; text-align:center; font-family: Cairo; font-size:18px;'>📑 مركز التقارير الحكومية</h2>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للرئيسية", use_container_width=True):
        st.session_state.page = "الرئيسية"; st.rerun()

    if 'last_report_html' not in st.session_state:
        st.session_state.last_report_html = ""; st.session_state.last_report_title = "تقرير"
    if 'last_report_df' not in st.session_state:
        st.session_state.last_report_df = pd.DataFrame()
    if 'last_report_meta' not in st.session_state:
        st.session_state.last_report_meta = {"region": "", "member": "", "manager": "", "gm": ""}

    tab1, tab2, tab3, tab4 = st.tabs(["📊 بيان بجميع الدعاوى المتداولة", "⚖️ بيان الاحكام", "📈 الإحصائيات", "📊 بيان عددي"])

    def report_header(region, title, مدير_عام, مدير_ادارة, عضو_قانوني, lawyer_name=""):
        parts = [p.strip() for p in str(title or "").splitlines() if p.strip()]
        main_title = parts[0] if parts else ""
        sub_line = parts[1] if len(parts) > 1 else format_lawyer_party_label(lawyer_name)
        sub_html = f"<div style='margin:6px 0 2px; font-size:13px; font-weight:900; text-align:center; color:#ff3b30; white-space:normal;'>{html.escape(sub_line)}</div>" if sub_line else ""
        return f"""
        <div style="text-align:right; margin-bottom:6px; line-height:1.35; color:#FFFFFF; font-size:11px; direction:rtl;">
            <p style="margin:1px 0;"><b>الهيئة القومية للتأمين الاجتماعي</b></p>
            <p style="margin:1px 0;">الإدارة المركزية للإدارات القانونية</p>
            <p style="margin:1px 0;">الإدارة العامة للقضايا</p>
            <p style="margin:1px 0;">ديوان عام منطقة: <b style="color:#B8860B;">{html.escape(str(region or ''))}</b></p>
        </div>
        <h3 style="margin:8px 0 4px; font-size:14px; text-align:center; color:#B8860B; white-space:normal;">{html.escape(main_title)}</h3>
        {sub_html}
        """

    def report_footer(مدير_عام, مدير_ادارة, عضو_قانوني):
        return _report_signatures_html(عضو_قانوني, مدير_ادارة, مدير_عام, datetime.now().date(), dark_theme=True)

    # ------------------------------------------------------------------
    # التبويب الأول: بيان بجميع الدعاوى المتداولة
    # ------------------------------------------------------------------
    with tab1:
        st.markdown("<h3 style='color:#B8860B; text-align:center; font-size:15px;'>📊 بيان بجميع الدعاوى المتداولة</h3>", unsafe_allow_html=True)
        with st.container(border=True):
            t1_chk_keys = ["t1_chk_all", "t1_chk_subject", "t1_chk_court_type_circle"]
            for k in t1_chk_keys:
                st.session_state.setdefault(k, k == "t1_chk_all")

            c_t1_1, c_t1_2, c_t1_3 = st.columns(3)
            with c_t1_1:
                st.checkbox("بيان بجميع الدعاوى المتداولة", key="t1_chk_all", on_change=_clear_other_report_checkboxes, args=("t1_chk_all", t1_chk_keys))
            with c_t1_2:
                st.checkbox("بيان بالدعاوى المتداولة حسب موضوع الدعوى", key="t1_chk_subject", on_change=_clear_other_report_checkboxes, args=("t1_chk_subject", t1_chk_keys))
            with c_t1_3:
                st.checkbox("بيان الدعاوى المتداولة حسب المحكمة ، النوع ، الدائرة", key="t1_chk_court_type_circle", on_change=_clear_other_report_checkboxes, args=("t1_chk_court_type_circle", t1_chk_keys))

            if st.session_state.get("t1_chk_court_type_circle"):
                نوع_التقرير_1 = "بيان الدعاوى المتداولة حسب المحكمة ، النوع ، الدائرة"
            elif st.session_state.get("t1_chk_subject"):
                نوع_التقرير_1 = "بيان بالدعاوى المتداولة حسب موضوع الدعوى"
            else:
                st.session_state["t1_chk_all"] = True
                نوع_التقرير_1 = "بيان بجميع الدعاوى المتداولة"

            فلتر_موضوع_1 = ""
            فلتر_محكمة_1 = ""
            فلتر_نوع_1 = ""
            فلتر_دائرة_1 = ""
            if نوع_التقرير_1 == "بيان بالدعاوى المتداولة حسب موضوع الدعوى":
                فلتر_موضوع_1 = st.text_input("موضوع الدعوى", placeholder="اكتب موضوع الدعوى أو جزء منه...", key="t1_subject")
            elif نوع_التقرير_1 == "بيان الدعاوى المتداولة حسب المحكمة ، النوع ، الدائرة":
                c1_m, c1_n, c1_d = st.columns(3)
                with c1_m:
                    فلتر_محكمة_1 = st.text_input("المحكمة", placeholder="اكتب اسم المحكمة...", key="t1_court")
                with c1_n:
                    فلتر_نوع_1 = st.text_input("النوع", placeholder="دعوى / استئناف / طعن...", key="t1_type")
                with c1_d:
                    فلتر_دائرة_1 = st.text_input("الدائرة", placeholder="اكتب رقم أو اسم الدائرة...", key="t1_circle")

            تفعيل_تاريخ_1 = st.checkbox("📅 تصفية بفترة زمنية (من تاريخ / إلى تاريخ)", value=False, key="t1_use_date")
            من_تاريخ_1_str = ""
            الى_تاريخ_1_str = ""
            if تفعيل_تاريخ_1:
                c_d1, c_d2 = st.columns(2)
                with c_d1:
                    من_تاريخ_1_str = date_input_dmy("من تاريخ (يوم / شهر / سنة)", value=datetime.now().date().replace(day=1), key="t1_from_date")
                with c_d2:
                    الى_تاريخ_1_str = date_input_dmy("إلى تاريخ (يوم / شهر / سنة)", value=datetime.now().date(), key="t1_to_date")

            col_r1, col_r2 = st.columns(2)
            with col_r1:
                منطقة1 = st.text_input("اسم المنطقة", placeholder="مثال: البحيرة", key="r1")
            with col_r2:
                محامي1 = st.text_input("طرف الأستاذ / المحامي", placeholder="اكتب اسم المحامي...", key="l1")

            col1, col2, col3 = st.columns(3)
            with col1: مدير_عام1 = st.text_input("مدير عام الإدارات القانونية", key="gm1")
            with col2: مدير_ادارة1 = st.text_input("مدير إدارة القضايا", key="m1")
            with col3: عضو_قانوني1 = st.text_input("العضو القانوني", key="u1")

        if st.button("📊 عرض وتجهيز البيان للطباعة", key="b1", use_container_width=True):
            قضايا = [c for c in my_cases if str(c.get('حالة', '')).strip() == 'متداولة']

            if نوع_التقرير_1 == "بيان بالدعاوى المتداولة حسب موضوع الدعوى":
                if not فلتر_موضوع_1.strip():
                    st.error("❌ من فضلك اكتب موضوع الدعوى")
                    st.stop()
                q_subj = _norm_ar(فلتر_موضوع_1)
                قضايا = [c for c in قضايا if q_subj in _norm_ar(c.get('موضوع', ''))]
                عنوان_1 = f"بيان بالدعاوى المتداولة عن موضوع: {فلتر_موضوع_1.strip()}"
            elif نوع_التقرير_1 == "بيان الدعاوى المتداولة حسب المحكمة ، النوع ، الدائرة":
                if not any(x.strip() for x in (فلتر_محكمة_1, فلتر_نوع_1, فلتر_دائرة_1)):
                    st.error("❌ من فضلك اكتب المحكمة أو النوع أو الدائرة")
                    st.stop()
                if فلتر_محكمة_1.strip():
                    q_court = _norm_ar(فلتر_محكمة_1)
                    قضايا = [c for c in قضايا if q_court in _norm_ar(c.get('محكمة_اسم', ''))]
                if فلتر_نوع_1.strip():
                    q_type = _norm_ar(فلتر_نوع_1)
                    قضايا = [c for c in قضايا if q_type in _norm_ar(c.get('نوع', ''))]
                if فلتر_دائرة_1.strip():
                    q_circ = _norm_ar(فلتر_دائرة_1)
                    قضايا = [c for c in قضايا if q_circ in _norm_ar(c.get('دائرة', ''))]
                parts_title = []
                if فلتر_محكمة_1.strip():
                    parts_title.append(f"محكمة: {فلتر_محكمة_1.strip()}")
                if فلتر_نوع_1.strip():
                    parts_title.append(f"النوع: {فلتر_نوع_1.strip()}")
                if فلتر_دائرة_1.strip():
                    parts_title.append(f"الدائرة: {فلتر_دائرة_1.strip()}")
                عنوان_1 = "بيان الدعاوى المتداولة حسب " + " - ".join(parts_title)
            else:
                عنوان_1 = "بيان بجميع الدعاوى المتداولة"

            if تفعيل_تاريخ_1:
                d_from_1 = parse_date_safe(من_تاريخ_1_str)
                d_to_1 = parse_date_safe(الى_تاريخ_1_str)
                if d_from_1 and d_to_1 and d_from_1 > d_to_1:
                    d_from_1, d_to_1 = d_to_1, d_from_1
                قضايا = [
                    c for c in قضايا
                    if _date_in_range(c.get('تاريخ_جلسة'), d_from_1, d_to_1, fallback_dates=_case_session_dates(c))
                ]
                عنوان_1 += f" (خلال الفترة من {fmt_date_ar(d_from_1)} إلى {fmt_date_ar(d_to_1)})"

            if not قضايا:
                st.warning("لا توجد دعاوى متداولة مطابقة للبحث")
            else:
                عنوان_1_كامل = _build_report_heading_text(عنوان_1, محامي1)
                html_rep = report_header(منطقة1, عنوان_1_كامل, مدير_عام1, مدير_ادارة1, عضو_قانوني1, محامي1)
                html_rep += '<div class="table-container"><table class="case-table"><tr><th>م</th><th>الرقم والسنة</th><th>النوع</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>تاريخ الجلسة</th><th>الاجراء</th><th>ملاحظات</th></tr>'
                rows_df = []
                for idx, c in enumerate(قضايا, 1):
                    رقم_كامل = f"{c.get('رقم', '')} لسنة {c.get('سنة', '')}"
                    محكمة_كاملة = f"{c.get('محكمة_اسم', '')}"
                    if c.get('مأمورية'): محكمة_كاملة += f" - مأمورية {c.get('مأمورية')}"
                    if c.get('دائرة'): محكمة_كاملة += f" - دائرة {c.get('دائرة')}"
                    نوع = c.get('نوع', '')
                    if "استئناف" in نوع: لقب1, لقب2 = "المستأنف", "المستأنف ضده"
                    elif "طعن" in نوع: لقب1, لقب2 = "الطاعن", "المطعون ضده"
                    else: لقب1, لقب2 = "المدعي", "المدعى عليه"
                    خصوم = f"<b>{لقب1}:</b> {html.escape(str(c.get('مدعي', '')))}<br><b>{لقب2}:</b> {html.escape(str(c.get('مدعي_عليه', '')))}"
                    خصوم_نص = f"{لقب1}: {c.get('مدعي', '')} / {لقب2}: {c.get('مدعي_عليه', '')}"
                    ملاحظات_نص = str(c.get('ملاحظات', '') or '')
                    if not ملاحظات_نص and c.get('جلسات'):
                        ملاحظات_نص = str(c['جلسات'][-1].get('ملاحظات', '') or '')
                    t_sess = fmt_date_ar(c.get('تاريخ_جلسة', ''))
                    html_rep += f"<tr><td>{idx}</td><td>{html.escape(رقم_كامل)}</td><td>{html.escape(str(نوع))}</td><td>{html.escape(محكمة_كاملة)}</td><td style='text-align:right'>{خصوم}</td><td>{html.escape(str(c.get('موضوع', '')))}</td><td style='white-space:nowrap'>{html.escape(t_sess)}</td><td>{html.escape(str(c.get('الاجراء', '')))}</td><td>{html.escape(ملاحظات_نص)}</td></tr>"
                    rows_df.append({
                        "م": idx,
                        "الرقم والسنة": رقم_كامل,
                        "النوع": نوع,
                        "المحكمة والدائرة": محكمة_كاملة,
                        "الخصوم": خصوم_نص,
                        "الموضوع": c.get('موضوع', ''),
                        "تاريخ الجلسة": t_sess,
                        "الاجراء": c.get('الاجراء', ''),
                        "ملاحظات": ملاحظات_نص
                    })
                html_rep += '</table></div>'
                html_rep += report_footer(مدير_عام1, مدير_ادارة1, عضو_قانوني1)
                st.session_state.last_report_html = html_rep
                st.session_state.last_report_title = عنوان_1_كامل
                st.session_state.last_report_df = pd.DataFrame(rows_df)
                st.session_state.last_report_meta = {"region": منطقة1, "member": عضو_قانوني1, "manager": مدير_ادارة1, "gm": مدير_عام1}
                st.markdown(html_rep, unsafe_allow_html=True)

    # ------------------------------------------------------------------
    # التبويب الثاني: بيان الأحكام
    # ------------------------------------------------------------------
    with tab2:
        st.markdown("<h3 style='color:#B8860B; text-align:center; font-size:15px;'>⚖️ بيان الاحكام</h3>", unsafe_allow_html=True)
        with st.container(border=True):
            t2_chk_keys = ["t2_chk_all", "t2_chk_win", "t2_chk_loss", "t2_chk_subject"]
            for k in t2_chk_keys:
                st.session_state.setdefault(k, k == "t2_chk_all")

            c_t2_1, c_t2_2, c_t2_3, c_t2_4 = st.columns(4)
            with c_t2_1:
                st.checkbox("بيان بجميع الأحكام الصادرة", key="t2_chk_all", on_change=_clear_other_report_checkboxes, args=("t2_chk_all", t2_chk_keys))
            with c_t2_2:
                st.checkbox("بيان بالأحكام الصادرة لصالح الهيئة", key="t2_chk_win", on_change=_clear_other_report_checkboxes, args=("t2_chk_win", t2_chk_keys))
            with c_t2_3:
                st.checkbox("بيان بالأحكام الصادرة ضد الهيئة", key="t2_chk_loss", on_change=_clear_other_report_checkboxes, args=("t2_chk_loss", t2_chk_keys))
            with c_t2_4:
                st.checkbox("بيان بالأحكام حسب موضوع الدعوى", key="t2_chk_subject", on_change=_clear_other_report_checkboxes, args=("t2_chk_subject", t2_chk_keys))

            if st.session_state.get("t2_chk_win"):
                نوع_بيان_الاحكام = "بيان بالأحكام الصادرة لصالح الهيئة"
            elif st.session_state.get("t2_chk_loss"):
                نوع_بيان_الاحكام = "بيان بالأحكام الصادرة ضد الهيئة"
            elif st.session_state.get("t2_chk_subject"):
                نوع_بيان_الاحكام = "بيان بالأحكام حسب موضوع الدعوى"
            else:
                st.session_state["t2_chk_all"] = True
                نوع_بيان_الاحكام = "بيان بجميع الأحكام الصادرة"

            فلتر_موضوع_2 = ""
            if نوع_بيان_الاحكام == "بيان بالأحكام حسب موضوع الدعوى":
                فلتر_موضوع_2 = st.text_input("موضوع الدعوى", placeholder="اكتب موضوع الدعوى أو جزء منه...", key="t2_subject")

            تفعيل_تاريخ_2 = st.checkbox("📅 تصفية بفترة زمنية (من تاريخ الحكم / إلى تاريخ الحكم)", value=False, key="t2_use_date")
            من_تاريخ_2_str = ""
            الى_تاريخ_2_str = ""
            if تفعيل_تاريخ_2:
                c2_d1, c2_d2 = st.columns(2)
                with c2_d1:
                    من_تاريخ_2_str = date_input_dmy("من تاريخ الحكم (يوم / شهر / سنة)", value=datetime.now().date().replace(day=1), key="t2_from_date")
                with c2_d2:
                    الى_تاريخ_2_str = date_input_dmy("إلى تاريخ الحكم (يوم / شهر / سنة)", value=datetime.now().date(), key="t2_to_date")

            col2_r1, col2_r2 = st.columns(2)
            with col2_r1:
                منطقة2 = st.text_input("اسم المنطقة", placeholder="مثال: البحيرة", key="r2")
            with col2_r2:
                محامي2 = st.text_input("طرف الأستاذ / المحامي", placeholder="اكتب اسم المحامي...", key="l2")

            col1, col2, col3 = st.columns(3)
            with col1: مدير_عام2 = st.text_input("مدير عام الإدارات القانونية", key="gm2")
            with col2: مدير_ادارة2 = st.text_input("مدير إدارة القضايا", key="m2")
            with col3: عضو_قانوني2 = st.text_input("العضو القانوني", key="u2")

        if st.button("⚖️ عرض وتجهيز بيان الاحكام", key="b2", use_container_width=True):
            قضايا = [c for c in my_cases if str(c.get('حالة', '')).strip() == 'منتهية']

            if نوع_بيان_الاحكام == "بيان بالأحكام الصادرة لصالح الهيئة":
                قضايا = [c for c in قضايا if str(c.get('مسندة_ل_الحكم', '')).strip() == 'الصالح']
                عنوان_2 = "بيان بالأحكام الصادرة لصالح الهيئة"
            elif نوع_بيان_الاحكام == "بيان بالأحكام الصادرة ضد الهيئة":
                قضايا = [c for c in قضايا if str(c.get('مسندة_ل_الحكم', '')).strip() == 'الضد']
                عنوان_2 = "بيان بالأحكام الصادرة ضد الهيئة"
            elif نوع_بيان_الاحكام == "بيان بالأحكام حسب موضوع الدعوى":
                if not فلتر_موضوع_2.strip():
                    st.error("❌ من فضلك اكتب موضوع الدعوى")
                    st.stop()
                q_subj2 = _norm_ar(فلتر_موضوع_2)
                قضايا = [c for c in قضايا if q_subj2 in _norm_ar(c.get('موضوع', ''))]
                عنوان_2 = f"بيان بالأحكام الصادرة عن موضوع: {فلتر_موضوع_2.strip()}"
            else:
                عنوان_2 = "بيان بجميع الأحكام الصادرة"

            if تفعيل_تاريخ_2:
                d_from_2 = parse_date_safe(من_تاريخ_2_str)
                d_to_2 = parse_date_safe(الى_تاريخ_2_str)
                if d_from_2 and d_to_2 and d_from_2 > d_to_2:
                    d_from_2, d_to_2 = d_to_2, d_from_2
                قضايا = [
                    c for c in قضايا
                    if _date_in_range(c.get('تاريخ_الحكم') or c.get('تاريخ_جلسة'), d_from_2, d_to_2, fallback_dates=_case_session_dates(c))
                ]
                عنوان_2 += f" (خلال الفترة من {fmt_date_ar(d_from_2)} إلى {fmt_date_ar(d_to_2)})"

            if not قضايا:
                st.warning("لا توجد احكام مسجلة مطابقة للبحث")
            else:
                عنوان_2_كامل = _build_report_heading_text(عنوان_2, محامي2)
                html_rep = report_header(منطقة2, عنوان_2_كامل, مدير_عام2, مدير_ادارة2, عضو_قانوني2, محامي2)
                html_rep += '<div class="table-container"><table class="case-table"><tr><th>م</th><th>الرقم والسنة</th><th>النوع</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>تاريخ الحكم</th><th>مسندة لـ</th><th>منطوق الحكم</th></tr>'
                rows_df = []
                for idx, c in enumerate(قضايا, 1):
                    رقم_كامل = f"{c.get('رقم', '')} لسنة {c.get('سنة', '')}"
                    محكمة_كاملة = f"{c.get('محكمة_اسم', '')}"
                    if c.get('مأمورية'): محكمة_كاملة += f" - مأمورية {c.get('مأمورية')}"
                    if c.get('دائرة'): محكمة_كاملة += f" - دائرة {c.get('دائرة')}"
                    نوع = c.get('نوع', '')
                    if "استئناف" in نوع: لقب1, لقب2 = "المستأنف", "المستأنف ضده"
                    elif "طعن" in نوع: لقب1, لقب2 = "الطاعن", "المطعون ضده"
                    else: لقب1, لقب2 = "المدعي", "المدعى عليه"
                    خصوم = f"<b>{لقب1}:</b> {html.escape(str(c.get('مدعي', '')))}<br><b>{لقب2}:</b> {html.escape(str(c.get('مدعي_عليه', '')))}"
                    خصوم_نص = f"{لقب1}: {c.get('مدعي', '')} / {لقب2}: {c.get('مدعي_عليه', '')}"
                    t_judg = fmt_date_ar(c.get('تاريخ_الحكم') or c.get('تاريخ_جلسة', ''))
                    مسندة = str(c.get('مسندة_ل_الحكم', '') or '')
                    لون_مسندة = "#4CAF50" if مسندة == "الصالح" else ("#FF5252" if مسندة == "الضد" else "#FFF")
                    html_rep += f"<tr><td>{idx}</td><td>{html.escape(رقم_كامل)}</td><td>{html.escape(str(نوع))}</td><td>{html.escape(محكمة_كاملة)}</td><td style='text-align:right'>{خصوم}</td><td>{html.escape(str(c.get('موضوع', '')))}</td><td style='white-space:nowrap'>{html.escape(t_judg)}</td><td style='color:{لون_مسندة}; font-weight:900'>{html.escape(مسندة)}</td><td>{html.escape(str(c.get('منطوق_الحكم', '')))}</td></tr>"
                    rows_df.append({
                        "م": idx,
                        "الرقم والسنة": رقم_كامل,
                        "النوع": نوع,
                        "المحكمة والدائرة": محكمة_كاملة,
                        "الخصوم": خصوم_نص,
                        "الموضوع": c.get('موضوع', ''),
                        "تاريخ الحكم": t_judg,
                        "مسندة لـ": مسندة,
                        "منطوق الحكم": c.get('منطوق_الحكم', '')
                    })
                html_rep += '</table></div>'
                html_rep += report_footer(مدير_عام2, مدير_ادارة2, عضو_قانوني2)
                st.session_state.last_report_html = html_rep
                st.session_state.last_report_title = عنوان_2_كامل
                st.session_state.last_report_df = pd.DataFrame(rows_df)
                st.session_state.last_report_meta = {"region": منطقة2, "member": عضو_قانوني2, "manager": مدير_ادارة2, "gm": مدير_عام2}
                st.markdown(html_rep, unsafe_allow_html=True)

    # ------------------------------------------------------------------
    # التبويب الثالث: الإحصائيات
    # ------------------------------------------------------------------
    with tab3:
        st.markdown("<h3 style='color:#B8860B; text-align:center; font-size:15px;'>📈 الإحصائيات</h3>", unsafe_allow_html=True)
        with st.container(border=True):
            t3_chk_keys = ["t3_chk_summary"]
            for k in t3_chk_keys:
                st.session_state.setdefault(k, True)
            st.checkbox("بيان إحصائي بحركة القضايا والأحكام", key="t3_chk_summary", on_change=_clear_other_report_checkboxes, args=("t3_chk_summary", t3_chk_keys))
            st.session_state["t3_chk_summary"] = True

            تفعيل_تاريخ_3 = st.checkbox("📅 تصفية الإحصائيات بفترة زمنية (من تاريخ / إلى تاريخ)", value=False, key="t3_use_date")
            من_تاريخ_3_str = ""
            الى_تاريخ_3_str = ""
            if تفعيل_تاريخ_3:
                c3_d1, c3_d2 = st.columns(2)
                with c3_d1:
                    من_تاريخ_3_str = date_input_dmy("من تاريخ (يوم / شهر / سنة)", value=datetime.now().date().replace(day=1), key="t3_from_date")
                with c3_d2:
                    الى_تاريخ_3_str = date_input_dmy("إلى تاريخ (يوم / شهر / سنة)", value=datetime.now().date(), key="t3_to_date")

            col3_r1, col3_r2 = st.columns(2)
            with col3_r1:
                منطقة3 = st.text_input("اسم المنطقة", placeholder="مثال: البحيرة", key="r3")
            with col3_r2:
                محامي3 = st.text_input("طرف الأستاذ / المحامي", placeholder="اكتب اسم المحامي...", key="l3")

            col1, col2, col3 = st.columns(3)
            with col1: مدير_عام3 = st.text_input("مدير عام الإدارات القانونية", key="gm3")
            with col2: مدير_ادارة3 = st.text_input("مدير إدارة القضايا", key="m3")
            with col3: عضو_قانوني3 = st.text_input("العضو القانوني", key="u3")

        if st.button("📈 عرض وتجهيز الإحصائيات", key="b3", use_container_width=True):
            try:
                قضايا_احصاء = list(my_cases)
                عنوان_3 = "بيان إحصائي بحركة القضايا والأحكام"

                if تفعيل_تاريخ_3:
                    d_from_3 = parse_date_safe(من_تاريخ_3_str)
                    d_to_3 = parse_date_safe(الى_تاريخ_3_str)
                    if d_from_3 and d_to_3 and d_from_3 > d_to_3:
                        d_from_3, d_to_3 = d_to_3, d_from_3
                    قضايا_احصاء = [
                        c for c in قضايا_احصاء
                        if _date_in_range(
                            c.get('تاريخ_الحكم') if str(c.get('حالة', '')).strip() == 'منتهية' else c.get('تاريخ_جلسة'),
                            d_from_3,
                            d_to_3,
                            fallback_dates=_case_session_dates(c) + ([parse_date_safe(c.get('تاريخ_الحكم'))] if c.get('تاريخ_الحكم') else [])
                        )
                    ]
                    عنوان_3 += f" (خلال الفترة من {fmt_date_ar(d_from_3)} إلى {fmt_date_ar(d_to_3)})"

                انواع_رئيسية = ["دعوى", "استئناف", "طعن"]
                rows_stat = []
                tot_all = tot_active = tot_ended = tot_win = tot_loss = 0

                for idx_t, t_name in enumerate(انواع_رئيسية, 1):
                    subset = [c for c in قضايا_احصاء if str(c.get('نوع', '')).strip() == t_name]
                    cnt_all = len(subset)
                    cnt_active = sum(1 for c in subset if str(c.get('حالة', '')).strip() == 'متداولة')
                    ended_sub = [c for c in subset if str(c.get('حالة', '')).strip() == 'منتهية']
                    cnt_ended = len(ended_sub)
                    cnt_win = sum(1 for c in ended_sub if str(c.get('مسندة_ل_الحكم', '')).strip() == 'الصالح')
                    cnt_loss = sum(1 for c in ended_sub if str(c.get('مسندة_ل_الحكم', '')).strip() == 'الضد')
                    win_pct = f"{round((cnt_win / cnt_ended) * 100, 1)}%" if cnt_ended > 0 else "0%"

                    tot_all += cnt_all
                    tot_active += cnt_active
                    tot_ended += cnt_ended
                    tot_win += cnt_win
                    tot_loss += cnt_loss

                    rows_stat.append({
                        "م": idx_t,
                        "نوع القضية": t_name,
                        "إجمالي القضايا": cnt_all,
                        "المتداولة": cnt_active,
                        "المنتهية بأحكام": cnt_ended,
                        "أحكام لصالح الهيئة": cnt_win,
                        "أحكام ضد الهيئة": cnt_loss,
                        "نسبة الصالح": win_pct,
                    })

                other_subset = [c for c in قضايا_احصاء if str(c.get('نوع', '')).strip() not in انواع_رئيسية]
                if other_subset:
                    cnt_all = len(other_subset)
                    cnt_active = sum(1 for c in other_subset if str(c.get('حالة', '')).strip() == 'متداولة')
                    ended_sub = [c for c in other_subset if str(c.get('حالة', '')).strip() == 'منتهية']
                    cnt_ended = len(ended_sub)
                    cnt_win = sum(1 for c in ended_sub if str(c.get('مسندة_ل_الحكم', '')).strip() == 'الصالح')
                    cnt_loss = sum(1 for c in ended_sub if str(c.get('مسندة_ل_الحكم', '')).strip() == 'الضد')
                    win_pct = f"{round((cnt_win / cnt_ended) * 100, 1)}%" if cnt_ended > 0 else "0%"

                    tot_all += cnt_all
                    tot_active += cnt_active
                    tot_ended += cnt_ended
                    tot_win += cnt_win
                    tot_loss += cnt_loss

                    rows_stat.append({
                        "م": len(rows_stat) + 1,
                        "نوع القضية": "أخرى",
                        "إجمالي القضايا": cnt_all,
                        "المتداولة": cnt_active,
                        "المنتهية بأحكام": cnt_ended,
                        "أحكام لصالح الهيئة": cnt_win,
                        "أحكام ضد الهيئة": cnt_loss,
                        "نسبة الصالح": win_pct,
                    })

                tot_win_pct = f"{round((tot_win / tot_ended) * 100, 1)}%" if tot_ended > 0 else "0%"
                rows_stat.append({
                    "م": len(rows_stat) + 1,
                    "نوع القضية": "الإجمالي العام",
                    "إجمالي القضايا": tot_all,
                    "المتداولة": tot_active,
                    "المنتهية بأحكام": tot_ended,
                    "أحكام لصالح الهيئة": tot_win,
                    "أحكام ضد الهيئة": tot_loss,
                    "نسبة الصالح": tot_win_pct,
                })

                عنوان_3_كامل = _build_report_heading_text(عنوان_3, محامي3)
                html_rep = report_header(منطقة3, عنوان_3_كامل, مدير_عام3, مدير_ادارة3, عضو_قانوني3, محامي3)
                html_rep += (
                    f"<div style='display:flex; flex-wrap:wrap; gap:8px; justify-content:center; margin:10px 0;'>"
                    f"<div style='background:#142038; border:1px solid #D4AF37; border-radius:10px; padding:8px 14px; color:#FFF;'><b>الإجمالي:</b> <span style='color:#FFD700'>{tot_all}</span></div>"
                    f"<div style='background:#142038; border:1px solid #D4AF37; border-radius:10px; padding:8px 14px; color:#FFF;'><b>المتداولة:</b> <span style='color:#4DA8DA'>{tot_active}</span></div>"
                    f"<div style='background:#142038; border:1px solid #D4AF37; border-radius:10px; padding:8px 14px; color:#FFF;'><b>المنتهية:</b> <span style='color:#D4AF37'>{tot_ended}</span></div>"
                    f"<div style='background:#142038; border:1px solid #4CAF50; border-radius:10px; padding:8px 14px; color:#FFF;'><b>الصالح:</b> <span style='color:#4CAF50'>{tot_win}</span></div>"
                    f"<div style='background:#142038; border:1px solid #FF5252; border-radius:10px; padding:8px 14px; color:#FFF;'><b>الضد:</b> <span style='color:#FF5252'>{tot_loss}</span></div>"
                    f"<div style='background:#142038; border:1px solid #D4AF37; border-radius:10px; padding:8px 14px; color:#FFF;'><b>نسبة الصالح:</b> <span style='color:#FFD700'>{tot_win_pct}</span></div>"
                    f"</div>"
                )
                html_rep += '<div class="table-container"><table class="case-table"><tr><th>م</th><th>نوع القضية</th><th>إجمالي القضايا</th><th>المتداولة</th><th>المنتهية بأحكام</th><th>أحكام لصالح الهيئة</th><th>أحكام ضد الهيئة</th><th>نسبة الصالح</th></tr>'
                for r in rows_stat:
                    is_total = r["نوع القضية"] == "الإجمالي العام"
                    row_style = " style='font-weight:900; background:#142038; color:#FFD700;'" if is_total else ""
                    html_rep += (
                        f"<tr{row_style}>"
                        f"<td>{r['م']}</td>"
                        f"<td>{html.escape(str(r['نوع القضية']))}</td>"
                        f"<td>{r['إجمالي القضايا']}</td>"
                        f"<td>{r['المتداولة']}</td>"
                        f"<td>{r['المنتهية بأحكام']}</td>"
                        f"<td style='color:#4CAF50; font-weight:900'>{r['أحكام لصالح الهيئة']}</td>"
                        f"<td style='color:#FF5252; font-weight:900'>{r['أحكام ضد الهيئة']}</td>"
                        f"<td style='color:#FFD700; font-weight:900'>{r['نسبة الصالح']}</td>"
                        f"</tr>"
                    )
                html_rep += '</table></div>'
                html_rep += report_footer(مدير_عام3, مدير_ادارة3, عضو_قانوني3)
                st.session_state.last_report_html = html_rep
                st.session_state.last_report_title = عنوان_3_كامل
                st.session_state.last_report_df = pd.DataFrame(rows_stat)
                st.session_state.last_report_meta = {"region": منطقة3, "member": عضو_قانوني3, "manager": مدير_ادارة3, "gm": مدير_عام3}
                st.markdown(html_rep, unsafe_allow_html=True)
            except Exception as exc:
                st.error(f"حدث خطأ أثناء تجهيز الإحصائيات: {exc}")

    # ------------------------------------------------------------------
    # التبويب الرابع: البيان العددي
    # ------------------------------------------------------------------
    with tab4:
        st.markdown("<h3 style='color:#B8860B; text-align:center; font-size:15px;'>📊 البيان العددي</h3>", unsafe_allow_html=True)
        with st.container(border=True):
            t4_chk_keys = ["t4_chk_type", "t4_chk_subject"]
            for k in t4_chk_keys:
                st.session_state.setdefault(k, k == "t4_chk_type")

            c_t4_1, c_t4_2 = st.columns(2)
            with c_t4_1:
                st.checkbox("بيان عددي شامل حسب نوع القضية", key="t4_chk_type", on_change=_clear_other_report_checkboxes, args=("t4_chk_type", t4_chk_keys))
            with c_t4_2:
                st.checkbox("بيان عددي حسب موضوع الدعوى", key="t4_chk_subject", on_change=_clear_other_report_checkboxes, args=("t4_chk_subject", t4_chk_keys))

            if st.session_state.get("t4_chk_subject"):
                نوع_البيان_العددي = "بيان عددي حسب موضوع الدعوى"
            else:
                st.session_state["t4_chk_type"] = True
                نوع_البيان_العددي = "بيان عددي شامل حسب نوع القضية"

            فلتر_موضوع_4 = ""
            if نوع_البيان_العددي == "بيان عددي حسب موضوع الدعوى":
                فلتر_موضوع_4 = st.text_input("تصفية بموضوع دعوى محدد (اتركه فارغاً لعرض كل الموضوعات)", placeholder="اكتب موضوع الدعوى أو اتركه فارغاً لتجميع كل الموضوعات...", key="t4_subject")

            تفعيل_تاريخ_4 = st.checkbox("📅 تصفية البيان العددي بفترة زمنية (من تاريخ / إلى تاريخ)", value=False, key="t4_use_date")
            من_تاريخ_4_str = ""
            الى_تاريخ_4_str = ""
            if تفعيل_تاريخ_4:
                c4_d1, c4_d2 = st.columns(2)
                with c4_d1:
                    من_تاريخ_4_str = date_input_dmy("من تاريخ (يوم / شهر / سنة)", value=datetime.now().date().replace(day=1), key="t4_from_date")
                with c4_d2:
                    الى_تاريخ_4_str = date_input_dmy("إلى تاريخ (يوم / شهر / سنة)", value=datetime.now().date(), key="t4_to_date")

            col4_r1, col4_r2 = st.columns(2)
            with col4_r1:
                منطقة4 = st.text_input("اسم المنطقة", placeholder="مثال: البحيرة", key="r4")
            with col4_r2:
                محامي4 = st.text_input("طرف الأستاذ / المحامي", placeholder="اكتب اسم المحامي...", key="l4")

            col1, col2, col3 = st.columns(3)
            with col1: مدير_عام4 = st.text_input("مدير عام الإدارات القانونية", key="gm4")
            with col2: مدير_ادارة4 = st.text_input("مدير إدارة القضايا", key="m4")
            with col3: عضو_قانوني4 = st.text_input("العضو القانوني", key="u4")

        if st.button("📊 عرض وتجهيز البيان العددي", key="b4", use_container_width=True):
            قضايا_عددي = list(my_cases)

            if تفعيل_تاريخ_4:
                d_from_4 = parse_date_safe(من_تاريخ_4_str)
                d_to_4 = parse_date_safe(الى_تاريخ_4_str)
                if d_from_4 and d_to_4 and d_from_4 > d_to_4:
                    d_from_4, d_to_4 = d_to_4, d_from_4
                قضايا_عددي = [
                    c for c in قضايا_عددي
                    if _date_in_range(
                        c.get('تاريخ_الحكم') if str(c.get('حالة', '')).strip() == 'منتهية' else c.get('تاريخ_جلسة'),
                        d_from_4,
                        d_to_4,
                        fallback_dates=_case_session_dates(c) + ([parse_date_safe(c.get('تاريخ_الحكم'))] if c.get('تاريخ_الحكم') else [])
                    )
                ]

            if نوع_البيان_العددي == "بيان عددي حسب موضوع الدعوى":
                if فلتر_موضوع_4.strip():
                    q_subj4 = _norm_ar(فلتر_موضوع_4)
                    قضايا_عددي = [c for c in قضايا_عددي if q_subj4 in _norm_ar(c.get('موضوع', ''))]
                    عنوان_4 = f"بيان عددي بالدعاوى والأحكام عن موضوع: {فلتر_موضوع_4.strip()}"
                else:
                    عنوان_4 = "بيان عددي بالدعاوى والأحكام حسب موضوع الدعوى"
            else:
                عنوان_4 = "بيان عددي شامل بالدعاوى والأحكام"

            if تفعيل_تاريخ_4:
                عنوان_4 += f" (خلال الفترة من {fmt_date_ar(d_from_4)} إلى {fmt_date_ar(d_to_4)})"

            if not قضايا_عددي:
                st.warning("لا توجد بيانات مطابقة للبحث")
            else:
                rows_num = []
                tot_all = tot_active = tot_ended = tot_win = tot_loss = 0

                if نوع_البيان_العددي == "بيان عددي حسب موضوع الدعوى":
                    grouped_subj = {}
                    for c in قضايا_عددي:
                        raw_subj = str(c.get('موضوع', '') or '').strip() or "غير محدد"
                        norm_k = _norm_ar(raw_subj) or "غير محدد"
                        if norm_k not in grouped_subj:
                            grouped_subj[norm_k] = {"label": raw_subj, "cases": []}
                        grouped_subj[norm_k]["cases"].append(c)

                    for idx_s, (_, item) in enumerate(sorted(grouped_subj.items(), key=lambda x: len(x[1]["cases"]), reverse=True), 1):
                        sub_cases = item["cases"]
                        cnt_all = len(sub_cases)
                        cnt_active = sum(1 for c in sub_cases if str(c.get('حالة', '')).strip() == 'متداولة')
                        ended_sub = [c for c in sub_cases if str(c.get('حالة', '')).strip() == 'منتهية']
                        cnt_ended = len(ended_sub)
                        cnt_win = sum(1 for c in ended_sub if str(c.get('مسندة_ل_الحكم', '')).strip() == 'الصالح')
                        cnt_loss = sum(1 for c in ended_sub if str(c.get('مسندة_ل_الحكم', '')).strip() == 'الضد')

                        tot_all += cnt_all
                        tot_active += cnt_active
                        tot_ended += cnt_ended
                        tot_win += cnt_win
                        tot_loss += cnt_loss

                        rows_num.append({
                            "م": idx_s,
                            "موضوع الدعوى": item["label"],
                            "العدد الإجمالي": cnt_all,
                            "المتداولة": cnt_active,
                            "المنتهية": cnt_ended,
                            "الصالح": cnt_win,
                            "الضد": cnt_loss,
                        })

                    rows_num.append({
                        "م": len(rows_num) + 1,
                        "موضوع الدعوى": "الإجمالي العام",
                        "العدد الإجمالي": tot_all,
                        "المتداولة": tot_active,
                        "المنتهية": tot_ended,
                        "الصالح": tot_win,
                        "الضد": tot_loss,
                    })
                    first_col_title = "موضوع الدعوى"
                else:
                    انواع_عددي = ["دعوى", "استئناف", "طعن"]
                    for idx_t, t_name in enumerate(انواع_عددي, 1):
                        sub_cases = [c for c in قضايا_عددي if str(c.get('نوع', '')).strip() == t_name]
                        cnt_all = len(sub_cases)
                        cnt_active = sum(1 for c in sub_cases if str(c.get('حالة', '')).strip() == 'متداولة')
                        ended_sub = [c for c in sub_cases if str(c.get('حالة', '')).strip() == 'منتهية']
                        cnt_ended = len(ended_sub)
                        cnt_win = sum(1 for c in ended_sub if str(c.get('مسندة_ل_الحكم', '')).strip() == 'الصالح')
                        cnt_loss = sum(1 for c in ended_sub if str(c.get('مسندة_ل_الحكم', '')).strip() == 'الضد')

                        tot_all += cnt_all
                        tot_active += cnt_active
                        tot_ended += cnt_ended
                        tot_win += cnt_win
                        tot_loss += cnt_loss

                        rows_num.append({
                            "م": idx_t,
                            "البيان": t_name,
                            "العدد الإجمالي": cnt_all,
                            "المتداولة": cnt_active,
                            "المنتهية": cnt_ended,
                            "الصالح": cnt_win,
                            "الضد": cnt_loss,
                        })

                    rows_num.append({
                        "م": len(rows_num) + 1,
                        "البيان": "الإجمالي العام",
                        "العدد الإجمالي": tot_all,
                        "المتداولة": tot_active,
                        "المنتهية": tot_ended,
                        "الصالح": tot_win,
                        "الضد": tot_loss,
                    })
                    first_col_title = "البيان"

                عنوان_4_كامل = _build_report_heading_text(عنوان_4, محامي4)
                html_rep = report_header(منطقة4, عنوان_4_كامل, مدير_عام4, مدير_ادارة4, عضو_قانوني4, محامي4)
                html_rep += f'<div class="table-container"><table class="case-table"><tr><th>م</th><th>{first_col_title}</th><th>العدد الإجمالي</th><th>المتداولة</th><th>المنتهية</th><th>الصالح</th><th>الضد</th></tr>'
                for r in rows_num:
                    label_val = r.get(first_col_title, "")
                    is_total = label_val == "الإجمالي العام"
                    row_style = " style='font-weight:900; background:#142038; color:#FFD700;'" if is_total else ""
                    html_rep += (
                        f"<tr{row_style}>"
                        f"<td>{r['م']}</td>"
                        f"<td>{html.escape(str(label_val))}</td>"
                        f"<td>{r['العدد الإجمالي']}</td>"
                        f"<td>{r['المتداولة']}</td>"
                        f"<td>{r['المنتهية']}</td>"
                        f"<td style='color:#4CAF50; font-weight:900'>{r['الصالح']}</td>"
                        f"<td style='color:#FF5252; font-weight:900'>{r['الضد']}</td>"
                        f"</tr>"
                    )
                html_rep += '</table></div>'
                html_rep += report_footer(مدير_عام4, مدير_ادارة4, عضو_قانوني4)
                st.session_state.last_report_html = html_rep
                st.session_state.last_report_title = عنوان_4_كامل
                st.session_state.last_report_df = pd.DataFrame(rows_num)
                st.session_state.last_report_meta = {"region": منطقة4, "member": عضو_قانوني4, "manager": مدير_ادارة4, "gm": مدير_عام4}
                st.markdown(html_rep, unsafe_allow_html=True)

    # ------------------------------------------------------------------
    # شريط تصدير التقرير المعروض (HTML / Word / Excel / PDF)
    # ------------------------------------------------------------------
    if st.session_state.get("last_report_html"):
        st.markdown("---")
        st.markdown("<h4 style='color:#B8860B; text-align:center;'>📥 تحميل وطباعة التقرير</h4>", unsafe_allow_html=True)
        meta = st.session_state.get("last_report_meta") or {}
        rep_title = st.session_state.get("last_report_title", "تقرير")
        rep_df = st.session_state.get("last_report_df")
        if rep_df is None or not isinstance(rep_df, pd.DataFrame):
            rep_df = pd.DataFrame()

        c_exp1, c_exp2, c_exp3, c_exp4 = st.columns(4)
        with c_exp1:
            export_html = get_export_html(st.session_state.last_report_html, rep_title)
            st.download_button(
                "🌐 تحميل HTML للطباعة",
                data=export_html.encode("utf-8"),
                file_name=report_download_name(rep_title, meta.get("member", ""), "", meta.get("region", ""), "html"),
                mime="text/html",
                use_container_width=True,
                key="dl_rep_html"
            )
        with c_exp2:
            word_bytes = to_word(
                rep_df, rep_title, meta.get("region", ""),
                member_name=meta.get("member", ""),
                manager_name=meta.get("manager", ""),
                general_manager=meta.get("gm", "")
            )
            st.download_button(
                "📝 تحميل Word",
                data=word_bytes,
                file_name=report_download_name(rep_title, meta.get("member", ""), "", meta.get("region", ""), "docx"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="dl_rep_word"
            )
        with c_exp3:
            excel_bytes = to_excel(rep_df, rep_title)
            st.download_button(
                "📊 تحميل Excel",
                data=excel_bytes,
                file_name=report_download_name(rep_title, meta.get("member", ""), "", meta.get("region", ""), "xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_rep_excel"
            )
        with c_exp4:
            pdf_bytes = to_pdf(
                rep_df, rep_title, meta.get("region", ""),
                member_name=meta.get("member", ""),
                manager_name=meta.get("manager", ""),
                general_manager=meta.get("gm", "")
            )
            st.download_button(
                "📄 تحميل PDF",
                data=pdf_bytes,
                file_name=report_download_name(rep_title, meta.get("member", ""), "", meta.get("region", ""), "pdf"),
                mime="application/pdf",
                use_container_width=True,
                key="dl_rep_pdf"
            )

elif st.session_state.page == "لوحة التحكم":
    admin_control_page()
elif st.session_state.page == "extract_member":
    extract_member_page()
elif st.session_state.page == "الرسائل":
    if st.button("⬅️ العودة للرئيسية", use_container_width=True, key="back_from_messages_page"):
        st.session_state.page = "الرئيسية"
        st.rerun()
    messaging_page()
elif st.session_state.page == "set_password":
    set_password_page()
elif st.session_state.page == "لوحة تحكمى":
    personal_control_page()
