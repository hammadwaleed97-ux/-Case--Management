import streamlit as st
import pandas as pd
import json, os, re, io, uuid, hashlib, secrets, smtplib, base64, urllib.request, urllib.error, zipfile, shutil
from datetime import datetime, timedelta
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from openpyxl.styles import Alignment, Font
import requests
from bs4 import BeautifulSoup
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement

def _fmt_date_dmy(val):
    if not val:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s[:10], fmt).strftime("%d/%m/%Y")
        except Exception:
            pass
    try:
        return datetime.fromisoformat(s).strftime("%d/%m/%Y")
    except Exception:
        return s

# --- إعداد التخزين المحلي المستمر (بدون حذف أو فقدان بيانات) ---
BASE_DIR = Path(__file__).resolve().parent
LOCAL_DATA_DIR = BASE_DIR / "local_data"
LOCAL_BACKUP_DIR = BASE_DIR / "local_backups"
LOCAL_UPLOADS_DIR = BASE_DIR / "local_uploads"
for _d in (LOCAL_DATA_DIR, LOCAL_BACKUP_DIR, LOCAL_UPLOADS_DIR):
    _d.mkdir(parents=True, exist_ok=True)

DATA_FILE = LOCAL_DATA_DIR / "cases_data.json"
USERS_FILE = LOCAL_DATA_DIR / "users_data.json"
MESSAGES_FILE = LOCAL_DATA_DIR / "messages_data.json"
BANNERS_FILE = LOCAL_DATA_DIR / "banners_data.json"
LIBRARY_FILE = LOCAL_DATA_DIR / "library_data.json"
ARCHIVE_FILE = LOCAL_DATA_DIR / "archive_data.json"
REPORTS_FILE = LOCAL_DATA_DIR / "reports_data.json"

LOCAL_KEY_MAP = {
    "cases": DATA_FILE,
    "users": USERS_FILE,
    "messages": MESSAGES_FILE,
    "banners": BANNERS_FILE,
    "library": LIBRARY_FILE,
    "archive": ARCHIVE_FILE,
    "reports": REPORTS_FILE,
}

# --- إعدادات السحابة (Supabase) مع الحفاظ الكامل على الحفظ المحلي ---
def _get_secret(key, default=""):
    try:
        if key in st.secrets:
            return str(st.secrets[key])
        for section in ("supabase", "connections", "app"):
            if section in st.secrets and key in st.secrets[section]:
                return str(st.secrets[section][key])
    except Exception:
        pass
    return os.environ.get(key, default)

SUPABASE_URL = _get_secret("SUPABASE_URL", "").strip().rstrip("/")
SUPABASE_KEY = (
    _get_secret("SUPABASE_SERVICE_ROLE_KEY", "").strip()
    or _get_secret("SUPABASE_KEY", "").strip()
    or _get_secret("SUPABASE_ANON_KEY", "").strip()
)
SUPABASE_TABLE = _get_secret("SUPABASE_TABLE", "app_state").strip() or "app_state"
SUPABASE_BUCKET = _get_secret("SUPABASE_BUCKET", "case-files").strip() or "case-files"
SUPABASE_ENABLED = bool(SUPABASE_URL and SUPABASE_KEY)

def _supabase_headers(prefer="return=representation"):
    return {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": prefer,
    }

def _supabase_storage_headers(content_type="application/octet-stream", upsert="true"):
    return {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": content_type,
        "x-upsert": upsert,
    }

def _safe_storage_filename(filename):
    base_name = os.path.basename(str(filename or "file.bin"))
    stem, ext = os.path.splitext(base_name)
    cleaned_stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._")
    if not cleaned_stem:
        cleaned_stem = f"file_{uuid.uuid4().hex[:8]}"
    cleaned_ext = re.sub(r"[^A-Za-z0-9.]+", "", ext)[:12]
    return f"{uuid.uuid4().hex[:8]}_{cleaned_stem}{cleaned_ext}"

def _save_local_upload_copy(file_bytes, filename):
    safe_name = _safe_storage_filename(filename)
    local_path = LOCAL_UPLOADS_DIR / safe_name
    with open(local_path, "wb") as f:
        f.write(file_bytes)
    return str(local_path)

def upload_bytes_to_supabase(file_bytes, filename, folder="attachments", content_type="application/octet-stream"):
    if isinstance(file_bytes, io.BytesIO):
        file_bytes = file_bytes.getvalue()
    local_saved_path = _save_local_upload_copy(file_bytes, filename)
    if not SUPABASE_ENABLED:
        return {"ok": True, "url": local_saved_path, "path": local_saved_path, "local_path": local_saved_path, "storage": "local"}
    safe_name = _safe_storage_filename(filename)
    clean_folder = re.sub(r"[^A-Za-z0-9/_-]+", "_", str(folder or "attachments")).strip("/")
    storage_path = f"{clean_folder}/{datetime.now().strftime('%Y/%m')}/{safe_name}" if clean_folder else f"{datetime.now().strftime('%Y/%m')}/{safe_name}"
    upload_url = f"{SUPABASE_URL}/storage/v1/object/{SUPABASE_BUCKET}/{storage_path}"
    try:
        resp = requests.post(
            upload_url,
            headers=_supabase_storage_headers(content_type=content_type, upsert="true"),
            data=file_bytes,
            timeout=30,
        )
        if resp.status_code in (200, 201):
            public_url = f"{SUPABASE_URL}/storage/v1/object/public/{SUPABASE_BUCKET}/{storage_path}"
            return {"ok": True, "url": public_url, "path": storage_path, "local_path": local_saved_path, "storage": "supabase"}
        return {"ok": True, "url": local_saved_path, "path": local_saved_path, "local_path": local_saved_path, "storage": "local", "error": f"HTTP {resp.status_code}: {resp.text}"}
    except Exception as e:
        return {"ok": True, "url": local_saved_path, "path": local_saved_path, "local_path": local_saved_path, "storage": "local", "error": str(e)}

def _load_from_supabase(key):
    if not SUPABASE_ENABLED:
        return None
    url = f"{SUPABASE_URL}/rest/v1/{SUPABASE_TABLE}"
    try:
        resp = requests.get(
            url,
            headers=_supabase_headers(),
            params={"key": f"eq.{key}", "select": "key,payload,updated_at"},
            timeout=15,
        )
        if resp.status_code == 200:
            rows = resp.json()
            if isinstance(rows, list) and rows:
                payload = rows[0].get("payload")
                if isinstance(payload, str):
                    return json.loads(payload)
                return payload
    except Exception:
        pass
    return None

def _save_to_supabase(key, payload):
    if not SUPABASE_ENABLED:
        return False
    url = f"{SUPABASE_URL}/rest/v1/{SUPABASE_TABLE}"
    body = [{
        "key": key,
        "payload": payload,
        "updated_at": datetime.now().astimezone().isoformat(),
    }]
    try:
        resp = requests.post(
            url,
            headers=_supabase_headers("resolution=merge-duplicates,return=minimal"),
            params={"on_conflict": "key"},
            json=body,
            timeout=15,
        )
        return resp.status_code in (200, 201, 204)
    except Exception:
        return False

def _create_local_backup(path: Path, raw_text: str):
    try:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        b_dir = LOCAL_BACKUP_DIR / path.stem
        b_dir.mkdir(parents=True, exist_ok=True)
        bak_file = b_dir / f"{path.name}.{stamp}.bak"
        bak_file.write_text(raw_text, encoding="utf-8")
        latest_bak = b_dir / f"{path.name}.bak"
        latest_bak.write_text(raw_text, encoding="utf-8")
        backups = sorted(b_dir.glob(f"{path.name}.*.bak"))
        for old in backups[:-30]:
            try:
                old.unlink()
            except Exception:
                pass
    except Exception:
        pass

def _write_local_json_atomic(path: Path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    raw_text = json.dumps(obj, ensure_ascii=False, indent=2)
    tmp_path = path.with_name(f".{path.name}.{os.getpid()}.{uuid.uuid4().hex[:6]}.tmp")
    tmp_path.write_text(raw_text, encoding="utf-8")
    os.replace(tmp_path, path)
    _create_local_backup(path, raw_text)

def _read_local_json_safe(path: Path):
    for candidate in (path, LOCAL_BACKUP_DIR / path.stem / f"{path.name}.bak"):
        if candidate.exists():
            try:
                return json.loads(candidate.read_text(encoding="utf-8"))
            except Exception:
                continue
    return None

def load_json_persistent(key, local_path: Path, default_factory):
    local_obj = _read_local_json_safe(local_path)
    cloud_obj = _load_from_supabase(key)
    if cloud_obj is not None:
        if local_obj is None:
            try:
                _write_local_json_atomic(local_path, cloud_obj)
            except Exception:
                pass
        return cloud_obj
    if local_obj is not None:
        if SUPABASE_ENABLED:
            _save_to_supabase(key, local_obj)
        return local_obj
    fresh = default_factory()
    try:
        _write_local_json_atomic(local_path, fresh)
    except Exception:
        pass
    if SUPABASE_ENABLED:
        _save_to_supabase(key, fresh)
    return fresh

def save_json_persistent(key, local_path: Path, obj):
    _write_local_json_atomic(local_path, obj)
    if SUPABASE_ENABLED:
        _save_to_supabase(key, obj)

def _default_cases_payload():
    return {"cases": [], "last_update": ""}

def load_data():
    d = load_json_persistent("cases", DATA_FILE, _default_cases_payload)
    if not isinstance(d, dict):
        d = _default_cases_payload()
    if "cases" not in d or not isinstance(d["cases"], list):
        d["cases"] = []
    d.setdefault("last_update", "")
    return d

def save_data(d):
    if not isinstance(d, dict):
        d = _default_cases_payload()
    d.setdefault("cases", [])
    d["last_update"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    save_json_persistent("cases", DATA_FILE, d)

# --- نظام المستخدمين والصلاحيات ---
ROLES = {
    "admin": "أدمن",
    "general_manager": "مدير عام",
    "department_manager": "مدير إدارة",
    "legal_member": "عضو قانوني",
    "lawyer": "محامي",
    "administrative": "إداري"
}

def _default_users_payload():
    return [
        {
            "id": "u_admin_1",
            "username": "admin",
            "password_hash": hashlib.sha256("123456".encode()).hexdigest(),
            "full_name": "مدير النظام الرئيسي",
            "role": "admin",
            "region": "البحيرة",
            "specialization": "إدارة عامة",
            "status": "active",
            "created_by": "system",
            "birth_date": "1980-01-01",
            "age_action": None
        }
    ]

def load_users():
    u = load_json_persistent("users", USERS_FILE, _default_users_payload)
    if not isinstance(u, list) or not u:
        u = _default_users_payload()
        save_json_persistent("users", USERS_FILE, u)
    return u

def save_users(users):
    save_json_persistent("users", USERS_FILE, users)

def hash_password(pw):
    return hashlib.sha256(str(pw).encode()).hexdigest()

def normalize_person_name(name):
    s = str(name or "").strip()
    s = re.sub(r"^(?:الأستاذ(?:ة)?|أ\.?|د\.?|المستشار(?:ة)?|المحامي(?:\s+الأستاذ)?)\s*/?\s*", "", s)
    s = re.sub(r"\s+المحامي(?:\s+الاستئنافي|\s+بالنقض)?$", "", s)
    s = re.sub(r"[أإآا]", "ا", s)
    s = re.sub(r"ى", "ي", s)
    s = re.sub(r"ة", "ه", s)
    s = re.sub(r"[\u064B-\u065F\u0670]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _user_matches_identifier(user, identifier):
    ident_norm = normalize_person_name(identifier)
    if not ident_norm:
        return False
    ident_compact = ident_norm.replace(" ", "")
    for field in ("full_name", "username", "name"):
        val = user.get(field)
        if not val:
            continue
        norm_val = normalize_person_name(val)
        if norm_val == ident_norm:
            return True
        compact_val = norm_val.replace(" ", "")
        if compact_val and compact_val == ident_compact:
            return True
    return False

def find_user_by_name(users, name):
    for u in users:
        if _user_matches_identifier(u, name):
            return u
    return None

def can_create_role(creator_role, target_role):
    if creator_role == "admin":
        return target_role in ["general_manager", "department_manager", "legal_member", "lawyer", "administrative"]
    if creator_role == "general_manager":
        return target_role in ["department_manager", "legal_member", "lawyer", "administrative"]
    if creator_role == "department_manager":
        return target_role in ["legal_member", "lawyer", "administrative"]
    return False

def get_creatable_roles(creator_role):
    return [r for r in ROLES.keys() if can_create_role(creator_role, r)]

def can_message(sender_role, receiver_role):
    if sender_role == "admin":
        return True
    if sender_role == "general_manager":
        return receiver_role in ["admin", "department_manager", "legal_member", "lawyer", "administrative"]
    if sender_role == "department_manager":
        return receiver_role in ["general_manager", "legal_member", "lawyer", "administrative"]
    if sender_role in ["legal_member", "lawyer", "administrative"]:
        return receiver_role in ["general_manager", "department_manager", "legal_member", "lawyer", "administrative"]
    return False

def can_view_case(user, case):
    role = user.get("role")
    uname = user.get("username")
    fname = user.get("full_name")
    if role in ["admin", "general_manager"]:
        return True
    if role == "department_manager":
        return (not user.get("region")) or (case.get("region") == user.get("region")) or (case.get("created_by") == uname)
    # باقي الأدوار: يرى القضايا المسندة إليه أو التي أنشأها
    assigned_list = case.get("assigned_to", [])
    if isinstance(assigned_list, str):
        assigned_list = [assigned_list]
    assigned_matches = any(
        _user_matches_identifier({"full_name": item, "username": item}, fname)
        or _user_matches_identifier({"full_name": item, "username": item}, uname)
        for item in assigned_list if item
    )
    return (case.get("created_by") == uname) or assigned_matches or (case.get("assigned_lawyer") in [uname, fname])

def can_edit_case(user, case):
    role = user.get("role")
    if role in ["admin", "general_manager", "department_manager"]:
        return can_view_case(user, case)
    if role in ["legal_member", "lawyer"]:
        return can_view_case(user, case)
    return False

def can_delete_case(user, case):
    role = user.get("role")
    if role in ["admin", "general_manager"]:
        return True
    if role == "department_manager":
        return can_view_case(user, case)
    return False

def can_manage_users(user):
    return user.get("role") in ["admin", "general_manager", "department_manager"]

def can_access_archive(user):
    return user.get("role") in ["admin", "general_manager", "department_manager", "legal_member", "lawyer", "administrative"]

def can_access_reports(user):
    return user.get("role") in ["admin", "general_manager", "department_manager", "legal_member", "lawyer"]

def is_admin_like(user):
    return user.get("role") in ["admin", "general_manager", "department_manager"]

# تهيئة حالة الجلسة للمستخدم
if "auth_user" not in st.session_state:
    st.session_state.auth_user = None
if "page" not in st.session_state:
    st.session_state.page = "تسجيل"
if "selected_case_id" not in st.session_state:
    st.session_state.selected_case_id = None
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "الرئيسية"
if "show_notifications_panel" not in st.session_state:
    st.session_state.show_notifications_panel = False

def _refresh_auth_user():
    u = st.session_state.get("auth_user")
    if not u:
        return
    users = load_users()
    for item in users:
        if item.get("id") == u.get("id") or item.get("username") == u.get("username"):
            st.session_state.auth_user = item
            return

# شاشة تسجيل الدخول إذا لم يكن المستخدم مسجلاً
def render_login_screen():
    st.markdown("""
    <style>
    .login-box {
        max-width: 460px;
        margin: 40px auto;
        padding: 30px;
        border-radius: 18px;
        background: linear-gradient(145deg, #132238, #1d3557);
        border: 2px solid #d4af37;
        box-shadow: 0 12px 30px rgba(0,0,0,0.35);
        color: #fff;
        direction: rtl;
        text-align: right;
    }
    .login-title {
        text-align: center;
        color: #f4d03f;
        font-size: 26px;
        font-weight: 800;
        margin-bottom: 8px;
    }
    .login-sub {
        text-align: center;
        color: #d6e4f0;
        font-size: 14px;
        margin-bottom: 22px;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="login-box"><div class="login-title">⚖️ منظومة إدارة القضايا</div><div class="login-sub">تسجيل الدخول إلى النظام القانوني</div>', unsafe_allow_html=True)
    with st.form("login_form"):
        username = st.text_input("اسم المستخدم", placeholder="أدخل اسم المستخدم")
        password = st.text_input("كلمة المرور", type="password", placeholder="أدخل كلمة المرور")
        submitted = st.form_submit_button("دخول", use_container_width=True)
        if submitted:
            users = load_users()
            pw_hash = hash_password(password)
            matched = None
            for u in users:
                if u.get("username", "").strip().lower() == username.strip().lower() and u.get("password_hash") == pw_hash:
                    matched = u
                    break
            if not matched:
                st.error("اسم المستخدم أو كلمة المرور غير صحيحة")
            elif matched.get("status") != "active":
                st.error("هذا الحساب موقوف أو غير نشط، يرجى مراجعة الإدارة")
            else:
                st.session_state.auth_user = matched
                st.session_state.page = "تسجيل"
                st.success(f"مرحباً بك {matched.get('full_name')}")
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

def calculate_age(birth_date, on_date=None):
    if not birth_date:
        return None
    if isinstance(birth_date, str):
        try:
            birth_date = datetime.strptime(birth_date, "%Y-%m-%d").date()
        except Exception:
            return None
    on_date = on_date or datetime.now().date()
    return on_date.year - birth_date.year - ((on_date.month, on_date.day) < (birth_date.month, birth_date.day))

def get_age_status(user):
    """يعيد حالة السن، مع مراعاة تاريخ مد العضوية إن وجد."""
    birth = user.get("birth_date")
    age = calculate_age(birth)
    if age is None:
        return ("unknown", None)
    extension = user.get("membership_extended_until")
    extended_active = False
    if extension:
        try:
            extended_active = datetime.fromisoformat(str(extension)).date() >= datetime.now().date()
        except Exception:
            try:
                extended_active = datetime.strptime(str(extension), "%Y-%m-%d").date() >= datetime.now().date()
            except Exception:
                extended_active = False
    if extended_active:
        return ("extended", age)
    if age >= 60:
        return ("over_60", age)
    return ("normal", age)

def render_user_management():
    st.markdown("## 👥 إدارة الأعضاء والصلاحيات")
    if st.button("⬅️ العودة للوحة التحكم", use_container_width=True):
        st.session_state.page = "لوحة التحكم"
        st.rerun()
    current_user = st.session_state.auth_user
    users = load_users()
    creatable = get_creatable_roles(current_user.get("role"))

    st.markdown("### ➕ إضافة عضو جديد")
    if not creatable:
        st.info("صلاحياتك الحالية لا تسمح بإضافة أعضاء جدد.")
    else:
        with st.form("create_user_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                full_name = st.text_input("الاسم الكامل")
                username = st.text_input("اسم المستخدم (للدخول)")
                password = st.text_input("كلمة المرور", type="password")
                birth_date = st.date_input(
                    "تاريخ الميلاد",
                    value=datetime.now().date().replace(year=max(1900, datetime.now().year - 30)),
                    min_value=datetime(1900,1,1).date(),
                    max_value=datetime.now().date(),
                    format="DD/MM/YYYY",
                    key="new_member_birth_date"
                )
            with c2:
                role = st.selectbox("الدور الوظيفي", options=creatable, format_func=lambda r: ROLES.get(r, r))
                region = st.text_input("المنطقة / الفرع", value=current_user.get("region", "البحيرة"))
                specialization = st.text_input("التخصص / الصفة", value="عضو إدارة قانونية")
            submitted = st.form_submit_button("إنشاء الحساب", use_container_width=True)
            if submitted:
                if not full_name.strip() or not username.strip() or not password.strip():
                    st.error("يرجى استكمال الاسم الكامل واسم المستخدم وكلمة المرور")
                elif not can_create_role(current_user.get("role"), role):
                    st.error("غير مسموح لك بإنشاء هذا الدور الوظيفي")
                else:
                    existing_user = next((u for u in users if u.get("username", "").strip().lower() == username.strip().lower()), None)
                    if existing_user:
                        existing_user["full_name"] = full_name.strip()
                        existing_user["password_hash"] = hash_password(password)
                        existing_user["role"] = role
                        existing_user["region"] = region.strip()
                        existing_user["specialization"] = specialization.strip()
                        existing_user["status"] = "active"
                        existing_user["birth_date"] = birth_date.isoformat()
                        existing_user["updated_by"] = current_user.get("username")
                        save_users(users)
                        st.success("اسم المستخدم كان موجوداً مسبقاً؛ تم تحديث بيانات الحساب وتنشيطه بنجاح")
                        st.rerun()
                    else:
                        new_u = {
                            "id": f"u_{uuid.uuid4().hex[:8]}",
                            "username": username.strip(),
                            "password_hash": hash_password(password),
                            "full_name": full_name.strip(),
                            "role": role,
                            "region": region.strip(),
                            "specialization": specialization.strip(),
                            "status": "active",
                            "created_by": current_user.get("username"),
                            "birth_date": birth_date.isoformat(),
                            "age_action": None
                        }
                        users.append(new_u)
                        save_users(users)
                        st.success("تم إنشاء حساب العضو بنجاح")
                        st.rerun()

    st.markdown("### 📋 قائمة الأعضاء")
    for idx, u in enumerate(users):
        if current_user.get("role") == "department_manager" and u.get("role") in ["admin", "general_manager"]:
            continue
        status_label = "نشط ✅" if u.get("status") == "active" else "موقوف ⛔"
        age_status, age_val = get_age_status(u)
        age_badge = ""
        if age_status == "over_60":
            age_badge = f" | ⚠️ تجاوز 60 سنة ({age_val} سنة)"
        elif age_status == "extended":
            age_badge = f" | ⏳ عضوية ممدودة ({age_val} سنة - حتى {_fmt_date_dmy(u.get('membership_extended_until'))})"
        elif age_val is not None:
            age_badge = f" | السن: {age_val} سنة"

        with st.expander(f"{u.get('full_name')} (@{u.get('username')}) — {ROLES.get(u.get('role'), u.get('role'))} — {status_label}{age_badge}"):
            st.write(f"المنطقة: {u.get('region', '-')} | التخصص: {u.get('specialization', '-')}")
            st.write(f"تاريخ الميلاد: {_fmt_date_dmy(u.get('birth_date')) or 'غير مسجل'}")
            if age_status == "over_60":
                st.warning(f"تنبيه: هذا العضو بلغ من العمر {age_val} سنة (تجاوز سن الستين).")
            elif age_status == "extended":
                st.info(f"تم مد عضوية هذا العضو حتى {_fmt_date_dmy(u.get('membership_extended_until'))}.")

            col_a, col_b, col_c = st.columns(3)
            with col_a:
                if u.get("username") != current_user.get("username"):
                    if u.get("status") == "active":
                        if st.button("إيقاف الحساب", key=f"stop_{u['id']}_{idx}"):
                            u["status"] = "suspended"
                            save_users(users)
                            st.rerun()
                    else:
                        if st.button("تفعيل الحساب", key=f"act_{u['id']}_{idx}"):
                            u["status"] = "active"
                            save_users(users)
                            st.rerun()
            with col_b:
                if age_status == "over_60" and is_admin_like(current_user):
                    if st.button("إنهاء العضوية لبلوغ السن", key=f"end_age_{u['id']}_{idx}"):
                        u["status"] = "suspended"
                        u["age_action"] = "terminated_at_60"
                        u["age_action_at"] = datetime.now().isoformat()
                        save_users(users)
                        st.success("تم إنهاء عضوية الحساب لبلوغ سن الستين")
                        st.rerun()
            with col_c:
                if is_admin_like(current_user):
                    new_pw = st.text_input("كلمة مرور جديدة", type="password", key=f"npw_{u['id']}_{idx}")
                    if st.button("تغيير كلمة المرور", key=f"chpw_{u['id']}_{idx}"):
                        if new_pw.strip():
                            u["password_hash"] = hash_password(new_pw.strip())
                            save_users(users)
                            st.success("تم تحديث كلمة المرور")

            if is_admin_like(current_user):
                with st.expander("⏳ مد العضوية لفترة محددة", expanded=False):
                    extension_date = st.date_input(
                        "مد العضوية حتى تاريخ",
                        min_value=datetime.now().date(),
                        value=(datetime.now() + timedelta(days=365)).date(),
                        format="DD/MM/YYYY",
                        key=f"ext_date_{u['id']}_{idx}"
                    )
                    if st.button("تأكيد مد العضوية", key=f"confirm_ext_{u['id']}_{idx}"):
                        u["status"] = "active"
                        u["membership_extended_until"] = extension_date.isoformat()
                        u["age_action"] = "extended"
                        u["age_action_at"] = datetime.now().isoformat()
                        save_users(users)
                        st.success(f"تم مد العضوية حتى {_fmt_date_dmy(extension_date)}")
                        st.rerun()

# --- نظام الرسائل الداخلية والبنرات ---
def load_messages():
    m = load_json_persistent("messages", MESSAGES_FILE, list)
    return m if isinstance(m, list) else []

def save_messages(msgs):
    save_json_persistent("messages", MESSAGES_FILE, msgs)

def load_banners():
    b = load_json_persistent("banners", BANNERS_FILE, list)
    return b if isinstance(b, list) else []

def save_banners(banners):
    save_json_persistent("banners", BANNERS_FILE, banners)

def get_user_notifications(user):
    role = user.get("role")
    uname = user.get("username")
    fname = user.get("full_name")
    msgs = load_messages()
    my_msgs = []
    for m in msgs:
        if m.get("to_user") in [uname, fname, "ALL"] or (m.get("to_role") and m.get("to_role") == role):
            my_msgs.append(m)
    unread_msgs = [m for m in my_msgs if uname not in m.get("read_by", [])]
    return my_msgs, unread_msgs

def render_messages_page():
    st.markdown("## ✉️ الرسائل الداخلية")
    if st.button("⬅️ العودة للوحة التحكم", use_container_width=True):
        st.session_state.page = "لوحة التحكم"
        st.rerun()
    user = st.session_state.auth_user
    users = load_users()
    role = user.get("role")
    uname = user.get("username")

    st.markdown("### 📤 إرسال رسالة جديدة")
    allowed_roles = [r for r in ROLES.keys() if can_message(role, r)]
    recipient_users = [u for u in users if u.get("role") in allowed_roles and u.get("username") != uname and u.get("status") == "active"]

    with st.form("send_msg_form", clear_on_submit=True):
        target_mode = st.radio("نوع المستلم", ["عضو محدد", "مجموعة / دور وظيفي"], horizontal=True)
        to_user = "ALL"
        to_role = None
        if target_mode == "عضو محدد":
            if recipient_users:
                selected_u = st.selectbox(
                    "اختر العضو",
                    options=recipient_users,
                    format_func=lambda x: f"{x.get('full_name')} ({ROLES.get(x.get('role'))})"
                )
                to_user = selected_u.get("username")
            else:
                st.info("لا يوجد أعضاء متاحون للإرسال لهم")
        else:
            if allowed_roles:
                to_role = st.selectbox("اختر الدور الوظيفي", options=allowed_roles, format_func=lambda r: ROLES.get(r, r))
                to_user = "ROLE_GROUP"
        subject = st.text_input("عنوان الرسالة")
        body = st.text_area("نص الرسالة")
        send_btn = st.form_submit_button("إرسال الرسالة", use_container_width=True)
        if send_btn:
            if not body.strip():
                st.error("يرجى كتابة نص الرسالة")
            else:
                msgs = load_messages()
                msgs.insert(0, {
                    "id": f"msg_{uuid.uuid4().hex[:8]}",
                    "from_user": uname,
                    "from_name": user.get("full_name"),
                    "from_role": role,
                    "to_user": to_user,
                    "to_role": to_role,
                    "subject": subject.strip() or "رسالة إدارية",
                    "body": body.strip(),
                    "created_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "read_by": [uname]
                })
                save_messages(msgs)
                st.success("تم إرسال الرسالة بنجاح")
                st.rerun()
                if age is not None:
                    age_line = f"**العمر:** {age} سنة"
                    if extended_active:
                        age_line += f" — **العضوية ممددة حتى:** {_fmt_date_dmy(user.get('membership_extended_until'))}"
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
                    if st.button("✏️ تعديل", key=f"edit_u_{user['id']}", use_container_width=True):
                        st.session_state[f"editing_user_{user['id']}"] = True
                with b:
                    if user.get("username") != st.session_state.auth_user.get("username"):
                        if st.button("🚫 إيقاف", key=f"del_u_{user['id']}", use_container_width=True):
                            user["status"] = "suspended_manual"
                            save_users(users)
                            st.rerun()
                with c:
                    if st.button("🔑 كلمة مرور", key=f"reset_pw_{user['id']}", use_container_width=True):
                        st.session_state[f"resetting_pw_{user['id']}"] = True

            if st.session_state.get(f"editing_user_{user['id']}"):
                with st.form(f"edit_user_form_{user['id']}"):
                    efull = st.text_input("الاسم الكامل", value=user.get("full_name", ""))
                    erole = st.selectbox(
                        "الدور",
                        options=list(ROLES.keys()),
                        index=list(ROLES.keys()).index(user.get("role", "lawyer")) if user.get("role") in ROLES else 4,
                        format_func=lambda r: ROLES.get(r, r)
                    )
                    ereg = st.text_input("المنطقة", value=user.get("region", ""))
                    espec = st.text_input("التخصص", value=user.get("specialization", ""))
                    ebirth_val = None
                    if user.get("birth_date"):
                        try:
                            ebirth_val = datetime.strptime(str(user.get("birth_date"))[:10], "%Y-%m-%d").date()
                        except Exception:
                            ebirth_val = None
                    ebirth = st.date_input(
                        "تاريخ الميلاد",
                        value=ebirth_val or datetime(1980, 1, 1).date(),
                        min_value=datetime(1900, 1, 1).date(),
                        max_value=datetime.now().date(),
                        format="DD/MM/YYYY",
                        key=f"edit_birth_{user['id']}"
                    )
                    if st.form_submit_button("💾 حفظ التعديلات", use_container_width=True):
                        user["full_name"] = efull.strip()
                        user["role"] = erole
                        user["region"] = ereg.strip()
                        user["specialization"] = espec.strip()
                        user["birth_date"] = ebirth.isoformat()
                        save_users(users)
                        st.session_state[f"editing_user_{user['id']}"] = False
                        st.rerun()

            if st.session_state.get(f"resetting_pw_{user['id']}"):
                with st.form(f"reset_pw_form_{user['id']}"):
                    npw = st.text_input("كلمة المرور الجديدة", type="password")
                    if st.form_submit_button("🔒 تحديث كلمة المرور", use_container_width=True):
                        if npw.strip():
                            user["password_hash"] = hash_password(npw.strip())
                            save_users(users)
                            st.session_state[f"resetting_pw_{user['id']}"] = False
                            st.success("تم تحديث كلمة المرور بنجاح")

def render_user_management():
    _refresh_auth_user()
    u = st.session_state.auth_user
    if not u or not can_manage_users(u):
        st.error("⛔ لا تملك صلاحية الوصول إلى إدارة الأعضاء.")
        return

    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>👥 إدارة الأعضاء والصلاحيات</h1>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للوحة التحكم", use_container_width=True):
        st.session_state.page = "لوحة التحكم"
        st.rerun()

    users = load_users()
    creatable_roles = get_creatable_roles(u.get("role"))

    st.markdown("### ➕ إضافة عضو جديد")
    if not creatable_roles:
        st.info("لا توجد أدوار متاحة للإنشاء بصلاحياتك الحالية.")
    else:
        with st.form("add_new_user_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                full_name = st.text_input("الاسم الكامل *")
                username = st.text_input("اسم المستخدم *")
                password = st.text_input("كلمة المرور *", type="password")
            with c2:
                role = st.selectbox("الدور الوظيفي", options=creatable_roles, format_func=lambda r: ROLES.get(r, r))
                region = st.text_input("المنطقة", value=u.get("region", "البحيرة"))
                specialization = st.text_input("التخصص / الصفة", value="محامي")
            with c3:
                email = st.text_input("البريد الإلكتروني (اختياري)")
                birth_date = st.date_input(
                    "تاريخ الميلاد",
                    value=datetime(1980, 1, 1).date(),
                    min_value=datetime(1900, 1, 1).date(),
                    max_value=datetime.now().date(),
                    format="DD/MM/YYYY",
                    key="add_user_birth_date"
                )

            if st.form_submit_button("✅ إضافة العضو", use_container_width=True):
                if not full_name.strip() or not username.strip() or not password.strip():
                    st.error("يرجى إدخال الاسم الكامل واسم المستخدم وكلمة المرور.")
                elif any(x.get("username", "").lower() == username.strip().lower() for x in users):
                    st.error("اسم المستخدم موجود بالفعل.")
                else:
                    new_member = {
                        "id": f"u_{uuid.uuid4().hex[:8]}",
                        "username": username.strip(),
                        "password_hash": hash_password(password.strip()),
                        "full_name": full_name.strip(),
                        "role": role,
                        "region": region.strip(),
                        "specialization": specialization.strip(),
                        "email": email.strip(),
                        "birth_date": birth_date.isoformat(),
                        "status": "active",
                        "created_by": u.get("username"),
                        "created_at": datetime.now().isoformat(),
                    }
                    users.append(new_member)
                    save_users(users)
                    st.success(f"تم إضافة العضو {full_name.strip()} بنجاح.")
                    st.rerun()

    st.markdown("---")
    st.markdown("### 📋 قائمة الأعضاء")
    for member in users:
        if u.get("role") == "department_manager" and member.get("role") in ("admin", "general_manager"):
            continue
        age_status, age_val = get_age_status(member)
        status = member.get("status", "active")
        badge = "🟢 نشط" if status == "active" else f"🔴 موقوف ({status})"
        if age_status == "over_60":
            badge += f" | ⚠️ سن {age_val}"
        elif age_status == "extended":
            badge += f" | ⏳ ممدد حتى {_fmt_date_dmy(member.get('membership_extended_until'))}"

        with st.expander(f"{member.get('full_name')} (@{member.get('username')}) — {ROLES.get(member.get('role'), member.get('role'))} — {badge}"):
            st.write(f"**المنطقة:** {member.get('region', '-')} | **التخصص:** {member.get('specialization', '-')} | **تاريخ الميلاد:** {_fmt_date_dmy(member.get('birth_date')) or '-'}")
            if age_status == "over_60":
                st.warning(f"العضو بلغ {age_val} سنة (تجاوز سن الستين).")
                if st.button("⚖️ اتخاذ إجراء بلوغ سن الستين", key=f"mgmt_age_{member['id']}"):
                    _age_action_dialog(member["id"])
            if status != "active":
                if st.button("✅ تفعيل العضو", key=f"mgmt_act_{member['id']}"):
                    member["status"] = "active"
                    save_users(users)
                    st.rerun()
            elif member.get("username") != u.get("username"):
                if st.button("🚫 إيقاف العضو", key=f"mgmt_sus_{member['id']}"):
                    member["status"] = "suspended_manual"
                    save_users(users)
                    st.rerun()

def _recipient_groups(candidates, key_prefix, preselected=None):
    preselected = set(preselected or [])
    selected_ids = []
    cols = st.columns(3)
    for idx, cand in enumerate(candidates):
        cid = cand.get("id")
        label = f"{cand.get('full_name')} ({ROLES.get(cand.get('role'), cand.get('role'))})"
        checked = st.checkbox(label, value=(cid in preselected), key=f"{key_prefix}_{cid}_{idx}")
        if checked:
            selected_ids.append(cid)
    return selected_ids

def render_messages_page():
    _refresh_auth_user()
    user = st.session_state.auth_user
    if not user:
        return
    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>✉️ الرسائل الداخلية</h1>", unsafe_allow_html=True)
    if st.button("⬅️ العودة للوحة التحكم", use_container_width=True):
        st.session_state.page = "لوحة التحكم"
        st.rerun()

    users = load_users()
    msgs = load_messages()
    uname = user.get("username")
    fname = user.get("full_name")
    role = user.get("role")

    tab1, tab2 = st.tabs(["📥 الوارد", "📤 إرسال رسالة"])
    with tab1:
        inbox = [m for m in msgs if m.get("to_user") in (uname, fname, "ALL") or (m.get("to_role") and m.get("to_role") == role)]
        if not inbox:
            st.info("لا توجد رسائل واردة.")
        for m in inbox:
            is_unread = uname not in m.get("read_by", [])
            title_prefix = "🆕 " if is_unread else "📩 "
            with st.expander(f"{title_prefix}{m.get('subject', 'رسالة')} — من: {m.get('from_name', m.get('from_user'))} ({_fmt_date_dmy(m.get('created_at'))})"):
                st.write(m.get("body", ""))
                if is_unread:
                    if st.button("تحديد كمقروءة", key=f"read_m_{m.get('id')}"):
                        m.setdefault("read_by", []).append(uname)
                        save_messages(msgs)
                        st.rerun()

    with tab2:
        candidates = [u for u in users if u.get("status") == "active" and u.get("id") != user.get("id") and can_message(role, u.get("role"))]
        if not candidates:
            st.info("لا يوجد مستلمون متاحون.")
        else:
            with st.form("compose_internal_msg", clear_on_submit=True):
                mode = st.radio("إرسال إلى", ["أعضاء محددين", "كل المتاحين"], horizontal=True)
                if mode == "كل المتاحين":
                    selected_ids = [u.get("id") for u in candidates]
                else:
                    selected_ids = _recipient_groups(candidates, f"msg_chk_{role}")
                subj = st.text_input("العنوان")
                body = st.text_area("نص الرسالة")
                if st.form_submit_button("📨 إرسال", use_container_width=True):
                    if not body.strip():
                        st.error("يرجى كتابة نص الرسالة.")
                    elif not selected_ids:
                        st.error("يرجى اختيار مستلم واحد على الأقل.")
                    else:
                        chosen = [u for u in candidates if u.get("id") in selected_ids]
                        for ch in chosen:
                            msgs.insert(0, {
                                "id": f"msg_{uuid.uuid4().hex[:8]}",
                                "from_user": uname,
                                "from_name": fname,
                                "from_role": role,
                                "to_user": ch.get("username"),
                                "to_role": ch.get("role"),
                                "subject": subj.strip() or "رسالة إدارية",
                                "body": body.strip(),
                                "created_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
                                "read_by": [uname]
                            })
                        save_messages(msgs)
                        st.success("تم إرسال الرسالة.")
                        st.rerun()

def show_banners(location="top"):
    banners = load_banners()
    if not banners:
        return
    user = st.session_state.get("auth_user") or {}
    uname = user.get("username", "")
    role = user.get("role", "")
    now = datetime.now()
    for b in banners:
        if b.get("active") is False:
            continue
        if b.get("location") and b.get("location") not in ("all", location):
            continue
        vis = b.get("visible_to")
        if vis and uname not in vis and role not in vis and "ALL" not in vis:
            continue
        exp = b.get("expire_at") or b.get("expire")
        if exp:
            try:
                if datetime.fromisoformat(str(exp)) < now:
                    continue
            except Exception:
                pass
        btype = b.get("type", "info")
        msg = b.get("text") or b.get("message") or ""
        if not msg:
            continue
        if btype == "warning":
            st.warning(msg)
        elif btype == "error":
            st.error(msg)
        elif btype == "success":
            st.success(msg)
        else:
            st.info(msg)

def render_admin_dashboard():
    _refresh_auth_user()
    user = st.session_state.auth_user
    if not user or not is_admin_like(user):
        st.error("⛔ لا تملك صلاحية الوصول إلى لوحة التحكم.")
        return

    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>⚖️ لوحة التحكم</h1>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("👥 إدارة الأعضاء والصلاحيات", use_container_width=True):
            st.session_state.page = "إدارة الأعضاء"
            st.rerun()
    with col2:
        if st.button("✉️ الرسائل الداخلية", use_container_width=True):
            st.session_state.page = "الرسائل"
            st.rerun()
    with col3:
        if st.button("📝 تسجيل قضية جديدة", use_container_width=True):
            st.session_state.page = "تسجيل"
            st.rerun()

    st.markdown("---")
    st.markdown("### 📢 إدارة البنرات والإعلانات")
    banners = load_banners()
    users = load_users()
    with st.form("create_banner_form", clear_on_submit=True):
        b_text = st.text_input("نص البنر")
        c1, c2 = st.columns(2)
        with c1:
            b_type = st.selectbox("النوع", ["info", "warning", "success", "error"])
            b_days = st.number_input("مدة الظهور بالأيام", min_value=1, max_value=365, value=7)
        with c2:
            banner_candidates = [u for u in users if u.get("status") == "active"]
            selected_banner_ids = _recipient_groups(banner_candidates, "banner_recipient_chk_v3")
            visible_to = [u.get("username") for u in banner_candidates if u.get("id") in selected_banner_ids]
        if st.form_submit_button("➕ نشر البنر", use_container_width=True):
            if b_text.strip():
                created = datetime.now()
                expire = created + timedelta(days=int(b_days))
                banners.insert(0, {
                    "id": f"bn_{uuid.uuid4().hex[:8]}",
                    "text": b_text.strip(),
                    "type": b_type,
                    "created_at": created.isoformat(),
                    "expire_at": expire.isoformat(),
                    "visible_to": visible_to or ["ALL"],
                    "active": True,
                    "location": "all"
                })
                save_banners(banners)
                st.success("تم نشر البنر.")
                st.rerun()

    for i, b in enumerate(banners):
        exp_str = _fmt_date_dmy(b.get("expire_at") or b.get("expire"))
        with st.expander(f"بنر: {b.get('text', '')[:40]} — ينتهي: {exp_str or '-'}"):
            expire_raw = b.get("expire_at") or b.get("expire")
            expire_dt = None
            if expire_raw:
                try:
                    expire_dt = datetime.fromisoformat(str(expire_raw))
                except Exception:
                    expire_dt = None
            new_expire = st.date_input(
                "تاريخ الانتهاء",
                value=expire_dt.date() if expire_dt else datetime.now().date(),
                format="DD/MM/YYYY",
                key=f"ed_{b.get('id', i)}"
            )
            c_s1, c_s2 = st.columns(2)
            with c_s1:
                if st.button("💾 حفظ تاريخ الانتهاء", key=f"save_exp_{b.get('id', i)}"):
                    old_dt = expire_dt or datetime.now()
                    b["expire_at"] = datetime.combine(new_expire, old_dt.time()).isoformat()
                    save_banners(banners)
                    st.success("تم التحديث")
                    st.rerun()
            with c_s2:
                if st.button("🗑️ حذف البنر", key=f"del_bn_{b.get('id', i)}"):
                    banners.pop(i)
                    save_banners(banners)
                    st.rerun()

def render_member_dashboard():
    _refresh_auth_user()
    user = st.session_state.auth_user
    if not user:
        return
    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>⚙️ لوحة تحكمى</h1>", unsafe_allow_html=True)
    st.write(f"**الاسم:** {user.get('full_name')} | **الصفة:** {ROLES.get(user.get('role'), user.get('role'))}")
    st.write(f"**تاريخ الميلاد:** {_fmt_date_dmy(user.get('birth_date')) or 'غير مسجل'}")
    if user.get("membership_extended_until"):
        st.info(f"العضوية ممددة حتى: {_fmt_date_dmy(user.get('membership_extended_until'))}")
    if st.button("✉️ الرسائل الداخلية", use_container_width=True):
        st.session_state.page = "الرسائل"
        st.rerun()

import arabic_reshaper
from bidi.algorithm import get_display
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from fpdf import FPDF

st.markdown('\n<style id="visible-controls-v4">\nbutton,.stButton button,.stDownloadButton button,\ndiv[data-testid="stFileUploader"] button,\ndiv[data-testid="stTabs"] button,\ndiv[data-testid="stExpander"] button {\n    opacity:1!important;\n    visibility:visible!important;\n}\nbutton p,button span,.stButton button p,.stButton button span,\n.stDownloadButton button p,.stDownloadButton button span {\n    opacity:1!important;\n    visibility:visible!important;\n    color:#fff!important;\n}\n</style>\n', unsafe_allow_html=True)
st.set_page_config(page_title="إدارة القضايا", layout="wide", page_icon="⚖️")

# ===== تثبيت ظهور أسماء الأزرار والأيقونات بوضوح =====
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
div[data-testid="stTabs"] button p {
    color:#ffffff !important;
    font-weight:900 !important;
}
div[data-testid="stTabs"] button[aria-selected="true"],
div[data-testid="stTabs"] button[aria-selected="true"] p {
    color:#D4AF37 !important;
}
div[data-testid="stTextInput"] label,
div[data-testid="stPasswordInput"] label,
div[data-testid="stDateInput"] label,
div[data-testid="stSelectbox"] label,
div[data-testid="stTextArea"] label {
    color:#ffffff !important;
    font-weight:900 !important;
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

def to_excel(df, title="التقرير"):
    df = df.fillna("").astype(str)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="التقرير")
        ws = writer.sheets["التقرير"]
        ws.sheet_view.rightToLeft = True
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for col in ws.columns:
            letter = col[0].column_letter
            length = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[letter].width = min(max(length + 3, 14), 45)
    return output.getvalue()

def to_word(df, title, region, member_name="", specialization="", manager_name="", general_manager=""):
    from docx.oxml.ns import qn
    doc = Document()
    for sec in doc.sections:
        sec.orientation = 1  # Landscape
        sec.page_width, sec.page_height = sec.page_height, sec.page_width
        sec.top_margin = Inches(0.6)
        sec.bottom_margin = Inches(0.6)
        sec.left_margin = Inches(0.6)
        sec.right_margin = Inches(0.6)

    def rtl(text, size=11, bold=False, align=WD_ALIGN_PARAGRAPH.RIGHT):
        p = doc.add_paragraph()
        p.alignment = align
        pPr = p._p.get_or_add_pPr()
        pPr.append(OxmlElement("w:bidi"))
        r = p.add_run(str(text))
        r.font.name = "Arial"
        r.font.size = Pt(size)
        r.bold = bold
        rPr = r._r.get_or_add_rPr()
        rPr.append(OxmlElement("w:rtl"))
        return p

    rtl("الهيئة القومية للتأمين الاجتماعي", 16, True, WD_ALIGN_PARAGRAPH.CENTER)
    rtl("الإدارة المركزية للإدارات القانونية", 13, True, WD_ALIGN_PARAGRAPH.CENTER)
    rtl("الإدارة العامة للقضايا", 13, True, WD_ALIGN_PARAGRAPH.CENTER)
    if region:
        rtl(f"ديوان عام منطقة {region}", 12, True, WD_ALIGN_PARAGRAPH.CENTER)
    rtl(title, 13, True, WD_ALIGN_PARAGRAPH.CENTER)

    cols = list(df.columns) if len(df.columns) else ["م"]
    # عكس الأعمدة في جدول Word ليبدأ من اليمين إلى اليسار (العمود الأول يميناً)
    rtl_cols = list(reversed(cols))
    table = doc.add_table(rows=1, cols=len(rtl_cols))
    table.style = "Table Grid"
    tblPr = table._tbl.tblPr
    bidi_visual = OxmlElement("w:bidiVisual")
    tblPr.append(bidi_visual)

    for i, col in enumerate(rtl_cols):
        cell = table.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
        r = p.add_run(str(col))
        r.font.name = "Arial"
        r.font.size = Pt(10)
        r.bold = True
        r._r.get_or_add_rPr().append(OxmlElement("w:rtl"))

    for _, row in df.iterrows():
        cells = table.add_row().cells
        rtl_vals = list(reversed(list(row)))
        for i, val in enumerate(rtl_vals):
            cell = cells[i]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
            disp_val = _fmt_date_dmy(val) if re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(val or "").strip()) else str(val)
            r = p.add_run(disp_val)
            r.font.name = "Arial"
            r.font.size = Pt(9.5)
            r._r.get_or_add_rPr().append(OxmlElement("w:rtl"))

    doc.add_paragraph()
    rtl(f"تحريرًا في: {datetime.now().strftime('%d/%m/%Y')}", 10, False, WD_ALIGN_PARAGRAPH.RIGHT)
    rtl("وتفضلوا بقبول وافر الاحترام والتقدير،", 11, True, WD_ALIGN_PARAGRAPH.CENTER)

    # صف التوقيعات: العضو القانوني يمين، ومدير إدارة القضايا شمال على نفس السطر
    sig_tbl = doc.add_table(rows=1, cols=2)
    sig_tbl.autofit = False
    sig_tbl._tbl.tblPr.append(OxmlElement("w:bidiVisual"))
    right_cell = sig_tbl.rows[0].cells[0]
    left_cell = sig_tbl.rows[0].cells[1]

    p_r1 = right_cell.paragraphs[0]
    p_r1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r1._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    rr1 = p_r1.add_run("العضو القانوني")
    rr1.font.name = "Arial"; rr1.font.size = Pt(10); rr1.bold = True
    rr1._r.get_or_add_rPr().append(OxmlElement("w:rtl"))

    p_r2 = right_cell.add_paragraph()
    p_r2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r2._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    rr2 = p_r2.add_run(str(member_name or ""))
    rr2.font.name = "Arial"; rr2.font.size = Pt(10)
    rr2._r.get_or_add_rPr().append(OxmlElement("w:rtl"))

    p_r3 = right_cell.add_paragraph()
    p_r3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    rr3 = p_r3.add_run("....................")
    rr3.font.name = "Arial"; rr3.font.size = Pt(10)

    p_l1 = left_cell.paragraphs[0]
    p_l1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_l1._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    rl1 = p_l1.add_run("مدير إدارة القضايا")
    rl1.font.name = "Arial"; rl1.font.size = Pt(10); rl1.bold = True
    rl1._r.get_or_add_rPr().append(OxmlElement("w:rtl"))

    p_l2 = left_cell.add_paragraph()
    p_l2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_l2._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    rl2 = p_l2.add_run(str(manager_name or ""))
    rl2.font.name = "Arial"; rl2.font.size = Pt(10)
    rl2._r.get_or_add_rPr().append(OxmlElement("w:rtl"))

    p_l3 = left_cell.add_paragraph()
    p_l3.alignment = WD_ALIGN_PARAGRAPH.LEFT
    rl3 = p_l3.add_run("....................")
    rl3.font.name = "Arial"; rl3.font.size = Pt(10)

    # مدير عام الإدارات القانونية في المنتصف أسفل السطر السابق
    rtl("مدير عام الإدارات القانونية", 10, True, WD_ALIGN_PARAGRAPH.CENTER)
    rtl(str(general_manager or ""), 10, False, WD_ALIGN_PARAGRAPH.CENTER)
    rtl("....................", 10, False, WD_ALIGN_PARAGRAPH.CENTER)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def to_pdf(df, title, region, member_name="", specialization="", manager_name="", general_manager=""):
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=12)
    fp = _report_font_path()
    font = "ArabicReport" if fp else "Helvetica"
    if fp:
        pdf.add_font(font, "", fp, uni=True)
    pdf.set_font(font, "", 13)

    def ar(v):
        s = "" if v is None else str(v)
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s.strip()):
            s = _fmt_date_dmy(s)
        try:
            return fix_arabic(s)
        except Exception:
            return s

    pdf.cell(0, 9, ar("الهيئة القومية للتأمين الاجتماعي"), 0, 1, "C")
    pdf.set_font(font, "", 11)
    pdf.cell(0, 7, ar("الإدارة المركزية للإدارات القانونية"), 0, 1, "C")
    pdf.cell(0, 7, ar("الإدارة العامة للقضايا"), 0, 1, "C")
    if region:
        pdf.cell(0, 7, ar(f"ديوان عام منطقة {region}"), 0, 1, "C")
    pdf.cell(0, 8, ar(title), 0, 1, "C")
    pdf.ln(3)

    if len(df.columns):
        pdf.set_font(font, "", 7.5)
        usable_w = 277.0
        n_cols = len(df.columns)
        width = usable_w / n_cols
        start_x = (297.0 - usable_w) / 2.0
        # ترتيب الأعمدة من اليمين إلى اليسار (العمود الأول يميناً)
        rtl_cols = list(reversed(list(df.columns)))
        pdf.set_x(start_x)
        for col in rtl_cols:
            pdf.cell(width, 7, ar(col), 1, 0, "C")
        pdf.ln()
        for _, row in df.iterrows():
            rtl_vals = list(reversed(list(row)))
            pdf.set_x(start_x)
            for val in rtl_vals:
                pdf.cell(width, 7, ar(val), 1, 0, "C")
            pdf.ln()

    pdf.ln(5)
    pdf.set_font(font, "", 9)
    pdf.cell(0, 7, ar(f"تحريرًا في: {datetime.now().strftime('%d/%m/%Y')}"), 0, 1, "R")
    pdf.ln(2)
    pdf.cell(0, 7, ar("وتفضلوا بقبول وافر الاحترام والتقدير،"), 0, 1, "C")
    pdf.ln(4)

    # صف التوقيعات: العضو القانوني يمين، ومدير إدارة القضايا شمال على نفس السطر
    half_w = 138.5
    left_x = 14.0
    right_x = left_x + half_w

    y_top = pdf.get_y()
    pdf.set_xy(right_x, y_top)
    pdf.cell(half_w, 6, ar("العضو القانوني"), 0, 1, "R")
    pdf.set_x(right_x)
    pdf.cell(half_w, 6, ar(member_name), 0, 1, "R")
    pdf.set_x(right_x)
    pdf.cell(half_w, 6, ar("...................."), 0, 1, "R")
    y_after_right = pdf.get_y()

    pdf.set_xy(left_x, y_top)
    pdf.cell(half_w, 6, ar("مدير إدارة القضايا"), 0, 1, "L")
    pdf.set_x(left_x)
    pdf.cell(half_w, 6, ar(manager_name), 0, 1, "L")
    pdf.set_x(left_x)
    pdf.cell(half_w, 6, ar("...................."), 0, 1, "L")
    y_after_left = pdf.get_y()

    pdf.set_y(max(y_after_right, y_after_left) + 4)
    pdf.cell(0, 6, ar("مدير عام الإدارات القانونية"), 0, 1, "C")
    pdf.cell(0, 6, ar(general_manager), 0, 1, "C")
    pdf.cell(0, 6, ar("...................."), 0, 1, "C")

    raw = pdf.output()
    return raw.encode("latin-1", errors="replace") if isinstance(raw, str) else bytes(raw)
def create_paper_pdf(case_data):
    papers_dir = BASE_DIR / "papers"
    papers_dir.mkdir(parents=True, exist_ok=True)
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
    name = str(papers_dir / f"صحيفة_{case_data.get('رقم')}_{case_data.get('سنة')}.pdf"); pdf.output(name); return name

# --- نهاية التعديلات الجذرية ---

def create_case_file_word(case):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Inches(0.8); sec.bottom_margin = Inches(0.8); sec.left_margin = Inches(0.8); sec.right_margin = Inches(0.8)
    def rtl(t, s=11, b=False, a=WD_ALIGN_PARAGRAPH.RIGHT):
        p = doc.add_paragraph(); p.alignment = a; p._p.get_or_add_pPr().append(OxmlElement('w:bidi'))
        r = p.add_run(str(t)); r.font.name = 'Arial'; r.font.size = Pt(s); r.bold = b; return p
    rtl("الهيئة القومية للتأمين الاجتماعي", 15, True, WD_ALIGN_PARAGRAPH.CENTER)
    rtl("الإدارة المركزية للإدارات القانونية - الإدارة العامة للقضايا", 12, True, WD_ALIGN_PARAGRAPH.CENTER)
    rtl(f"ملف قضية رقم {case.get('رقم')} لسنة {case.get('سنة')} ({case.get('نوع')})", 13, True, WD_ALIGN_PARAGRAPH.CENTER)
    rtl("____________________________________________________", 10, False, WD_ALIGN_PARAGRAPH.CENTER)
    rtl("1- بيانات القضية والمحكمة:", 12, True)
    mamoria = f" - مأمورية {case.get('مأمورية')}" if case.get('مأمورية') else ""; rtl(f"• المحكمة: {case.get('محكمة_اسم')}{mamoria}")
    rtl(f"• الدائرة: {case.get('دائرة')}")
    rtl(f"• المدعي: {case.get('مدعي')}")
    rtl(f"• المدعي عليه: {case.get('مدعي_عليه')}")
    rtl(f"• الموضوع: {case.get('موضوع')}")
    rtl(f"• الحالة: {case.get('حالة')} | المكان: {case.get('مكان')}", 11, True)
    if case.get('جلسات'):
        rtl("2- سجل الجلسات:", 12, True)
        t = doc.add_table(rows=1, cols=4); t.style = 'Table Grid'
        for i, h in enumerate(['الرول', 'التاريخ', 'الاجراء', 'ملاحظات']): t.rows[0].cells[i].text = h
        for ج in case.get('جلسات', []):
            row = t.add_row().cells
            row[0].text = str(ج.get('الرول','')); row[1].text = str(_fmt_date_dmy(ج.get('تاريخ',''))); row[2].text = str(ج.get('الاجراء','')); row[3].text = str(ج.get('ملاحظات',''))
    if case.get('حالة') == 'منتهية':
        rtl("3- بيانات الحكم:", 12, True)
        rtl(f"• تاريخ الحكم: {_fmt_date_dmy(case.get('تاريخ_الحكم'))} | مسندة لـ: {case.get('مسندة_ل_الحكم')}")
        rtl(f"• منطوق الحكم: {case.get('منطوق_الحكم')}")
    if case.get('تم_الحفظ_النهائي'):
        rtl("4- بيانات الحفظ النهائي:", 12, True)
        rtl(f"• تاريخ الحفظ: {_fmt_date_dmy(case.get('تاريخ_الحفظ'))} | سبب الحفظ: {case.get('سبب_الحفظ')}")
    out = io.BytesIO(); doc.save(out); return out.getvalue()

def create_case_file_html(case):
    html_doc = f"""<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="UTF-8"><style>
    body {{ font-family: 'Arial', sans-serif; padding: 20px; color: #000; background: #fff; }}
    .header {{ text-align: center; border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 20px; }}
    .section {{ margin-bottom: 20px; }}
    .section-title {{ background: #eee; padding: 8px; font-weight: bold; border-right: 4px solid #000; margin-bottom: 10px; }}
    .row {{ margin-bottom: 6px; }}
    .label {{ font-weight: bold; display: inline-block; width: 140px; }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
    th, td {{ border: 1px solid #000; padding: 8px; text-align: right; }}
    th {{ background: #f2f2f2; }}
    </style></head><body>
    <div class="header"><h2>الهيئة القومية للتأمين الاجتماعي</h2><h3>الإدارة المركزية للإدارات القانونية - الإدارة العامة للقضايا</h3>
    <h3>ملف قضية رقم {case.get('رقم')} لسنة {case.get('سنة')} ({case.get('نوع')})</h3></div>
    <div class="section"><div class="section-title">1- بيانات القضية والمحكمة</div>
    <div class="row"><span class="label">المحكمة:</span> {case.get('محكمة_اسم')} {f'- مأمورية {case.get("مأمورية")}' if case.get('مأمورية') else ''}</div>
    <div class="row"><span class="label">الدائرة:</span> {case.get('دائرة')}</div>
    <div class="row"><span class="label">المدعي:</span> {case.get('مدعي')}</div>
    <div class="row"><span class="label">المدعي عليه:</span> {case.get('مدعي_عليه')}</div>
    <div class="row"><span class="label">الموضوع:</span> {case.get('موضوع')}</div>
    <div class="row"><span class="label">الحالة والمكان:</span> {case.get('حالة')} - {case.get('مكان')}</div></div>"""
    if case.get('جلسات'):
        html_doc += """<div class="section"><div class="section-title">2- سجل الجلسات</div><table><tr><th>الرول</th><th>التاريخ</th><th>الاجراء</th><th>ملاحظات</th></tr>"""
        for ج in case.get('جلسات', []):
            html_doc += f"<tr><td>{ج.get('الرول','')}</td><td>{_fmt_date_dmy(ج.get('تاريخ',''))}</td><td>{ج.get('الاجراء','')}</td><td>{ج.get('ملاحظات','')}</td></tr>"
        html_doc += "</table></div>"
    if case.get('حالة') == 'منتهية':
        html_doc += f"""<div class="section"><div class="section-title">3- بيانات الحكم</div>
        <div class="row"><span class="label">تاريخ الحكم:</span> {_fmt_date_dmy(case.get('تاريخ_الحكم'))}</div>
        <div class="row"><span class="label">مسندة لـ:</span> {case.get('مسندة_ل_الحكم')}</div>
        <div class="row"><span class="label">منطوق الحكم:</span> {case.get('منطوق_الحكم')}</div></div>"""
    if case.get('تم_الحفظ_النهائي'):
        html_doc += f"""<div class="section"><div class="section-title">4- بيانات الحفظ النهائي</div>
        <div class="row"><span class="label">تاريخ الحفظ:</span> {_fmt_date_dmy(case.get('تاريخ_الحفظ'))}</div>
        <div class="row"><span class="label">سبب الحفظ:</span> {case.get('سبب_الحفظ')}</div></div>"""
    html_doc += "</body></html>"
    return html_doc

def create_case_file_pdf(case):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    fp = _report_font_path()
    font = 'Cairo' if fp else 'Helvetica'
    if fp:
        pdf.add_font('Cairo', '', fp, uni=True)
    pdf.set_font(font, '', 14); pdf.set_right_margin(15)
    pdf.cell(0,10,fix_arabic("الهيئة القومية للتأمين الاجتماعي"),ln=1,align='C')
    pdf.set_font(font, '', 11)
    pdf.cell(0,8,fix_arabic(f"ملف قضية رقم {case.get('رقم')} لسنة {case.get('سنة')}"),ln=1,align='C')
    pdf.ln(4)
    pdf.cell(0,8,fix_arabic(f"المحكمة: {case.get('محكمة_اسم','')}"),ln=1,align='R')
    pdf.cell(0,8,fix_arabic(f"المدعي: {case.get('مدعي','')}"),ln=1,align='R')
    pdf.cell(0,8,fix_arabic(f"المدعي عليه: {case.get('مدعي_عليه','')}"),ln=1,align='R')
    pdf.multi_cell(0,8,fix_arabic(f"الموضوع: {case.get('موضوع','')}"),align='R')
    raw = pdf.output()
    return raw.encode("latin-1", errors="replace") if isinstance(raw, str) else bytes(raw)

# ==============================================================================
# 🗄️ التخزين المحلي الدائم + مزامنة اختيارية مع Supabase / GitHub
# ==============================================================================
LOCAL_DATA_DIR = BASE_DIR / "local_storage_data"
LOCAL_DATA_DIR.mkdir(parents=True, exist_ok=True)
USERS_FILE = LOCAL_DATA_DIR / "users_data.json"
DATA_FILE = LOCAL_DATA_DIR / "cases_data.json"
TOKENS_FILE = LOCAL_DATA_DIR / "tokens_data.json"
LIBRARY_FILE = LOCAL_DATA_DIR / "library_data.json"

def _read_json_file(path, default_factory):
    try:
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                txt = f.read().strip()
                if txt:
                    return json.loads(txt)
    except Exception:
        pass
    return default_factory() if callable(default_factory) else default_factory

def _write_json_file(path, payload):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def load_json_persistent(key, local_path, default_value):
    default_fn = (lambda: copy.deepcopy(default_value)) if not callable(default_value) else default_value
    if key in st.session_state and st.session_state[key] is not None:
        return st.session_state[key]
    data_val = _read_json_file(local_path, default_fn)
    st.session_state[key] = data_val
    return data_val

def save_json_persistent(key, local_path, payload):
    st.session_state[key] = payload
    _write_json_file(local_path, payload)
    return True

def upload_bytes_to_supabase(file_bytes, filename, folder="uploads"):
    up_dir = LOCAL_DATA_DIR / folder
    up_dir.mkdir(parents=True, exist_ok=True)
    safe_name = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{re.sub(r'[^A-Za-z0-9._-]', '_', filename)}"
    target = up_dir / safe_name
    with open(target, "wb") as f:
        f.write(file_bytes)
    return {"ok": True, "path": str(target), "url": str(target)}

def load_data():
    data=load_json_persistent("data", DATA_FILE, {"cases":[],"library":[]})
    if not isinstance(data,dict):
        data={}
    data.setdefault("cases",[]); data.setdefault("library",[])
    changed = False
    admin = next((u for u in load_users() if u.get("role") == "admin"), None)
    admin_id = admin.get("id") if admin else None
    admin_username = admin.get("username") if admin else ""
    for case in data["cases"]:
        if "owner_user_id" not in case and case.get("assigned_to"):
            owner = next((u for u in load_users() if u.get("username") == case.get("assigned_to")), None)
            if owner:
                case["owner_user_id"] = owner.get("id")
                changed = True
        if "assigned_to" not in case and admin_id is not None:
            case["assigned_to"] = admin_username
            case["owner_user_id"] = admin_id
            changed = True
    if changed:
        save_json_persistent("data", DATA_FILE, data)
    return data
def load_notifications():
    n = load_json_persistent("notifications", LOCAL_DATA_DIR / "notifications_data.json", list)
    return n if isinstance(n, list) else []

def save_notifications(n):
    save_json_persistent("notifications", LOCAL_DATA_DIR / "notifications_data.json", n)

def compute_notifications(cases):
    alerts = []
    today = datetime.now().date()
    appeal_days = 40
    for c in cases:
        if c.get("تم_الحفظ_النهائي"):
            continue
        case_title = f"قضية رقم {c.get('رقم','')} لسنة {c.get('سنة','')}"
        if c.get("حالة") == "متداولة" and c.get("تاريخ_جلسة"):
            try:
                session_date = datetime.strptime(str(c["تاريخ_جلسة"])[:10], "%Y-%m-%d").date()
                days_left = (session_date - today).days
                if -2 <= days_left <= 7:
                    alerts.append({
                        "type": "session",
                        "case_id": c.get("id"),
                        "title": case_title,
                        "date": _fmt_date_dmy(c["تاريخ_جلسة"]),
                        "days_left": days_left,
                        "case": c
                    })
            except Exception:
                pass
        if c.get("حالة") == "منتهية" and c.get("مسندة_ل_الحكم") == "الضد" and c.get("تاريخ_الحكم"):
            try:
                judgment_date = datetime.strptime(str(c["تاريخ_الحكم"])[:10], "%Y-%m-%d").date()
                last_appeal_day = judgment_date + timedelta(days=appeal_days)
                days_left = (last_appeal_day - today).days
                if -5 <= days_left <= 15:
                    case_copy = dict(c)
                    case_copy["deadline"] = _fmt_date_dmy(last_appeal_day)
                    alerts.append({
                        "type": "appeal",
                        "case_id": c.get("id"),
                        "title": case_title,
                        "date": _fmt_date_dmy(last_appeal_day),
                        "days_left": days_left,
                        "case": case_copy
                    })
            except Exception:
                pass
    return alerts

def send_email_alerts(to_email, alerts):
    if not to_email or not alerts:
        return False, "بيانات البريد أو التنبيهات غير متوفرة"
    smtp_server = _get_secret("SMTP_SERVER", "")
    smtp_port = int(_get_secret("SMTP_PORT", "587") or "587")
    smtp_user = _get_secret("SMTP_USER", "")
    smtp_pass = _get_secret("SMTP_PASS", "")
    if not (smtp_server and smtp_user and smtp_pass):
        return False, "إعدادات SMTP غير مضافة في Secrets"
    subject = f"🔔 تنبيهات ادارة القضايا - {datetime.now().strftime('%d/%m/%Y')}"
    body = "<div dir='rtl' style='font-family:Arial;'><h3>تنبيهات القضايا والجلسات</h3>"
    for a in alerts:
        if a["type"] == "session":
            body += f"<p><b>{a['title']}</b> - الجلسة: {a['date']} - <b style='color:red;'>متبقي {a['days_left']} يوم</b></p>"
        else:
            body += f"<p><b>{a['title']}</b> - آخر ميعاد طعن: {a['date']} - <b style='color:red;'>متبقي {a['days_left']} يوم</b></p>"
    body += "</div>"
    try:
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart
        msg = MIMEMultipart()
        msg["From"] = smtp_user
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html", "utf-8"))
        with smtplib.SMTP(smtp_server, smtp_port, timeout=15) as server:
            server.starttls()
            server.login(smtp_user, smtp_pass)
            server.sendmail(smtp_user, to_email, msg.as_string())
        return True, "تم إرسال التنبيهات عبر البريد الإلكتروني بنجاح"
    except Exception as e:
        return False, str(e)

LIBRARY_CATEGORIES = {
    "احكام المحكمة الدستورية العليا": "#DC143C",
    "احكام محكمة النقض": "#B22222",
    "احكام المحكمة الإدارية العليا": "#8B0000",
    "احكام المحكمة الإدارية": "#DEB887",
    "منشورات القضاء العادى": "#5F9EA0",
    "منشورات مجلس الدولة": "#4682B4",
}

ANWA3_MOSTANDAT = [
    "صحيفة دعوى", "صحيفة استئناف", "صحيفة طعن", "مذكرة دفاع", "حافظة مستندات",
    "تقرير خبير", "تقرير طب شرعى", "تقرير لجنة طبية", "صحيفة تجديد من الشطب",
    "صحيفة تعجيل من الوقف", "صورة حكم تمهيدى", "أخرى"
]

def load_library():
    lib = load_json_persistent("library", LIBRARY_FILE, list)
    return lib if isinstance(lib, list) else []

def save_library(lib):
    save_json_persistent("library", LIBRARY_FILE, lib)

# --- التحقق من تسجيل الدخول ---
if not st.session_state.auth_user:
    render_login_screen()
    st.stop()

_refresh_auth_user()
current_user = st.session_state.auth_user
if current_user.get("status") != "active":
    st.error("⛔ حسابك غير نشط حالياً. يرجى التواصل مع الإدارة.")
    if st.button("تسجيل الخروج"):
        st.session_state.auth_user = None
        st.rerun()
    st.stop()

data = load_data()
all_cases = data.get("cases", [])
visible_cases = [c for c in all_cases if can_view_case(current_user, c)]

# --- شريط التنقل العلوي ---
show_banners("top")

top_cols = st.columns([2, 2, 2, 2, 2, 2])
with top_cols[0]:
    if st.button("📝 تسجيل", use_container_width=True):
        st.session_state.page = "تسجيل"
        st.rerun()
with top_cols[1]:
    if st.button("⚖️ المتداولة", use_container_width=True):
        st.session_state.page = "المتداولة"
        st.rerun()
with top_cols[2]:
    if st.button("📂 الأرشيف", use_container_width=True):
        st.session_state.page = "الأرشيف"
        st.rerun()
with top_cols[3]:
    if st.button("📊 التقارير", use_container_width=True):
        st.session_state.page = "التقارير"
        st.rerun()
with top_cols[4]:
    if st.button("📚 المكتبة", use_container_width=True):
        st.session_state.page = "مكتبة"
        st.rerun()
with top_cols[5]:
    if is_admin_like(current_user):
        if st.button("⚙️ لوحة التحكم", use_container_width=True, key="admin_control_main_unique"):
            st.session_state.page = "لوحة التحكم"
            st.rerun()
    else:
        if st.button("⚙️ لوحة تحكمى", use_container_width=True, key="member_control_main_unique"):
            st.session_state.page = "لوحة تحكمى"
            st.rerun()

# شريط إضافي للبحث والتنبيهات والرسائل وتسجيل الخروج
sub_cols = st.columns([2, 2, 2, 2])
with sub_cols[0]:
    if st.button("🔍 بحث متقدم", use_container_width=True):
        st.session_state.page = "بحث"
        st.rerun()
with sub_cols[1]:
    _, unread_list = get_user_notifications(current_user)
    msg_label = f"✉️ الرسائل ({len(unread_list)})" if unread_list else "✉️ الرسائل"
    if st.button(msg_label, use_container_width=True):
        st.session_state.page = "الرسائل"
        st.rerun()
with sub_cols[2]:
    alerts_now = compute_notifications(visible_cases)
    alert_label = f"🔔 التنبيهات ({len(alerts_now)})" if alerts_now else "🔔 التنبيهات"
    if st.button(alert_label, use_container_width=True):
        st.session_state.page = "التنبيهات"
        st.rerun()
with sub_cols[3]:
    if st.button("🚪 خروج", use_container_width=True):
        st.session_state.auth_user = None
        st.rerun()

st.markdown(f"<div style='text-align:right; color:#D4AF37; font-weight:800; margin:6px 0;'>👤 {current_user.get('full_name')} — {ROLES.get(current_user.get('role'), current_user.get('role'))} | آخر تحديث: {_fmt_date_dmy(data.get('last_update')) or '-'}</div>", unsafe_allow_html=True)

# --- صفحات التطبيق ---
if st.session_state.page == "تسجيل":
    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>📝 تسجيل دعوى / طعن جديد</h1>", unsafe_allow_html=True)
    with st.form("add_case_form", clear_on_submit=True):
        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>1- بيانات المحكمة</div>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            نوع = st.selectbox("نوع الدعوى / الطعن", ["ابتدائي", "استئناف", "نقض", "إداري", "إدارية عليا", "دستورية", "أخرى"])
            رقم = st.text_input("رقم الدعوى / الطعن")
            سنة = st.text_input("السنة")
        with c2:
            محكمة_اسم = st.text_input("اسم المحكمة", key="court_name_add")
            مأمورية = st.text_input("المأمورية")
            دائرة = st.text_input("الدائرة")
        with c3:
            region = st.text_input("المنطقة", value=current_user.get("region", "البحيرة"))
            مسندة_ل = st.selectbox("الصفة / المسندة", ["مدعي", "مدعي عليه", "مستأنف", "مستأنف ضده", "طاعن", "مطعون ضده"])
            مكان = st.text_input("مكان الملف / الإدارة", value=current_user.get("region", "البحيرة"))

        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>2- بيانات الخصوم والموضوع</div>", unsafe_allow_html=True)
        p1, p2 = st.columns(2)
        with p1:
            مدعي = st.text_input("اسم المدعي / المستأنف / الطاعن")
            موضوع = st.text_area("موضوع الدعوى")
        with p2:
            مدعي_عليه = st.text_input("اسم المدعي عليه / المستأنف ضده / المطعون ضده")
            assigned_lawyer = st.text_input("المحامي / العضو المختص", value=current_user.get("full_name", ""))

        st.markdown("<div style='color:#D4AF37; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>4- بيانات الجلسة</div>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            تاريخ_جلسة = st.date_input("تاريخ أول جلسة", value=datetime.now().date(), key="session_date_add", format="DD/MM/YYYY")
        with col2:
            الرول = st.text_input("الرول")
        with col3:
            الاجراء = st.text_input("الإجراء", value="أول جلسة")
        ملاحظات = st.text_area("ملاحظات")

        submitted_case = st.form_submit_button("💾 حفظ القضية", use_container_width=True, type="primary")
        if submitted_case:
            if not رقم.strip() or not سنة.strip():
                st.error("يرجى إدخال رقم الدعوى والسنة.")
            else:
                existing_ids = [int(x.get("id")) for x in data.get("cases", []) if str(x.get("id", "")).isdigit()]
                next_id = (max(existing_ids) + 1) if existing_ids else (len(data.get("cases", [])) + 1)
                case_for_pdf = {
                    "نوع": نوع, "رقم": رقم, "سنة": سنة, "دائرة": دائرة,
                    "محكمة_اسم": محكمة_اسم, "مدعي": مدعي, "مدعي_عليه": مدعي_عليه,
                    "موضوع": موضوع, "تاريخ_جلسة": str(تاريخ_جلسة), "مسندة_ل": مسندة_ل
                }
                try:
                    create_paper_pdf(case_for_pdf)
                except Exception:
                    pass
                new_case = {
                    "id": next_id,
                    "نوع": نوع,
                    "رقم": رقم.strip(),
                    "سنة": سنة.strip(),
                    "محكمة_اسم": محكمة_اسم.strip(),
                    "مأمورية": مأمورية.strip(),
                    "دائرة": دائرة.strip(),
                    "region": region.strip(),
                    "مسندة_ل": مسندة_ل,
                    "مكان": مكان.strip(),
                    "مدعي": مدعي.strip(),
                    "مدعي_عليه": مدعي_عليه.strip(),
                    "موضوع": موضوع.strip(),
                    "تاريخ_جلسة": str(تاريخ_جلسة),
                    "الرول": الرول.strip(),
                    "الاجراء": الاجراء.strip(),
                    "ملاحظات": ملاحظات.strip(),
                    "حالة": "متداولة",
                    "assigned_lawyer": assigned_lawyer.strip(),
                    "assigned_to": [assigned_lawyer.strip()] if assigned_lawyer.strip() else [],
                    "created_by": current_user.get("username"),
                    "created_at": datetime.now().isoformat(),
                    "جلسات": [{
                        "تاريخ": str(تاريخ_جلسة),
                        "الرول": الرول.strip(),
                        "الاجراء": الاجراء.strip(),
                        "ملاحظات": ملاحظات.strip(),
                        "الحالة": "متداولة"
                    }],
                    "مستندات": []
                }
                data["cases"].append(new_case)
                save_data(data)
                st.success("✅ تم تسجيل القضية بنجاح.")
                st.rerun()

elif st.session_state.page == "المتداولة":
    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>⚖️ القضايا المتداولة</h1>", unsafe_allow_html=True)
    active_cases = [c for c in visible_cases if c.get("حالة") == "متداولة" and not c.get("تم_الحفظ_النهائي")]
    for case in active_cases:
        if case.get("جلسات"):
            جلسات_مرتبة = sorted(case["جلسات"], key=lambda x: x.get("تاريخ", "9999-12-31"), reverse=True)
            اخر_جلسة = جلسات_مرتبة[0]
            case["تاريخ_جلسة"] = اخر_جلسة.get("تاريخ", case.get("تاريخ_جلسة", ""))
            case["الاجراء"] = اخر_جلسة.get("الاجراء", case.get("الاجراء", ""))
            case["الحالة"] = اخر_جلسة.get("الحالة", case.get("الحالة", "متداولة"))

    sorted_cases = sorted(active_cases, key=lambda x: x.get("تاريخ_جلسة", "9999-12-31"))
    today = datetime.now().date()
    start_week = today - timedelta(days=today.weekday())
    end_week = start_week + timedelta(days=6)
    this_week = len([c for c in active_cases if c.get('تاريخ_جلسة') and start_week <= datetime.strptime(str(c['تاريخ_جلسة'])[:10], '%Y-%m-%d').date() <= end_week])
    reserved = len([c for c in active_cases if any(k in str(c.get('الاجراء', '')) for k in ['حكم', 'للحكم', 'الحكم'])])

    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f"<div style='font-size:28px; font-weight:900; color:#D4AF37'>{len(sorted_cases)}</div><div style='font-size:18px; color:#FFF; font-weight:700'>إجمالي المتداولة</div>", unsafe_allow_html=True)
    with m2:
        st.markdown(f"<div style='font-size:28px; font-weight:900; color:#4CAF50'>{this_week}</div><div style='font-size:18px; color:#FFF; font-weight:700'>جلسات هذا الأسبوع</div>", unsafe_allow_html=True)
    with m3:
        st.markdown(f"<div style='font-size:28px; font-weight:900; color:#FF5252'>⚖️ {reserved}</div><div style='font-size:18px; color:#FFF; font-weight:700'>المحجوز للحكم</div>", unsafe_allow_html=True)

    st.markdown("""
    <style>
    .case-table { width:100%; border-collapse:collapse; direction:rtl; text-align:right; background:#142038; color:#FFF; margin-top:12px; }
    .case-table th { background:#1E2A47; color:#FFD700; padding:10px; border:1px solid #D4AF37; font-weight:900; }
    .case-table td { padding:9px; border:1px solid #2d426b; vertical-align:top; }
    .date-gold { color:#FFD700; font-weight:900; }
    .status-green { color:#4CAF50; font-weight:900; }
    </style>
    """, unsafe_allow_html=True)

    if sorted_cases:
        headers = ["م", "الرقم والسنة", "المحكمة والدائرة", "الخصوم", "الموضوع", "اخر جلسة", "الاجراء", "الحالة"]
        table_html = "<table class='case-table'><tr>" + "".join(f"<th>{h}</th>" for h in headers) + "</tr>"
        for idx, case in enumerate(sorted_cases, 1):
            رقم_كامل = f"{case.get('رقم','')}/{case.get('سنة','')}"
            محكمة_كاملة = f"{case.get('نوع','')} {case.get('محكمة_اسم','')}"
            if case.get('مأمورية', ''):
                محكمة_كاملة += f"<br>مأمورية {case.get('مأمورية','')}"
            دائرة_كاملة = case.get('دائرة', '')
            if دائرة_كاملة:
                محكمة_كاملة += f"<br>دائرة {دائرة_كاملة}"
            خصوم = f"{case.get('مدعي','')} ضد {case.get('مدعي_عليه','')}"
            table_html += f"<tr><td>{idx}</td><td>{رقم_كامل}</td><td>{محكمة_كاملة}</td><td>{خصوم}</td><td>{case.get('موضوع','')}</td><td class='date-gold'>{_fmt_date_dmy(case.get('تاريخ_جلسة',''))}</td><td>{case.get('الاجراء','')}</td><td class='status-green'>{case.get('حالة','متداولة')}</td></tr>"
        table_html += "</table>"
        st.markdown(table_html, unsafe_allow_html=True)

        st.markdown("### 📌 اختيار قضية لمتابعة الجلسات أو تسجيل الحكم")
        options = {c.get("id"): f"{c.get('رقم')}/{c.get('سنة')} — {c.get('محكمة_اسم')} ({c.get('مدعي')} ضد {c.get('مدعي_عليه')})" for c in sorted_cases}
        chosen_id = st.selectbox("اختر القضية", options=list(options.keys()), format_func=lambda k: options.get(k, str(k)))
        case = next((c for c in sorted_cases if c.get("id") == chosen_id), None)

        if case:
            st.markdown("---")
            c_a, c_b, c_c = st.columns(3)
            with c_a:
                st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>القضية</div><div style='color:#FFF; font-weight:700; font-size:14px'>{case.get('رقم')}/{case.get('سنة')} ({case.get('نوع','')})</div></div>", unsafe_allow_html=True)
            with c_b:
                محكمة_كاملة = f"{case.get('محكمة_اسم')}"
                if case.get('مأمورية'):
                    محكمة_كاملة += f" - مأمورية {case.get('مأمورية')}"
                st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>المحكمة</div><div style='color:#FFF; font-weight:700; font-size:14px'>{محكمة_كاملة}</div></div>", unsafe_allow_html=True)
            with c_c:
                st.markdown(f"<div style='background:#142038; padding:12px; border-radius:12px; border:1px solid #D4AF37; margin-bottom:10px; text-align:center'><div style='color:#D4AF37; font-weight:900; font-size:14px'>آخر جلسة</div><div style='color:#FFD700; font-weight:900; font-size:14px'>{_fmt_date_dmy(case.get('تاريخ_جلسة',''))}</div></div>", unsafe_allow_html=True)

            with st.expander("✏️ تعديل بيانات القضية"):
                with st.form(f"edit_case_form_{case['id']}"):
                    رقم_جديد = st.text_input("الرقم", value=case.get('رقم', ''))
                    سنة_جديد = st.text_input("السنة", value=case.get('سنة', ''))
                    نوع_جديد = st.text_input("النوع", value=case.get('نوع', ''))
                    محكمة_جديد = st.text_input("اسم المحكمة", value=case.get('محكمة_اسم', ''))
                    مأمورية_جديد = st.text_input("المأمورية", value=case.get('مأمورية', ''))
                    دائرة_جديد = st.text_input("الدائرة", value=case.get('دائرة', ''))
                    مدعي_جديد = st.text_input("المدعي", value=case.get('مدعي', ''))
                    مدعي_عليه_جديد = st.text_input("المدعي عليه", value=case.get('مدعي_عليه', ''))
                    حالة_جديد = st.selectbox("الحالة", ["متداولة", "منتهية"], index=0 if case.get('حالة', 'متداولة') == "متداولة" else 1)
                    موضوع_جديد = st.text_area("الموضوع", value=case.get('موضوع', ''))
                    if st.form_submit_button("💾 حفظ تعديل القضية", use_container_width=True):
                        case['رقم'] = رقم_جديد
                        case['سنة'] = سنة_جديد
                        case['نوع'] = نوع_جديد
                        case['محكمة_اسم'] = محكمة_جديد
                        case['مأمورية'] = مأمورية_جديد
                        case['دائرة'] = دائرة_جديد
                        case['مدعي'] = مدعي_جديد
                        case['مدعي_عليه'] = مدعي_عليه_جديد
                        case['حالة'] = حالة_جديد
                        case['موضوع'] = موضوع_جديد
                        save_data(data)
                        st.success("✅ تم حفظ تعديل القضية")
                        st.rerun()

            st.markdown("### 📅 سجل الجلسات")
            case.setdefault("جلسات", [])
            for i, ج in enumerate(case["جلسات"]):
                st.markdown(f"<div style='background:#142038; padding:15px; border-radius:12px; border:2px solid #D4AF37; margin-bottom:10px; text-align:right; direction:rtl'><div style='display:flex; justify-content:flex-end; margin-bottom:10px'><div style='background:#D4AF37; color:#000; padding:5px 15px; border-radius:8px; font-weight:900; font-size:16px'>جلسة {i+1}</div></div><div style='margin-bottom:8px'><span style='color:#D4AF37; font-weight:900'>التاريخ:</span> <span style='color:#FFF'>{_fmt_date_dmy(ج.get('تاريخ'))}</span></div><div style='margin-bottom:8px'><span style='color:#D4AF37; font-weight:900'>الرول:</span> <span style='color:#FFF'>{ج.get('الرول')}</span></div><div style='margin-bottom:8px'><span style='color:#D4AF37; font-weight:900'>الاجراء:</span> <span style='color:#FFF'>{ج.get('الاجراء')}</span></div><div><span style='color:#D4AF37; font-weight:900'>ملاحظات:</span> <span style='color:#FFF'>{ج.get('ملاحظات')}</span></div></div>", unsafe_allow_html=True)
                if st.button("✏️ تعديل الجلسة", key=f"edit_session_{case['id']}_{i}", use_container_width=True):
                    st.session_state["edit_session_case_id"] = case["id"]
                    st.session_state["edit_session_index"] = i
                    st.rerun()

            if st.session_state.get("edit_session_case_id") == case["id"] and "edit_session_index" in st.session_state:
                idx = st.session_state.edit_session_index
                if 0 <= idx < len(case["جلسات"]):
                    جلسة = case["جلسات"][idx]
                    st.warning(f"تعديل الجلسة رقم {idx+1}")
                    try:
                        init_dt = datetime.strptime(str(جلسة.get('تاريخ'))[:10], '%Y-%m-%d').date()
                    except Exception:
                        init_dt = datetime.now().date()
                    تاريخ_تعديل = st.date_input("التاريخ", value=init_dt, format="DD/MM/YYYY", key=f"ed_sess_dt_{case['id']}_{idx}")
                    رول_تعديل = st.text_input("الرول", value=جلسة.get('الرول', ''), key=f"ed_sess_roll_{case['id']}_{idx}")
                    اجراء_تعديل = st.text_input("الاجراء", value=جلسة.get('الاجراء', ''), key=f"ed_sess_act_{case['id']}_{idx}")
                    ملاحظات_تعديل = st.text_area("الملاحظات", value=جلسة.get('ملاحظات', ''), key=f"ed_sess_note_{case['id']}_{idx}")
                    with st.form(f"save_session_edit_form_{case['id']}_{idx}"):
                        if st.form_submit_button("💾 حفظ تعديل الجلسة", use_container_width=True):
                            case["جلسات"][idx] = {"تاريخ": str(تاريخ_تعديل), "الرول": رول_تعديل, "الاجراء": اجراء_تعديل, "ملاحظات": ملاحظات_تعديل, "الحالة": جلسة.get("الحالة", "متداولة")}
                            جلسات_مرتبة = sorted(case["جلسات"], key=lambda x: x.get("تاريخ", "9999-12-31"), reverse=True)
                            case["تاريخ_جلسة"] = جلسات_مرتبة[0].get("تاريخ", "")
                            case["الاجراء"] = جلسات_مرتبة[0].get("الاجراء", "")
                            save_data(data)
                            st.session_state.pop("edit_session_index", None)
                            st.success("✅ تم حفظ تعديل الجلسة")
                            st.rerun()

            with st.expander("اضافة جلسة جديدة"):
                with st.form(f"add_new_session_form_{case['id']}", clear_on_submit=True):
                    تاريخ_جديد = st.date_input("تاريخ الجلسة", value=datetime.now().date(), format="DD/MM/YYYY", key=f"new_sess_dt_{case['id']}")
                    رول_جديد = st.text_input("الرول")
                    الاجراء_جديد = st.text_input("الاجراء")
                    ملاحظات_جديدة = st.text_area("ملاحظات")
                    if st.form_submit_button("حفظ الجلسة"):
                        case["جلسات"].append({"تاريخ": str(تاريخ_جديد), "الرول": رول_جديد, "الاجراء": الاجراء_جديد, "ملاحظات": ملاحظات_جديدة, "الحالة": "متداولة"})
                        case["تاريخ_جلسة"] = str(تاريخ_جديد)
                        case["الاجراء"] = الاجراء_جديد
                        save_data(data)
                        st.success("تم اضافة الجلسة")
                        st.rerun()

            with st.expander("📎 إرفاق مستند للقضية"):
                uploaded_file = st.file_uploader("اختر الملف", type=['pdf', 'jpg', 'jpeg', 'png', 'doc', 'docx'], key=f"case_upl_{case['id']}")
                نوع_مستند = st.selectbox("نوع المستند", ANWA3_MOSTANDAT, key=f"case_doc_type_{case['id']}")
                if st.button("رفع المستند", key=f"btn_upl_case_{case['id']}") and uploaded_file is not None:
                    up_res = upload_bytes_to_supabase(uploaded_file.getvalue(), uploaded_file.name, folder=f"cases/{case['id']}")
                    case.setdefault("مستندات", []).append({
                        "name": uploaded_file.name,
                        "type": نوع_مستند,
                        "url": up_res.get("url"),
                        "uploaded_at": datetime.now().strftime("%d/%m/%Y %H:%M")
                    })
                    save_data(data)
                    st.success("تم إرفاق المستند بنجاح")
                    st.rerun()

            st.markdown("<div style='color:#FF5252; font-size:20px; font-weight:900; text-align:center; margin-bottom:10px'>5- جلسة الحكم</div>", unsafe_allow_html=True)
            if case.get('تاريخ_الحكم'):
                لون = "#4CAF50" if case.get('مسندة_ل_الحكم') == "الصالح" else "#FF5252"
                st.markdown(f"<div style='background:#142038; padding:15px; border-radius:12px; border:2px solid {لون}; margin-bottom:10px'><b style='color:{لون}'>تاريخ جلسة الحكم:</b> {_fmt_date_dmy(case.get('تاريخ_الحكم'))}<br><b style='color:{لون}'>مسندة لـ:</b> {case.get('مسندة_ل_الحكم')}<br><b style='color:{لون}'>منطوق الحكم:</b> {case.get('منطوق_الحكم')}</div>", unsafe_allow_html=True)
                with st.expander("✏️ تعديل بيانات الحكم"):
                    with st.form(f"edit_judgment_form_{case['id']}"):
                        try:
                            j_dt = datetime.strptime(str(case.get('تاريخ_الحكم'))[:10], '%Y-%m-%d').date()
                        except Exception:
                            j_dt = datetime.now().date()
                        تاريخ_حكم_تعديل = st.date_input("تاريخ الحكم", value=j_dt, format="DD/MM/YYYY", key=f"ed_j_dt_{case['id']}")
                        منطوق_الحكم_تعديل = st.text_area("منطوق الحكم", value=case.get('منطوق_الحكم', ''), height=150)
                        مسندة_ل_تعديل = st.selectbox("مسندة لـ", ["الصالح", "الضد"], index=["الصالح", "الضد"].index(case.get('مسندة_ل_الحكم', 'الصالح')) if case.get('مسندة_ل_الحكم') in ["الصالح", "الضد"] else 0)
                        if st.form_submit_button("💾 حفظ تعديل الحكم", use_container_width=True, type="primary"):
                            case['تاريخ_الحكم'] = str(تاريخ_حكم_تعديل)
                            case['منطوق_الحكم'] = منطوق_الحكم_تعديل
                            case['مسندة_ل_الحكم'] = مسندة_ل_تعديل
                            for ج in case.get('جلسات', []):
                                if 'الحكم' in ج.get('الاجراء', ''):
                                    ج['تاريخ'] = str(تاريخ_حكم_تعديل)
                                    ج['الاجراء'] = f'الحكم - مسندة لـ {مسندة_ل_تعديل}'
                                    ج['ملاحظات'] = منطوق_الحكم_تعديل
                                    break
                            case['تاريخ_جلسة'] = str(تاريخ_حكم_تعديل)
                            case['الاجراء'] = f'الحكم - مسندة لـ {مسندة_ل_تعديل}'
                            save_data(data)
                            st.success("✅ تم تعديل الحكم")
                            st.rerun()

            with st.form(f"record_judgment_form_{case['id']}"):
                st.markdown("<label style='color:#FFD700; font-weight:900; font-size:16px'>1- تاريخ الجلسة</label>", unsafe_allow_html=True)
                تاريخ_حكم = st.date_input("تاريخ الجلسة", value=datetime.now().date(), label_visibility="collapsed", format="DD/MM/YYYY", key=f"rec_j_dt_{case['id']}")
                st.markdown("<label style='color:#FFD700; font-weight:900; font-size:16px'>2- منطوق الحكم</label>", unsafe_allow_html=True)
                منطوق_الحكم = st.text_area("منطوق الحكم", height=150, placeholder="اكتب منطوق الحكم هنا...", label_visibility="collapsed")
                مسندة_ل = st.selectbox("3- مسندة لـ", ["الصالح", "الضد"], key=f"rec_j_side_{case['id']}")
                if st.form_submit_button("💾 حفظ الحكم", use_container_width=True, type="primary"):
                    if not منطوق_الحكم.strip():
                        st.error("❌ لازم تكتب منطوق الحكم")
                    else:
                        case['حالة'] = 'منتهية'
                        case['تاريخ_الحكم'] = str(تاريخ_حكم)
                        case['منطوق_الحكم'] = منطوق_الحكم.strip()
                        case['مسندة_ل_الحكم'] = مسندة_ل
                        case['جلسات'].append({
                            'تاريخ': str(تاريخ_حكم),
                            'الرول': '-',
                            'الاجراء': f'الحكم - مسندة لـ {مسندة_ل}',
                            'ملاحظات': منطوق_الحكم.strip(),
                            'الحالة': 'منتهية'
                        })
                        case['تاريخ_جلسة'] = str(تاريخ_حكم)
                        case['الاجراء'] = f'الحكم - مسندة لـ {مسندة_ل}'
                        save_data(data)
                        st.success("✅ حفظت ونقلت للارشيف للمتابعه يتم الانتقال للارشيف")
                        st.session_state.page = "الأرشيف"
                        st.rerun()
    else:
        st.info("لا توجد قضايا متداولة حالياً.")

elif st.session_state.page == "الأرشيف":
    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>📂 الأرشيف والأحكام</h1>", unsafe_allow_html=True)
    if not can_access_archive(current_user):
        st.error("⛔ لا تملك صلاحية الوصول إلى الأرشيف.")
    else:
        tab_judgments, tab_saved = st.tabs(["⚖️ الأحكام الصادرة", "🗄️ الحفظ النهائي"])
        with tab_judgments:
            st.markdown("<div style='color:#FFF; font-size:18px; font-weight:900; text-align:center; margin-bottom:10px'>🔍 البحث عن قضية صدر فيها الحكم</div>", unsafe_allow_html=True)
            q_judg = st.text_input("ابحث برقم القضية أو السنة أو الخصوم أو المنطوق", key="search_judgments")
            judg_cases = [c for c in visible_cases if c.get('حالة') == 'منتهية' and not c.get('تم_الحفظ_النهائي')]
            if q_judg.strip():
                qq = q_judg.strip()
                judg_cases = [c for c in judg_cases if any(qq in str(c.get(k, '')) for k in ['رقم', 'سنة', 'مدعي', 'مدعي_عليه', 'منطوق_الحكم', 'محكمة_اسم'])]

            for case in judg_cases:
                لون = "#4CAF50" if case.get('مسندة_ل_الحكم') == "الصالح" else "#FF5252"
                محكمة_str = f"{case.get('محكمة_اسم')} {f'- مأمورية {case.get(\"مأمورية\")}' if case.get('مأمورية') else ''}"
                st.markdown(f"""
                <table style='width:100%; border-collapse:collapse; direction:rtl; margin-bottom:12px;'>
                <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; width:35%; font-weight:900;'>القضية</td><td style='background:#FFF; color:#000; padding:8px; font-weight:700;'>{case.get('رقم')}/{case.get('سنة')} ({case.get('نوع','')})</td></tr>
                <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; width:35%; font-weight:900;'>المحكمة</td><td style='background:#FFF; color:#000; padding:8px; font-weight:700;'>{محكمة_str}</td></tr>
                <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; font-weight:900;'>الخصوم</td><td style='background:#FFF; color:#000; padding:8px; font-weight:700;'>{case.get('مدعي')} ضد {case.get('مدعي_عليه')}</td></tr>
                <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; font-weight:900;'>تاريخ الحكم</td><td style='background:#FFF; color:{لون}; padding:8px; font-weight:900;'>{_fmt_date_dmy(case.get('تاريخ_الحكم'))}</td></tr>
                <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; font-weight:900;'>المنطوق</td><td style='background:#FFF; color:{لون}; padding:8px; font-weight:900;'>{case.get('منطوق_الحكم')}</td></tr>
                <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; font-weight:900; border-radius:0 0 0 8px;'>مسندة لـ</td><td style='background:#FFF; color:{لون}; padding:8px; font-weight:900; border-radius:0 0 8px 0;'>{case.get('مسندة_ل_الحكم')}</td></tr>
                </table>
                """, unsafe_allow_html=True)

                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    st.download_button("📄 ملف القضية Word", data=create_case_file_word(case), file_name=f"ملف_قضية_{case.get('رقم')}_{case.get('سنة')}.docx", key=f"w_j_{case['id']}")
                with col_f2:
                    st.download_button("📑 ملف القضية HTML", data=create_case_file_html(case), file_name=f"ملف_قضية_{case.get('رقم')}_{case.get('سنة')}.html", key=f"h_j_{case['id']}")

                with st.expander("🗄️ نقل إلى الحفظ النهائي"):
                    with st.form(f"final_save_form_{case['id']}"):
                        سبب_الحفظ = st.text_area("سبب الحفظ", placeholder="مثال: تم الطعن / حكم نهائي / عدم جدوى")
                        مستندات_الحفظ = st.file_uploader("ارفع مستندات الحفظ", type=['pdf', 'jpg', 'png', 'doc', 'docx'], accept_multiple_files=True, key=f"f_upl_{case['id']}")
                        if st.form_submit_button("تأكيد الحفظ النهائي", use_container_width=True):
                            for mf in (مستندات_الحفظ or []):
                                up_r = upload_bytes_to_supabase(mf.getvalue(), mf.name, folder=f"archive/{case['id']}")
                                case.setdefault("مستندات", []).append({
                                    "name": mf.name,
                                    "type": "مستند حفظ نهائي",
                                    "url": up_r.get("url"),
                                    "uploaded_at": datetime.now().strftime("%d/%m/%Y %H:%M")
                                })
                            case['تم_الحفظ_النهائي'] = True
                            case['تاريخ_الحفظ'] = str(datetime.now().date())
                            case['سبب_الحفظ'] = سبب_الحفظ.strip()
                            save_data(data)
                            st.success("تم نقل القضية إلى الحفظ النهائي")
                            st.rerun()

        with tab_saved:
            saved_cases = [c for c in visible_cases if c.get('تم_الحفظ_النهائي')]
            q_sav = st.text_input("ابحث في القضايا المحفوظة نهائياً", key="search_saved")
            if q_sav.strip():
                qq = q_sav.strip()
                saved_cases = [c for c in saved_cases if any(qq in str(c.get(k, '')) for k in ['رقم', 'سنة', 'مدعي', 'مدعي_عليه', 'سبب_الحفظ', 'محكمة_اسم'])]
            for case in saved_cases:
                محكمة_str = f"{case.get('محكمة_اسم')} {f'- مأمورية {case.get(\"مأمورية\")}' if case.get('مأمورية') else ''}"
                st.markdown(f"""
                <table style='width:100%; border-collapse:collapse; direction:rtl; margin-bottom:12px;'>
                <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; width:35%; font-weight:900;'>القضية</td><td style='background:#FFF; color:#000; padding:8px; font-weight:700;'>{case.get('رقم')}/{case.get('سنة')} ({case.get('نوع','')})</td></tr>
                <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; width:35%; font-weight:900;'>المحكمة</td><td style='background:#FFF; color:#000; padding:8px; font-weight:700;'>{محكمة_str}</td></tr>
                <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; font-weight:900;'>الخصوم</td><td style='background:#FFF; color:#000; padding:8px; font-weight:700;'>{case.get('مدعي')} ضد {case.get('مدعي_عليه')}</td></tr>
                <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; font-weight:900;'>تاريخ الحفظ</td><td style='background:#FFF; color:#4CAF50; padding:8px; font-weight:900;'>{_fmt_date_dmy(case.get('تاريخ_الحفظ'))}</td></tr>
                <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; font-weight:900;'>سبب الحفظ</td><td style='background:#FFF; color:#000; padding:8px; font-weight:700;'>{case.get('سبب_الحفظ','')}</td></tr>
                </table>
                """, unsafe_allow_html=True)
elif st.session_state.page == "بحث":
    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>🔍 البحث المتقدم</h1>", unsafe_allow_html=True)
    q = st.text_input("ابحث برقم الدعوى أو السنة أو الخصوم أو المحكمة أو الموضوع", key="adv_search_q")
    status_filter = st.selectbox("حالة القضية", ["الكل", "متداولة", "منتهية", "محفوظة نهائياً"])
    results = visible_cases
    if status_filter == "متداولة":
        results = [c for c in results if c.get("حالة") == "متداولة" and not c.get("تم_الحفظ_النهائي")]
    elif status_filter == "منتهية":
        results = [c for c in results if c.get("حالة") == "منتهية" and not c.get("تم_الحفظ_النهائي")]
    elif status_filter == "محفوظة نهائياً":
        results = [c for c in results if c.get("تم_الحفظ_النهائي")]

    if q.strip():
        qq = q.strip()
        results = [
            c for c in results
            if any(qq in str(c.get(k, "")) for k in ["رقم", "سنة", "مدعي", "مدعي_عليه", "محكمة_اسم", "مأمورية", "دائرة", "موضوع", "منطوق_الحكم"])
        ]

    st.write(f"عدد النتائج: {len(results)}")
    for c in results:
        محكمة_كاملة = f"{c.get('نوع','')} {c.get('محكمة_اسم','')}"
        if c.get('مأمورية', ''):
            محكمة_كاملة += f"<br>مأمورية {c.get('مأمورية','')}"
        if c.get('دائرة', ''):
            محكمة_كاملة += f"<br>دائرة {c.get('دائرة','')}"
        st.markdown(f"""
        <table style='width:100%; border-collapse:collapse; direction:rtl; margin-bottom:10px;'>
        <tr><th>م</th><th>الرقم والسنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>اخر جلسة</th><th>الحالة</th><th>المكان</th></tr>
        <tr>
            <td>{c.get('id','')}</td>
            <td>{c.get('رقم','')}/{c.get('سنة','')}</td>
            <td>{محكمة_كاملة}</td>
            <td>{c.get('مدعي','')} ضد {c.get('مدعي_عليه','')}</td>
            <td>{c.get('موضوع','')}</td>
            <td style='color:#FFD700; font-weight:900'>{_fmt_date_dmy(c.get('تاريخ_جلسة','')) or '-'}</td>
            <td>{c.get('حالة','')}</td>
            <td>{c.get('مكان','')}</td>
        </tr>
        </table>
        """, unsafe_allow_html=True)

        with st.expander(f"عرض تفاصيل القضية {c.get('رقم')}/{c.get('سنة')}"):
            محكمة_تفصيل = f"{c.get('محكمة_اسم')} {f'- مأمورية {c.get(\"مأمورية\")}' if c.get('مأمورية') else ''}"
            st.markdown(f"""
            <table style='width:100%; border-collapse:collapse; direction:rtl; margin-bottom:10px;'>
            <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; font-weight:900;'>المحكمة</td><td style='background:#FFF; color:#000; padding:8px; font-weight:700;'>{محكمة_تفصيل}</td></tr>
            <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; font-weight:900;'>الخصوم</td><td style='background:#FFF; color:#000; padding:8px; font-weight:700;'>{c.get('مدعي','')} ضد {c.get('مدعي_عليه','')}</td></tr>
            <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; font-weight:900;'>الموضوع</td><td style='background:#FFF; color:#000; padding:8px; font-weight:700;'>{c.get('موضوع','')}</td></tr>
            <tr><td style='background:#1E2A47; color:#FFD700; padding:8px; width:30%; font-weight:900;'>تاريخ اخر جلسة</td><td style='background:#FFF; color:#000; padding:8px; font-weight:700;'>{_fmt_date_dmy(c.get('تاريخ_جلسة','')) or '-'}</td></tr>
            </table>
            """, unsafe_allow_html=True)

            if c.get('جلسات'):
                جلسات_مرتبة = sorted(c['جلسات'], key=lambda x: x.get("تاريخ", ""), reverse=True)
                for ج in جلسات_مرتبة:
                    st.markdown(f"<div style='background:#1E2A47; padding:10px; border-radius:8px; margin-bottom:5px; border:1px solid #D4AF37'><b style='color:#FFD700'>تاريخ:</b> <span style='color:#FFF'>{_fmt_date_dmy(ج.get('تاريخ'))}</span> | <b style='color:#FFD700'>الاجراء:</b> <span style='color:#FFF'>{ج.get('الاجراء')}</span> | <b style='color:#FFD700'>الحالة:</b> <span style='color:#FFF'>{ج.get('الحالة')}</span></div>", unsafe_allow_html=True)

elif st.session_state.page == "التنبيهات":
    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>🔔 التنبيهات والمواعيد</h1>", unsafe_allow_html=True)
    alerts = compute_notifications(visible_cases)
    if current_user.get("email"):
        if st.button("📧 إرسال التنبيهات إلى بريدي الإلكتروني", use_container_width=True):
            ok, msg = send_email_alerts(current_user.get("email"), alerts)
            if ok:
                st.success(msg)
            else:
                st.warning(msg)

    if not alerts:
        st.info("لا توجد تنبيهات عاجلة حالياً.")
    for a in alerts:
        if a["type"] == "session":
            st.warning(f"📅 **{a['title']}** — الجلسة: {a['date']} (متبقي {a['days_left']} يوم)")
        else:
            body_txt = f"⚖️ **{a['title']}** — آخر ميعاد طعن: {a['date']} (متبقي {a['days_left']} يوم)"
            st.error(body_txt)

elif st.session_state.page == "مكتبة":
    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>📚 المكتبة القانونية والأحكام والمنشورات</h1>", unsafe_allow_html=True)
    lib_items = load_library()
    with st.expander("➕ إضافة حكم أو منشور جديد"):
        with st.form("add_lib_form", clear_on_submit=True):
            cat = st.selectbox("التصنيف", list(LIBRARY_CATEGORIES.keys()))
            title_l = st.text_input("العنوان / رقم الحكم أو المنشور")
            summary_l = st.text_area("الملخص أو النص القانوني")
            doc_file = st.file_uploader("إرفاق ملف (اختياري)", type=["pdf", "doc", "docx", "jpg", "png"])
            if st.form_submit_button("💾 حفظ في المكتبة", use_container_width=True):
                if not title_l.strip():
                    st.error("يرجى إدخال العنوان.")
                else:
                    file_url = None
                    file_name = None
                    if doc_file is not None:
                        up_res = upload_bytes_to_supabase(doc_file.getvalue(), doc_file.name, folder="library")
                        file_url = up_res.get("url")
                        file_name = doc_file.name
                    lib_items.insert(0, {
                        "id": f"lib_{uuid.uuid4().hex[:8]}",
                        "category": cat,
                        "title": title_l.strip(),
                        "summary": summary_l.strip(),
                        "file_name": file_name,
                        "file_url": file_url,
                        "added_by": current_user.get("full_name"),
                        "created_at": datetime.now().strftime("%d/%m/%Y")
                    })
                    save_library(lib_items)
                    st.success("تمت الإضافة إلى المكتبة بنجاح.")
                    st.rerun()

    sel_cat = st.selectbox("تصفية حسب التصنيف", ["الكل"] + list(LIBRARY_CATEGORIES.keys()))
    q_lib = st.text_input("بحث في المكتبة القانونية")
    filtered_lib = lib_items
    if sel_cat != "الكل":
        filtered_lib = [x for x in filtered_lib if x.get("category") == sel_cat]
    if q_lib.strip():
        ql = q_lib.strip()
        filtered_lib = [x for x in filtered_lib if ql in str(x.get("title", "")) or ql in str(x.get("summary", ""))]

    for item in filtered_lib:
        color = LIBRARY_CATEGORIES.get(item.get("category"), "#D4AF37")
        st.markdown(f"<div style='background:#142038; padding:14px; border-radius:12px; border-right:5px solid {color}; margin-bottom:10px;'><div style='color:{color}; font-weight:900;'>{item.get('category')} — {_fmt_date_dmy(item.get('created_at'))}</div><div style='color:#FFF; font-size:17px; font-weight:800; margin:6px 0;'>{item.get('title')}</div><div style='color:#DCE6F2;'>{item.get('summary','')}</div></div>", unsafe_allow_html=True)
        if item.get("file_url") and os.path.exists(str(item.get("file_url"))):
            with open(str(item.get("file_url")), "rb") as fp_f:
                st.download_button(f"📥 تحميل {item.get('file_name') or 'الملف'}", data=fp_f.read(), file_name=item.get("file_name") or "document.pdf", key=f"dl_lib_{item.get('id')}")

elif st.session_state.page == "التقارير":
    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>📊 استخراج البيانات والتقارير الرسمية</h1>", unsafe_allow_html=True)
    if not can_access_reports(current_user):
        st.error("⛔ لا تملك صلاحية استخراج التقارير.")
    else:
        rep_tab1, rep_tab2, rep_tab3 = st.tabs(["📋 بيان الدعاوى المتداولة", "⚖️ بيان الأحكام", "🔢 بيان عددي"])

        with rep_tab1:
            colA, colB, colC = st.columns(3)
            with colA:
                نوع_تقرير_متداولة = st.selectbox("نوع البيان", ["بيان بالدعاوى المتداولة حسب موضوع الدعوى", "بيان بالدعاوى المتداولة حسب المحكمة", "بيان عام بالدعاوى المتداولة"], key="rep1_type")
                region = st.text_input("المنطقة", value=current_user.get("region", "البحيرة"), key="reg1")
                عضو_قانوني1 = st.text_input("العضو القانوني", value=current_user.get("full_name", ""), key="mem1")
                lawyer = st.text_input("اسم المحامي / العضو", value=current_user.get("full_name", ""), key="law1")
            with colB:
                مدير_عام1 = st.text_input("اسم مدير عام الإدارات القانونية", key="modir1")
            with colC:
                مدير_ادارة1 = st.text_input("اسم مدير إدارة القضايا", key="modir_idara1")

            col1, col2 = st.columns(2)
            with col1:
                from_date = st.date_input("من الفترة", key="from1", format="DD/MM/YYYY")
            with col2:
                to_date = st.date_input("حتى الفترة", key="to1", format="DD/MM/YYYY")

            if st.button("📊 إعداد بيان المتداولة", use_container_width=True, key="btn_rep1"):
                cases = [c for c in visible_cases if c.get("حالة") == "متداولة" and not c.get("تم_الحفظ_النهائي")]
                filtered = []
                for c in cases:
                    if c.get('تاريخ_جلسة'):
                        try:
                            ت_جلسة = datetime.strptime(str(c['تاريخ_جلسة'])[:10], '%Y-%m-%d').date()
                            if from_date <= ت_جلسة <= to_date:
                                filtered.append(c)
                        except Exception:
                            filtered.append(c)
                    else:
                        filtered.append(c)
                cases = sorted(filtered, key=lambda x: x.get("تاريخ_جلسة", "0000-00-00"), reverse=True)
                title = f"{نوع_تقرير_متداولة} خلال الفترة من {from_date.strftime('%d/%m/%Y')} إلى {to_date.strftime('%d/%m/%Y')} طرف الاستاذ/ {lawyer} المحامي"
                st.success(f"✅ تم العثور على {len(cases)} دعوى")
                df_data = []
                html_tbl = "<table class='case-table'><tr><th>م</th><th>رقم</th><th>سنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>اخر جلسة</th><th>الاجراء</th><th>ملاحظات</th></tr>"
                for idx, c in enumerate(cases, 1):
                    محكمة = f"{c.get('نوع','')} {c.get('محكمة_اسم','')}"
                    if c.get('مأمورية'):
                        محكمة += f" - مأمورية {c.get('مأمورية')}"
                    if c.get('دائرة'):
                        محكمة += f" - دائرة {c.get('دائرة')}"
                    مدعي = c.get('مدعي', '')
                    مدعي_عليه = c.get('مدعي_عليه', '')
                    خصوم_html = f"{مدعي} ضد {مدعي_عليه}"
                    html_tbl += f"<tr><td>{idx}</td><td>{c.get('رقم','')}</td><td>{c.get('سنة','')}</td><td>{محكمة}</td><td>{خصوم_html}</td><td>{c.get('موضوع','')}</td><td><b style='color:#B8860B'>{_fmt_date_dmy(c.get('تاريخ_جلسة',''))}</b></td><td>{c.get('الاجراء','')}</td><td>{c.get('ملاحظات','')}</td></tr>"
                    df_data.append({
                        'م': idx,
                        'رقم': c.get('رقم', ''),
                        'سنة': c.get('سنة', ''),
                        'المحكمة': محكمة,
                        'الخصوم': f"{مدعي} ضد {مدعي_عليه}",
                        'الموضوع': c.get('موضوع', ''),
                        'اخر جلسة': _fmt_date_dmy(c.get('تاريخ_جلسة', '')),
                        'الاجراء': c.get('الاجراء', ''),
                        'ملاحظات': c.get('ملاحظات', '')
                    })
                html_tbl += "</table>"
                st.markdown(html_tbl, unsafe_allow_html=True)
                footer = f"""<div style="margin-top:25px; color:#B8860B; font-size:12px;"><p style="text-align:center; margin-bottom:20px; font-size:13px; font-weight:700;">تفضلوا بقبول وافر الاحترام والتقدير،</p><table style="width:100%;"><tr><td style="width:50%; text-align:right; vertical-align:top;"><div style="font-weight:900;">العضو القانوني</div><div>{عضو_قانوني1}</div><div style="margin-top:12px;">....................</div><div style="margin-top:20px;">تحرر في: {datetime.now().strftime('%d/%m/%Y')}</div></td><td style="width:50%; text-align:left; vertical-align:top;"><div style="font-weight:900;">مدير إدارة القضايا</div><div>{مدير_ادارة1}</div><div style="margin-top:12px;">....................</div></td></tr></table><div style="text-align:center; margin-top:20px;"><div style="font-weight:900; color:#dc3545;">مدير عام الإدارات القانونية</div><div>{مدير_عام1}</div><div style="margin-top:12px;">....................</div></div></div>"""
                st.markdown(footer, unsafe_allow_html=True)

                d1, d2, d3 = st.columns(3)
                with d1:
                    st.download_button("📄 استخراج PDF", data=to_pdf(pd.DataFrame(df_data), title, region, عضو_قانوني1, lawyer, مدير_ادارة1, مدير_عام1), file_name=report_download_name("بيان_الدعاوى_المتداولة", عضو_قانوني1, lawyer, region, "pdf"), mime="application/pdf", use_container_width=True, key="pdf_dlx1")
                with d2:
                    st.download_button("📝 Word", data=to_word(pd.DataFrame(df_data), title, region, عضو_قانوني1, lawyer, مدير_ادارة1, مدير_عام1), file_name=report_download_name("بيان_الدعاوى_المتداولة", عضو_قانوني1, lawyer, region, "docx"), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="word_dlx1")
                with d3:
                    st.download_button("📊 Excel", data=to_excel(pd.DataFrame(df_data), title), file_name=report_download_name("بيان_الدعاوى_المتداولة", عضو_قانوني1, lawyer, region, "xlsx"), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="xls_dlx1")

        with rep_tab2:
            colA, colB, colC = st.columns(3)
            with colA:
                نوع_التقرير = st.selectbox("نوع بيان الأحكام", ["بيان بالأحكام الصادرة", "بيان بالأحكام الصادرة لصالح الهيئة", "بيان بالأحكام الصادرة ضد الهيئة"], key="rep2_type")
                region2 = st.text_input("المنطقة", value=current_user.get("region", "البحيرة"), key="reg2")
                عضو_قانوني2 = st.text_input("العضو القانوني", value=current_user.get("full_name", ""), key="mem2")
                lawyer2 = st.text_input("اسم المحامي / العضو", value=current_user.get("full_name", ""), key="law2")
            with colB:
                مدير_عام2 = st.text_input("اسم مدير عام الإدارات القانونية", key="modir2")
            with colC:
                مدير_ادارة2 = st.text_input("اسم مدير إدارة القضايا", key="modir_idara2")

            col1, col2 = st.columns(2)
            with col1:
                from_date2 = st.date_input("من الفترة", key="from2", format="DD/MM/YYYY")
            with col2:
                to_date2 = st.date_input("حتى الفترة", key="to2", format="DD/MM/YYYY")

            if st.button("⚖️ إعداد بيان الأحكام", use_container_width=True, key="btn_rep2"):
                cases = [c for c in visible_cases if c.get("حالة") == "منتهية"]
                filtered = []
                for c in cases:
                    if c.get('تاريخ_الحكم'):
                        try:
                            ت_حكم = datetime.strptime(str(c['تاريخ_الحكم'])[:10], '%Y-%m-%d').date()
                            if from_date2 <= ت_حكم <= to_date2:
                                filtered.append(c)
                        except Exception:
                            filtered.append(c)
                cases = filtered
                if "لصالح" in نوع_التقرير:
                    cases = [c for c in cases if c.get('مسندة_ل_الحكم') == 'الصالح']
                elif "ضد" in نوع_التقرير:
                    cases = [c for c in cases if c.get('مسندة_ل_الحكم') == 'الضد']
                cases = sorted(cases, key=lambda x: x.get("تاريخ_الحكم", "0000-00-00"), reverse=True)
                title = f"{نوع_التقرير} خلال الفترة من {from_date2.strftime('%d/%m/%Y')} إلى {to_date2.strftime('%d/%m/%Y')} طرف الاستاذ/ {lawyer2} المحامي"
                st.success(f"✅ تم العثور على {len(cases)} حكم")
                df_data = []
                html_tbl = "<table class='case-table'><tr><th>م</th><th>رقم</th><th>سنة</th><th>المحكمة والدائرة</th><th>الخصوم</th><th>الموضوع</th><th>تاريخ الحكم</th><th>منطوق الحكم</th><th>مسندة ل</th><th>ملاحظات</th></tr>"
                for idx, c in enumerate(cases, 1):
                    محكمة = f"{c.get('نوع','')} {c.get('محكمة_اسم','')}"
                    if c.get('مأمورية'):
                        محكمة += f" - مأمورية {c.get('مأمورية')}"
                    if c.get('دائرة'):
                        محكمة += f" - دائرة {c.get('دائرة')}"
                    مدعي = c.get('مدعي', '')
                    مدعي_عليه = c.get('مدعي_عليه', '')
                    خصوم_html = f"{مدعي} ضد {مدعي_عليه}"
                    لون_مسندة = "#28a745" if c.get('مسندة_ل_الحكم') == 'الصالح' else "#dc3545"
                    html_tbl += f"<tr><td>{idx}</td><td>{c.get('رقم','')}</td><td>{c.get('سنة','')}</td><td>{محكمة}</td><td>{خصوم_html}</td><td>{c.get('موضوع','')}</td><td><b style='color:#B8860B'>{_fmt_date_dmy(c.get('تاريخ_الحكم',''))}</b></td><td>{c.get('منطوق_الحكم','')}</td><td><b style='color:{لون_مسندة}'>{c.get('مسندة_ل_الحكم')}</b></td><td>{c.get('ملاحظات','')}</td></tr>"
                    df_data.append({
                        'م': idx,
                        'رقم': c.get('رقم', ''),
                        'سنة': c.get('سنة', ''),
                        'المحكمة': محكمة,
                        'الخصوم': f"{مدعي} ضد {مدعي_عليه}",
                        'الموضوع': c.get('موضوع', ''),
                        'تاريخ الحكم': _fmt_date_dmy(c.get('تاريخ_الحكم', '')),
                        'منطوق الحكم': c.get('منطوق_الحكم', ''),
                        'مسندة ل': c.get('مسندة_ل_الحكم'),
                        'ملاحظات': c.get('ملاحظات', '')
                    })
                html_tbl += "</table>"
                st.markdown(html_tbl, unsafe_allow_html=True)
                footer = f"""<div style="margin-top:25px; color:#B8860B; font-size:12px;"><p style="text-align:center; margin-bottom:20px; font-size:13px; font-weight:700;">تفضلوا بقبول وافر الاحترام والتقدير،</p><table style="width:100%;"><tr><td style="width:50%; text-align:right; vertical-align:top;"><div style="font-weight:900;">العضو القانوني</div><div>{عضو_قانوني2}</div><div style="margin-top:12px;">....................</div><div style="margin-top:20px;">تحرر في: {datetime.now().strftime('%d/%m/%Y')}</div></td><td style="width:50%; text-align:left; vertical-align:top;"><div style="font-weight:900;">مدير إدارة القضايا</div><div>{مدير_ادارة2}</div><div style="margin-top:12px;">....................</div></td></tr></table><div style="text-align:center; margin-top:20px;"><div style="font-weight:900; color:#dc3545;">مدير عام الإدارات القانونية</div><div>{مدير_عام2}</div><div style="margin-top:12px;">....................</div></div></div>"""
                st.markdown(footer, unsafe_allow_html=True)

                d1, d2, d3 = st.columns(3)
                with d1:
                    st.download_button("📄 استخراج PDF", data=to_pdf(pd.DataFrame(df_data), title, region2, عضو_قانوني2, lawyer2, مدير_ادارة2, مدير_عام2), file_name=report_download_name("بيان_الاحكام", عضو_قانوني2, lawyer2, region2, "pdf"), mime="application/pdf", use_container_width=True, key="pdf_dlx2")
                with d2:
                    st.download_button("📝 Word", data=to_word(pd.DataFrame(df_data), title, region2, عضو_قانوني2, lawyer2, مدير_ادارة2, مدير_عام2), file_name=report_download_name("بيان_الاحكام", عضو_قانوني2, lawyer2, region2, "docx"), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="word_dlx2")
                with d3:
                    st.download_button("📊 Excel", data=to_excel(pd.DataFrame(df_data), title), file_name=report_download_name("بيان_الاحكام", عضو_قانوني2, lawyer2, region2, "xlsx"), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="xls_dlx2")

        with rep_tab3:
            col1, col2 = st.columns(2)
            with col1:
                stat_from = st.date_input("من تاريخ", key="s1", format="DD/MM/YYYY")
            with col2:
                stat_to = st.date_input("حتى تاريخ", key="s2", format="DD/MM/YYYY")

            if st.button("📈 عرض ملخص الإحصاء", use_container_width=True, key="btn_stat_summary"):
                متداولة = [c for c in visible_cases if c.get('حالة') == 'متداولة' and c.get('تاريخ_جلسة') and stat_from <= datetime.strptime(str(c['تاريخ_جلسة'])[:10], '%Y-%m-%d').date() <= stat_to]
                احكام = [c for c in visible_cases if c.get('حالة') == 'منتهية' and not c.get('تم_الحفظ_النهائي') and c.get('تاريخ_الحكم') and stat_from <= datetime.strptime(str(c['تاريخ_الحكم'])[:10], '%Y-%m-%d').date() <= stat_to]
                للصالح = [c for c in احكام if c.get('مسندة_ل_الحكم') == 'الصالح']
                للضد = [c for c in احكام if c.get('مسندة_ل_الحكم') == 'الضد']
                s1, s2, s3, s4 = st.columns(4)
                s1.metric("متداولة", len(متداولة))
                s2.metric("أحكام صادرة", len(احكام))
                s3.metric("لصالح الهيئة", len(للصالح))
                s4.metric("ضد الهيئة", len(للضد))

            st.markdown("---")
            colA, colB, colC = st.columns(3)
            with colA:
                اسم_البيان = st.text_input("اسم البيان العددي", value="الدعاوى والأحكام", key="stat_name")
                region_stat = st.text_input("المنطقة", value=current_user.get("region", "البحيرة"), key="reg_stat")
                عضو_قانوني_stat = st.text_input("العضو القانوني", value=current_user.get("full_name", ""), key="mem_stat")
                lawyer_stat = st.text_input("اسم المحامي / العضو", value=current_user.get("full_name", ""), key="law_stat")
            with colB:
                مدير_عام_stat = st.text_input("اسم مدير عام الإدارات القانونية", key="modir_stat")
            with colC:
                مدير_ادارة_stat = st.text_input("اسم مدير إدارة القضايا", key="modir_idara_stat")

            col1, col2 = st.columns(2)
            with col1:
                stat_from2 = st.date_input("من الفترة", key="s_from2", format="DD/MM/YYYY")
            with col2:
                stat_to2 = st.date_input("حتى الفترة", key="s_to2", format="DD/MM/YYYY")

            العدد = st.number_input("العدد", min_value=0, step=1, key="stat_count_val")
            ملاحظات = st.text_area("ملاحظات", key="stat_notes_val")

            if st.button("📄 إعداد البيان العددي", use_container_width=True, key="btn_make_stat"):
                title_stat = f"بيان عددي ب{اسم_البيان} خلال الفترة من {stat_from2.strftime('%d/%m/%Y')} إلى {stat_to2.strftime('%d/%m/%Y')} طرف الاستاذ/ {lawyer_stat} المحامي"
                row_dict = {'م': 1, 'البيان': اسم_البيان, 'العدد': العدد, 'ملاحظات': ملاحظات}
                st.markdown(f"<table class='case-table'><tr><th>م</th><th>البيان</th><th>العدد</th><th>ملاحظات</th></tr><tr><td>1</td><td>{اسم_البيان}</td><td>{العدد}</td><td>{ملاحظات}</td></tr></table>", unsafe_allow_html=True)
                footer = f"""<div style="margin-top:25px; color:#B8860B; font-size:12px;"><p style="text-align:center; margin-bottom:20px; font-size:13px; font-weight:700;">تفضلوا بقبول وافر الاحترام والتقدير،</p><table style="width:100%;"><tr><td style="width:50%; text-align:right; vertical-align:top;"><div style="font-weight:900;">العضو القانوني</div><div>{عضو_قانوني_stat}</div><div style="margin-top:12px;">....................</div><div style="margin-top:20px;">تحرر في: {datetime.now().strftime('%d/%m/%Y')}</div></td><td style="width:50%; text-align:left; vertical-align:top;"><div style="font-weight:900;">مدير إدارة القضايا</div><div>{مدير_ادارة_stat}</div><div style="margin-top:12px;">....................</div></td></tr></table><div style="text-align:center; margin-top:20px;"><div style="font-weight:900; color:#dc3545;">مدير عام الإدارات القانونية</div><div>{مدير_عام_stat}</div><div style="margin-top:12px;">....................</div></div></div>"""
                st.markdown(footer, unsafe_allow_html=True)

                d1, d2 = st.columns(2)
                with d1:
                    st.download_button("📄 استخراج PDF", data=to_pdf(pd.DataFrame([row_dict]), title_stat, region_stat, عضو_قانوني_stat, lawyer_stat, مدير_ادارة_stat, مدير_عام_stat), file_name=report_download_name("بيان_عددي", عضو_قانوني_stat, lawyer_stat, region_stat, "pdf"), mime="application/pdf", use_container_width=True, key="pdf_dlx4")
                with d2:
                    st.download_button("📝 Word", data=to_word(pd.DataFrame([row_dict]), title_stat, region_stat, عضو_قانوني_stat, lawyer_stat, مدير_ادارة_stat, مدير_عام_stat), file_name=report_download_name("بيان_عددي", عضو_قانوني_stat, lawyer_stat, region_stat, "docx"), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="word_dlx4")

elif st.session_state.page == "إدارة الأعضاء":
    render_user_management()

elif st.session_state.page == "الرسائل":
    render_messages_page()

elif st.session_state.page == "لوحة التحكم":
    render_admin_dashboard()

elif st.session_state.page == "لوحة تحكمى":
    render_member_dashboard()
