# === بداية الجزء 1 من 3 ===
from pathlib import Path

import json, os, bcrypt, smtplib, random, io, requests, base64, html, secrets, re, sqlite3, hashlib
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


# ==== دوال توحيد وتنسيق التواريخ (يوم / شهر / سنة) بقائمة منسدلة =====
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
    return f"{d.day:02d}/{d.month:02d}/{d.year:04d}"


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

    c_d, c_m, c_y = st.columns([1, 1, 1.15], gap="small")
    with c_y:
        sel_y = st.selectbox(
            "السنة",
            years,
            index=years.index(d_val.year),
            key=f"{base_key}_y",
            label_visibility="collapsed",
        )
    months = list(range(1, 13))
    with c_m:
        sel_m = st.selectbox(
            "الشهر",
            months,
            format_func=lambda m: f"{m:02d}",
            index=months.index(d_val.month),
            key=f"{base_key}_m",
            label_visibility="collapsed",
        )
    days = list(range(1, 32))
    safe_day_idx = min(max(d_val.day - 1, 0), 30)
    with c_d:
        sel_d = st.selectbox(
            "اليوم",
            days,
            format_func=lambda d: f"{d:02d}",
            index=safe_day_idx,
            key=f"{base_key}_d",
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
TOP_BAR_FILE = DATA_DIR / "top_bar_config.json"

SITE_AUDIO_FILE = DATA_DIR / "site_audio_config.json"

DEFAULT_SITE_AUDIO_CONFIG = {
    "enabled": False,          # True = يعمل، False = متوقف
    "source_type": "none",     # "upload" | "url" | "none"
    "title": "",
    "file_name": "",
    "mime_type": "audio/mpeg",
    "audio_b64": "",
    "audio_url": "",
    "loop": True,
    "volume": 0.7,
    "updated_at": "",
    "updated_by": "",
}


def _normalize_site_audio_config(raw):
    cfg = dict(DEFAULT_SITE_AUDIO_CONFIG)
    if isinstance(raw, dict):
        for k in cfg:
            if k in raw and raw[k] is not None:
                cfg[k] = raw[k]
    cfg["enabled"] = bool(cfg.get("enabled", False))
    cfg["loop"] = bool(cfg.get("loop", True))
    try:
        cfg["volume"] = max(0.0, min(1.0, float(cfg.get("volume", 0.7))))
    except Exception:
        cfg["volume"] = 0.7
    if cfg.get("source_type") not in ("upload", "url", "none"):
        cfg["source_type"] = "none"
    return cfg


def load_site_audio_config():
    try:
        data = load_json_persistent("site_audio_config", SITE_AUDIO_FILE, dict(DEFAULT_SITE_AUDIO_CONFIG))
        return _normalize_site_audio_config(data)
    except Exception:
        return dict(DEFAULT_SITE_AUDIO_CONFIG)


def save_site_audio_config(cfg):
    normalized = _normalize_site_audio_config(cfg)
    normalized["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ok = save_json_persistent("site_audio_config", SITE_AUDIO_FILE, normalized)
    st.session_state["site_audio_config_cache"] = normalized
    return ok


def get_site_audio_config(force_reload=False):
    if force_reload or "site_audio_config_cache" not in st.session_state:
        st.session_state["site_audio_config_cache"] = load_site_audio_config()
    return _normalize_site_audio_config(st.session_state["site_audio_config_cache"])


def _guess_site_audio_mime(filename, uploaded_type=""):
    if uploaded_type and str(uploaded_type).startswith("audio/"):
        return uploaded_type
    ext = Path(filename or "").suffix.lower()
    return {
        ".mp3": "audio/mpeg",
        ".wav": "audio/wav",
        ".ogg": "audio/ogg",
        ".m4a": "audio/mp4",
        ".aac": "audio/aac",
        ".webm": "audio/webm",
    }.get(ext, "audio/mpeg")


def _has_valid_site_audio(cfg):
    stype = cfg.get("source_type")
    if stype == "upload":
        return bool(str(cfg.get("audio_b64") or "").strip())
    if stype == "url":
        return bool(str(cfg.get("audio_url") or "").strip())
    return False


def render_site_audio_player():
    """تشغيل صوت الموقع بعد تسجيل الدخول فقط عندما يكون مفعّلاً من قِبل الأدمن."""
    if not st.session_state.get("user"):
        return
    cfg = get_site_audio_config()
    if not cfg.get("enabled") or not _has_valid_site_audio(cfg):
        return

    title = html.escape(str(cfg.get("title") or cfg.get("file_name") or "صوت الموقع"))
    loop_flag = bool(cfg.get("loop", True))
    vol = float(cfg.get("volume", 0.7))

    if cfg.get("source_type") == "upload":
        src = f"data:{cfg.get('mime_type', 'audio/mpeg')};base64,{cfg.get('audio_b64')}"
    else:
        src = str(cfg.get("audio_url") or "").strip()

    if not src:
        return

    with st.sidebar:
        st.markdown(f"🔊 **صوت الموقع:** {title}")
        if cfg.get("source_type") == "upload":
            try:
                raw_bytes = base64.b64decode(cfg["audio_b64"])
                st.audio(raw_bytes, format=cfg.get("mime_type", "audio/mpeg"), loop=loop_flag)
            except Exception:
                pass
        else:
            st.audio(src, loop=loop_flag)

        st.components.v1.html(
            f"""
            <script>
            (function() {{
                const a = new Audio("{src}");
                a.loop = {str(loop_flag).lower()};
                a.volume = {vol};
                a.play().catch(() => {{}});
            }})();
            </script>
            """,
            height=0,
        )


def render_admin_audio_manager(actor_username="admin"):
    """لوحة تحكم الأدمن لإضافة الصوت، تغييره، إيقافه/تشغيله، أو إلغائه."""
    st.markdown("---")
    st.markdown("<h3 style='color:#D4AF37;text-align:center;'>🎵 إدارة صوت الموقع أثناء التشغيل</h3>", unsafe_allow_html=True)
    cfg = get_site_audio_config(force_reload=True)

    has_audio = _has_valid_site_audio(cfg)
    is_running = bool(cfg.get("enabled")) and has_audio

    if has_audio:
        state_lbl = "🟢 يعمل الآن" if is_running else "⏸️ متوقف"
        src_lbl = (
            f"ملف مرفوع: {cfg.get('file_name') or 'ملف صوتي'}"
            if cfg.get("source_type") == "upload"
            else f"رابط صوتي: {cfg.get('audio_url')}"
        )
        st.info(
            f"**الحالة:** {state_lbl}  \n"
            f"**العنوان:** {cfg.get('title') or 'بدون عنوان'}  \n"
            f"**المصدر:** {src_lbl}"
        )
    else:
        st.warning("لا يوجد صوت مضاف للموقع حاليًا.")

    c_stop, c_play, c_cancel = st.columns(3)
    with c_stop:
        if st.button("⏸️ إيقاف الصوت", key="admin_audio_stop_btn", use_container_width=True, disabled=not has_audio):
            cfg["enabled"] = False
            cfg["updated_by"] = actor_username
            if save_site_audio_config(cfg):
                st.success("تم إيقاف صوت الموقع.")
                st.rerun()
    with c_play:
        if st.button("▶️ تشغيل الصوت", key="admin_audio_play_btn", type="primary", use_container_width=True, disabled=not has_audio):
            cfg["enabled"] = True
            cfg["updated_by"] = actor_username
            if save_site_audio_config(cfg):
                st.success("تم تشغيل صوت الموقع.")
                st.rerun()
    with c_cancel:
        if st.button("🗑️ إلغاء وحذف الصوت", key="admin_audio_cancel_btn", use_container_width=True, disabled=not (has_audio or cfg.get("enabled"))):
            reset_cfg = dict(DEFAULT_SITE_AUDIO_CONFIG)
            reset_cfg["updated_by"] = actor_username
            if save_site_audio_config(reset_cfg):
                st.success("تم إلغاء وحذف صوت الموقع.")
                st.rerun()

    with st.form("admin_site_audio_form"):
        form_heading = "🔄 تغيير الصوت الحالي" if has_audio else "➕ إضافة صوت للموقع"
        st.markdown(f"#### {form_heading}")
        audio_title = st.text_input("اسم أو وصف الصوت", value=str(cfg.get("title") or ""), key="admin_audio_title_inp")
        source_mode = st.radio(
            "مصدر الصوت",
            ["ملف صوتي مرفوع", "رابط صوتي مباشر"],
            index=0 if cfg.get("source_type") != "url" else 1,
            horizontal=True,
            key="admin_audio_src_mode",
        )

        uploaded_audio = None
        audio_url_val = ""
        if source_mode == "ملف صوتي مرفوع":
            uploaded_audio = st.file_uploader(
                "اختر ملف الصوت (MP3, WAV, OGG, M4A, AAC, WEBM)",
                type=["mp3", "wav", "ogg", "m4a", "aac", "webm"],
                key="admin_audio_file_uploader",
            )
        else:
            audio_url_val = st.text_input(
                "رابط الصوت المباشر",
                value=str(cfg.get("audio_url") or ""),
                key="admin_audio_url_inp",
            )

        f1, f2, f3 = st.columns(3)
        with f1:
            enable_audio = st.checkbox("تشغيل الصوت فور الحفظ", value=bool(cfg.get("enabled", True)), key="admin_audio_enable_chk")
        with f2:
            loop_audio = st.checkbox("تكرار الصوت تلقائيًا", value=bool(cfg.get("loop", True)), key="admin_audio_loop_chk")
        with f3:
            vol_audio = st.slider("مستوى الصوت", min_value=0.0, max_value=1.0, value=float(cfg.get("volume", 0.7)), step=0.05, key="admin_audio_vol_sld")

        if st.form_submit_button("💾 حفظ إعدادات الصوت", type="primary", use_container_width=True):
            updated_audio = dict(cfg)
            updated_audio["title"] = audio_title.strip()
            updated_audio["loop"] = bool(loop_audio)
            updated_audio["volume"] = float(vol_audio)
            updated_audio["updated_by"] = actor_username

            if source_mode == "ملف صوتي مرفوع":
                if uploaded_audio is not None:
                    raw_bytes = uploaded_audio.getvalue()
                    updated_audio["source_type"] = "upload"
                    updated_audio["file_name"] = uploaded_audio.name
                    updated_audio["mime_type"] = _guess_site_audio_mime(uploaded_audio.name, uploaded_audio.type or "")
                    updated_audio["audio_b64"] = base64.b64encode(raw_bytes).decode("utf-8")
                    updated_audio["audio_url"] = ""
                elif cfg.get("source_type") != "upload" or not cfg.get("audio_b64"):
                    st.error("يرجى رفع ملف صوتي أولًا.")
                    st.stop()
            else:
                clean_u = audio_url_val.strip()
                if clean_u:
                    updated_audio["source_type"] = "url"
                    updated_audio["audio_url"] = clean_u
                    updated_audio["audio_b64"] = ""
                    updated_audio["file_name"] = ""
                elif cfg.get("source_type") != "url" or not cfg.get("audio_url"):
                    st.error("يرجى إدخال رابط صوتي صحيح.")
                    st.stop()

            updated_audio["enabled"] = bool(enable_audio)
            if save_site_audio_config(updated_audio):
                st.success("✅ تم حفظ إعدادات صوت الموقع بنجاح.")
                st.rerun()

JOB_TITLES_FILE = DATA_DIR / "job_titles.json"

DEFAULT_JOB_TITLES = [
    "عضو قانونى",
    "مدير إدارة",
    "مدير عام",
    "قيادي مركز رئيسى 1",
    "قيادي مركز رئيسى 2",
    "قيادي مركز رئيسى 3",
]

def load_job_titles():
    raw = load_json_persistent("job_titles", JOB_TITLES_FILE, list(DEFAULT_JOB_TITLES))
    if not isinstance(raw, list):
        raw = list(DEFAULT_JOB_TITLES)
    seen = set()
    out = []
    for item in DEFAULT_JOB_TITLES + [str(x or "").strip() for x in raw]:
        s = str(item or "").strip()
        if s and s not in seen:
            seen.add(s)
            out.append(s)
    return out

def save_job_titles(titles):
    seen = set()
    cleaned = []
    for item in (titles or []):
        s = str(item or "").strip()
        if s and s not in seen:
            seen.add(s)
            cleaned.append(s)
    for d in DEFAULT_JOB_TITLES:
        if d not in seen:
            seen.add(d)
            cleaned.append(d)
    ok = save_json_persistent("job_titles", JOB_TITLES_FILE, cleaned)
    try:
        _cached_job_titles_snapshot.clear()
    except Exception:
        pass
    return ok

@st.cache_data(ttl=15, show_spinner=False)
def _cached_job_titles_snapshot():
    return load_job_titles()

def get_all_job_titles(include_users=True):
    titles = list(_cached_job_titles_snapshot())
    seen = set(titles)
    if include_users:
        try:
            for u in _cached_users_snapshot():
                for key_name in ("job_title", "job", "وظيفة", "role_name"):
                    val = str(u.get(key_name) or "").strip()
                    if val and val not in seen:
                        seen.add(val)
                        titles.append(val)
        except Exception:
            pass
    return titles

def register_job_title(title_text):
    clean = str(title_text or "").strip()
    if not clean:
        return False, "أدخل اسم الوظيفة أولاً"
    current = get_all_job_titles(include_users=True)
    if clean in current:
        return True, "الوظيفة موجودة بالفعل ومتاحة للجميع"
    current.append(clean)
    if save_job_titles(current):
        return True, f"تمت إضافة الوظيفة ({clean}) وأصبحت متاحة للأدمن ولكل الأعضاء"
    return False, "تعذر حفظ الوظيفة"

DEFAULT_TOP_BAR_CONFIG = {
    "enabled": True,
    "text": "مع تحيات وليد حماد - الإدارة العامة للشئون القانونية بديوان عام منطقة البحيرة بالهيئة القومية للتأمين الاجتماعي",
    "bg_start": "#D4AF37",
    "bg_mid": "#FFD700",
    "bg_end": "#D4AF37",
    "text_color": "#0A1428",
    "font_size": 16,
    "speed_seconds": 15,
    "mode": "marquee",
}

def load_top_bar_config():
    cfg = load_json_persistent("top_bar_config", TOP_BAR_FILE, dict(DEFAULT_TOP_BAR_CONFIG))
    if not isinstance(cfg, dict):
        cfg = dict(DEFAULT_TOP_BAR_CONFIG)
    merged = dict(DEFAULT_TOP_BAR_CONFIG)
    for k, v in cfg.items():
        if v is not None:
            merged[k] = v
    merged["enabled"] = bool(merged.get("enabled", True))
    try:
        merged["font_size"] = max(11, min(36, int(merged.get("font_size", 16))))
    except Exception:
        merged["font_size"] = 16
    try:
        merged["speed_seconds"] = max(4, min(120, int(merged.get("speed_seconds", 15))))
    except Exception:
        merged["speed_seconds"] = 15
    if merged.get("mode") not in ("marquee", "static"):
        merged["mode"] = "marquee"
    return merged

def save_top_bar_config(cfg):
    merged = dict(DEFAULT_TOP_BAR_CONFIG)
    if isinstance(cfg, dict):
        for k, v in cfg.items():
            if v is not None:
                merged[k] = v
    merged["enabled"] = bool(merged.get("enabled", True))
    try:
        merged["font_size"] = max(11, min(36, int(merged.get("font_size", 16))))
    except Exception:
        merged["font_size"] = 16
    try:
        merged["speed_seconds"] = max(4, min(120, int(merged.get("speed_seconds", 15))))
    except Exception:
        merged["speed_seconds"] = 15
    if merged.get("mode") not in ("marquee", "static"):
        merged["mode"] = "marquee"
    ok = save_json_persistent("top_bar_config", TOP_BAR_FILE, merged)
    try:
        _cached_top_bar_snapshot.clear()
    except Exception:
        pass
    return ok

@st.cache_data(ttl=15, show_spinner=False)
def _cached_top_bar_snapshot():
    return load_top_bar_config()

def render_top_bar():
    cfg = _cached_top_bar_snapshot()
    if not cfg.get("enabled", True):
        return
    txt = str(cfg.get("text") or "").strip()
    if not txt:
        return
    bg1 = html.escape(str(cfg.get("bg_start") or "#D4AF37"))
    bg2 = html.escape(str(cfg.get("bg_mid") or "#FFD700"))
    bg3 = html.escape(str(cfg.get("bg_end") or "#D4AF37"))
    fg = html.escape(str(cfg.get("text_color") or "#0A1428"))
    fs = int(cfg.get("font_size", 16))
    spd = int(cfg.get("speed_seconds", 15))
    mode = cfg.get("mode", "marquee")
    safe_txt = html.escape(txt)
    if mode == "static":
        st.markdown(
            f"""<div style="background:linear-gradient(90deg,{bg1},{bg2},{bg3});color:{fg};padding:12px 16px;font-weight:900;font-size:{fs}px;text-align:center;border-radius:0 0 15px 15px;box-shadow:0 4px 14px rgba(0,0,0,.25);overflow-wrap:break-word;">{safe_txt}</div>""",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"""<div class="marquee" style="background:linear-gradient(90deg,{bg1},{bg2},{bg3});color:{fg};font-size:{fs}px;"><span style="animation:marquee {spd}s linear infinite;">{safe_txt}</span></div>""",
            unsafe_allow_html=True,
        )

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

try:
    from supabase import create_client as _create_supabase_client
    supabase = _create_supabase_client(SUPABASE_URL, SUPABASE_KEY) if CLOUD_ENABLED else None
except Exception:
    supabase = None

_CLOUD_LAST_ERROR = ""

def cloud_status():
    try:
        return bool(SUPABASE_URL and SUPABASE_KEY)
    except Exception:
        return False

def cloud_success_message(action, section=""):
    target = f" في {section}" if section else ""
    return f"تم {action}{target} والسحابة بنجاح ✓"

def cloud_error_message(action, error):
    return f"تم {action} محليًا، ولكن تعذر إتمام العملية في السحابة: {error}"

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
            _CLOUD_LAST_ERROR = ""
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
            _CLOUD_LAST_ERROR = ""
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
            _CLOUD_LAST_ERROR = ""
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
            _CLOUD_LAST_ERROR = ""
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

def cloud_delete_record(table_name, record_id):
    try:
        if not cloud_status() or not record_id:
            return False, "اتصال السحابة غير متاح"
        if supabase is not None:
            supabase.table(table_name).delete().eq("id", record_id).execute()
            return True, ""
        r = requests.delete(
            f"{SUPABASE_URL}/rest/v1/{table_name}",
            params={"id": f"eq.{record_id}"},
            headers=_cloud_headers(),
            timeout=15,
        )
        return (r.ok, "" if r.ok else r.text[:300])
    except Exception as exc:
        return False, str(exc)

def cloud_upload_file(file_bytes, file_name, content_type="application/octet-stream"):
    try:
        if not cloud_status():
            return None, "اتصال السحابة غير متاح"
        safe_name = re.sub(r"[^A-Za-z0-9._/-]+", "_", str(file_name))
        storage_path = f"{datetime.now().strftime('%Y/%m')}/{safe_name}"
        if supabase is not None:
            supabase.storage.from_(SUPABASE_STORAGE_BUCKET).upload(
                storage_path, file_bytes,
                {"content-type": content_type, "upsert": "true"}
            )
            return storage_path, ""
        r = requests.post(
            f"{SUPABASE_URL}/storage/v1/object/{SUPABASE_STORAGE_BUCKET}/{storage_path}",
            headers=_storage_headers(content_type),
            data=file_bytes,
            timeout=25,
        )
        if r.ok:
            return storage_path, ""
        return None, f"HTTP {r.status_code}: {r.text[:300]}"
    except Exception as exc:
        return None, str(exc)

def cloud_delete_file(storage_path):
    try:
        if not cloud_status() or not storage_path:
            return False, "مسار الملف السحابي غير موجود"
        if supabase is not None:
            supabase.storage.from_(SUPABASE_STORAGE_BUCKET).remove([storage_path])
            return True, ""
        r = requests.delete(
            f"{SUPABASE_URL}/storage/v1/object/{SUPABASE_STORAGE_BUCKET}/{storage_path}",
            headers={
                "apikey": SUPABASE_KEY,
                "Authorization": f"Bearer {SUPABASE_KEY}",
            },
            timeout=15,
        )
        return (r.ok, "" if r.ok else r.text[:300])
    except Exception as exc:
        return False, str(exc)

def fix_arabic(text):
    """ بتظبط العربي عشان ميطلعش متقطع """
    if not text:
        return ""
    reshaped_text = arabic_reshaper.reshape(str(text))
    bidi_text = get_display(reshaped_text)
    return bidi_text

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
            elif isinstance(cloud_obj, (dict, list)) and isinstance(local_obj, (dict, list)):
                cloud_updated = str(cloud.get("updated_at") or "")
                if cloud_updated and not local_obj:
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

def sync_all_local_to_cloud():
    """مزامنة جميع ملفات النظام المحلية إلى السحابة."""
    if not CLOUD_ENABLED:
        return False, "إعدادات SUPABASE_URL و SUPABASE_KEY غير مفعلة."
    pairs = [
        ("users", Path(os.path.join(LOCAL_DATA_DIR, "users.json"))),
        ("data", Path(os.path.join(LOCAL_DATA_DIR, "cases_data.json"))),
        ("messages", Path(os.path.join(LOCAL_DATA_DIR, "messages_v1.json"))),
        ("banners", BANNERS_FILE),
        ("top_bar_config", TOP_BAR_FILE),
        ("site_audio_config", SITE_AUDIO_FILE),
        ("job_titles", JOB_TITLES_FILE),
        ("tokens", Path(os.path.join(LOCAL_DATA_DIR, "tokens.json"))),
    ]
    synced = 0
    errors = []
    for k, p in pairs:
        obj = _read_local_json_safe(p)
        if obj is not None:
            if _cloud_put(k, obj):
                synced += 1
            else:
                errors.append(f"{k}: {_CLOUD_LAST_ERROR}")
    if errors:
        return False, f"تمت مزامنة {synced} ملفات، وحدثت ملاحظات: {' | '.join(errors)}"
    return True, f"تمت مزامنة {synced} ملفات مع السحابة بنجاح ✓"

def pull_all_cloud_to_local():
    """سحب أحدث البيانات من السحابة وتحديث الملفات المحلية."""
    if not CLOUD_ENABLED:
        return False, "إعدادات SUPABASE_URL و SUPABASE_KEY غير مفعلة."
    pairs = [
        ("users", Path(os.path.join(LOCAL_DATA_DIR, "users.json"))),
        ("data", Path(os.path.join(LOCAL_DATA_DIR, "cases_data.json"))),
        ("messages", Path(os.path.join(LOCAL_DATA_DIR, "messages_v1.json"))),
        ("banners", BANNERS_FILE),
        ("top_bar_config", TOP_BAR_FILE),
        ("site_audio_config", SITE_AUDIO_FILE),
        ("job_titles", JOB_TITLES_FILE),
        ("tokens", Path(os.path.join(LOCAL_DATA_DIR, "tokens.json"))),
    ]
    pulled = 0
    for k, p in pairs:
        row = _cloud_get(k)
        if row and row.get("payload") is not None:
            _write_local_json_atomic(p, row["payload"])
            pulled += 1
    try:
        st.cache_data.clear()
    except Exception:
        pass
    return True, f"تم تحديث {pulled} ملفات من السحابة بنجاح ✓"

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
    if not isinstance(user, dict):
        return ""
    job = str(user.get("job_title") or user.get("job") or user.get("وظيفة") or user.get("role_name") or "").strip()
    return "مدير عام (قيادى مناطق)" if job == "مدير عام" else job

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
                st.markdown(f"<div style='background:#14233d;border:2px solid #D4AF37;border-radius:14px;padding:14px;margin-top:14px;color:#fff;text-align:center;font-weight:900;'>تم العثور على العضوية: {html.escape(str(target.get('display_name') or target.get('username')))}<br>أدخل البريد الإلكتروني لاسترداد الحساب وكلمة السر التي تريد استخدامها.</div>", unsafe_allow_html=True)
                recovery_email_act = st.text_input("البريد الإلكتروني لاسترداد الحساب (اختياري)", value=target.get("email") or "", placeholder="example@domain.com", key="activation_recovery_email")
                first = st.text_input("كلمة السر", type="password", key="activation_pass_1")
                second = st.text_input("تأكيد كلمة السر", type="password", key="activation_pass_2")
                activation_photo = st.file_uploader("الصورة الشخصية", type=["jpg","jpeg","png"], key="activation_profile_photo")
                if st.button("💾 حفظ كلمة السر وتفعيل العضوية", type="primary", use_container_width=True, key="confirm_activation_v2"):
                    cleaned_email = recovery_email_act.strip()
                    if not first:
                        st.error("اكتب كلمة السر أولاً")
                    elif first != second:
                        st.error("كلمتا السر غير متطابقتين")
                    elif cleaned_email and ("@" not in cleaned_email or "." not in cleaned_email.split("@")[-1]):
                        st.error("أدخل بريدًا إلكترونيًا صحيحًا لاسترداد الحساب")
                    else:
                        for u in users:
                            if u.get("id") == activation_id:
                                u["password"] = bcrypt.hashpw(first.encode(), bcrypt.gensalt()).decode()
                                u["password_set"] = True
                                if cleaned_email:
                                    u["email"] = cleaned_email
                                    u["recovery_email"] = cleaned_email
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

    render_return_home_button("⬅️ العودة للرئيسية", key="extract_member_home")
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
        all_jobs_extract = get_all_job_titles(include_users=True)
        if st.session_state.get("user", {}).get("role") == "admin":
            with st.expander("➕ إضافة وظيفة جديدة تظهر للأدمن ولكل الأعضاء"):
                c_jt1, c_jt2 = st.columns([3, 1])
                with c_jt1:
                    new_jt_inp = st.text_input("اسم الوظيفة الجديدة", placeholder="مثال: باحث قانوني أول / مدير عام الشؤون القانونية", key="quick_add_job_title_inp")
                with c_jt2:
                    st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
                    if st.button("إضافة الوظيفة", key="quick_add_job_title_btn", use_container_width=True):
                        ok_jt, msg_jt = register_job_title(new_jt_inp)
                        if ok_jt:
                            st.success(msg_jt)
                            st.rerun()
                        else:
                            st.error(msg_jt)
        job_role = st.selectbox("الوظيفة", all_jobs_extract, key="new_member_job_role")
        phone_number = st.text_input("رقم الموبايل (اختياري)", key="new_member_phone")
        profile_photo = st.file_uploader("الصورة الشخصية", type=["jpg","jpeg","png"], key="new_member_profile_photo")

        if st.button("🪪 استخراج العضوية", use_container_width=True, type="primary", key="extract_member_new_v3"):
            name = new_username.strip()
            area = insurance_area.strip()
            mobile_val = phone_number.strip()
            if not name:
                st.error("لازم تكتب اسم العضو")
            elif not area:
                st.error("لازم تكتب المنطقة التأمينية")
            else:
                users = _cached_users_snapshot()
                existing_user = next((u for u in users if u.get("username") == name or u.get("display_name") == name), None)
                birth_date_str = fmt_date_ar(birth_date)
                role_val = {"عضو قانونى":"member", "مدير إدارة":"manager", "مدير عام":"general_manager", "قيادي مركز رئيسى 1":"leader1", "قيادي مركز رئيسى 2":"leader2", "قيادي مركز رئيسى 3":"leader3"}.get(job_role, "member")
                if existing_user:
                    existing_user["display_name"] = existing_user.get("display_name") or name
                    existing_user["birth_date"] = birth_date_str
                    existing_user["insurance_area"] = area
                    register_job_title(job_role)
                    existing_user["job_title"] = job_role
                    existing_user["phone"] = mobile_val
                    existing_user["mobile"] = mobile_val
                    if profile_photo:
                        existing_user["profile_photo"] = base64.b64encode(profile_photo.getvalue()).decode("utf-8")
                        existing_user["profile_photo_type"] = profile_photo.type
                    existing_user["role"] = role_val
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
                        "display_name": name,
                        "password": "",
                        "email": "",
                        "recovery_email": "",
                        "role": role_val,
                        "job_title": job_role,
                        "phone": mobile_val,
                        "mobile": mobile_val,
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
    cur_u = st.session_state.get("user") or {}
    if cur_u.get("role") != "admin" and not is_assistant_admin(cur_u):
        st.error("هذه الصفحة متاحة للادمن فقط")
        return
    if not is_primary_admin(cur_u):
        st.info(f"أنت تعمل بصلاحيات الأدمن المساعد ({html.escape(get_assistant_admin_title(cur_u))}). لا يمكن حذف أو إيقاف أو تعديل عضوية الأدمن 1 أو إغلاق الموقع.")

    st.markdown("<h2 style='color:#D4AF37;text-align:center;'>👥 إدارة الأعضاء</h2>", unsafe_allow_html=True)
    render_return_home_button("⬅️ العودة للرئيسية", key="manage_users_top_home")
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

            with st.expander("🔑 استخراج / تعيين كلمة سر جديدة للعضو (إذا نسيها)"):
                st.caption("يحتفظ العضو بنفس عضويته وبياناته وقضاياه، ويمكنه بعد الدخول تغيير كلمة السر من لوحة تحكمه.")
                admin_new_pass = st.text_input("كلمة السر الجديدة للعضو", type="password", key=f"admin_set_pass_inp_{user['id']}")
                admin_confirm_pass = st.text_input("تأكيد كلمة السر الجديدة", type="password", key=f"admin_set_pass_conf_{user['id']}")
                if st.button("🔑 حفظ كلمة السر الجديدة للعضو", key=f"admin_set_pass_btn_{user['id']}", type="primary", use_container_width=True):
                    if not admin_new_pass:
                        st.error("اكتب كلمة السر الجديدة أولاً")
                    elif admin_new_pass != admin_confirm_pass:
                        st.error("كلمتا السر غير متطابقتين")
                    else:
                        user["password"] = bcrypt.hashpw(admin_new_pass.encode(), bcrypt.gensalt()).decode()
                        user["password_set"] = True
                        if user.get("status") in ("suspended_temp", "suspended_final"):
                            user["status"] = "active"
                        save_users(users)
                        st.success(f"تم استخراج وتعيين كلمة سر جديدة للعضو {user.get('display_name') or user.get('username')} بنفس عضويته وبياناته.")
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

DB_PATH = os.environ.get("APP_DB_PATH", str(Path(LOCAL_DATA_DIR) / "app_permissions_system.db"))

ROLES = {
    "admin": "الادمن",
    "member": "عضو قانونى",
    "manager": "مدير إدارة",
    "general_manager": "مدير عام",
    "leader1": "قيادي مركز رئيسى 1",
    "leader2": "قيادي مركز رئيسى 2",
    "leader3": "قيادي مركز رئيسى 3",
}

ALL_ROLES = list(ROLES.keys())

DEFAULT_ROLE_PERMISSIONS = {
    "admin": {
        "can_send_messages": True,
        "message_scope": "all_users",
        "allowed_message_roles": ALL_ROLES,
        "can_manage_permissions": True,
        "can_view_cases": True,
        "can_edit_cases": True,
        "can_delete_cases": True,
        "can_export_reports": True,
        "can_manage_users": True,
    },
    "member": {
        "can_send_messages": True,
        "message_scope": "roles",
        "allowed_message_roles": ["admin", "member", "manager"],
        "can_manage_permissions": False,
        "can_view_cases": True,
        "can_edit_cases": True,
        "can_delete_cases": False,
        "can_export_reports": True,
        "can_manage_users": False,
    },
    "manager": {
        "can_send_messages": True,
        "message_scope": "roles",
        "allowed_message_roles": ["admin", "member", "manager", "general_manager"],
        "can_manage_permissions": False,
        "can_view_cases": True,
        "can_edit_cases": True,
        "can_delete_cases": False,
        "can_export_reports": True,
        "can_manage_users": False,
    },
    "general_manager": {
        "can_send_messages": True,
        "message_scope": "roles",
        "allowed_message_roles": ["admin", "member", "manager", "general_manager", "leader2", "leader3"],
        "can_manage_permissions": False,
        "can_view_cases": True,
        "can_edit_cases": True,
        "can_delete_cases": True,
        "can_export_reports": True,
        "can_manage_users": False,
    },
    "leader1": {
        "can_send_messages": True,
        "message_scope": "roles",
        "allowed_message_roles": ["admin", "member", "manager", "general_manager", "leader1", "leader2", "leader3"],
        "can_manage_permissions": False,
        "can_view_cases": True,
        "can_edit_cases": True,
        "can_delete_cases": True,
        "can_export_reports": True,
        "can_manage_users": False,
    },
    "leader2": {
        "can_send_messages": True,
        "message_scope": "roles",
        "allowed_message_roles": ["admin", "member", "manager", "general_manager", "leader1", "leader2", "leader3"],
        "can_manage_permissions": False,
        "can_view_cases": True,
        "can_edit_cases": True,
        "can_delete_cases": True,
        "can_export_reports": True,
        "can_manage_users": False,
    },
    "leader3": {
        "can_send_messages": True,
        "message_scope": "roles",
        "allowed_message_roles": ["admin", "member", "manager", "general_manager", "leader2", "leader3"],
        "can_manage_permissions": False,
        "can_view_cases": True,
        "can_edit_cases": True,
        "can_delete_cases": True,
        "can_export_reports": True,
        "can_manage_users": False,
    },
}

PERMISSION_LABELS = {
    "can_send_messages": "السماح بإرسال الرسائل",
    "can_manage_permissions": "إدارة وتعديل صلاحيات الأعضاء",
    "can_view_cases": "عرض القضايا والملفات",
    "can_edit_cases": "إضافة وتعديل البيانات والقضايا",
    "can_delete_cases": "حذف السجلات والقضايا",
    "can_export_reports": "تصدير التقارير",
    "can_manage_users": "إدارة حسابات الأعضاء",
}

SCOPE_LABELS = {
    "all_users": "جميع الأعضاء النشطين",
    "roles": "حسب الأدوار الوظيفية المحددة",
    "custom_users": "أعضاء محددون بالاسم",
}


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def get_connection():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            full_name TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL,
            phone TEXT DEFAULT "",
            photo_path TEXT DEFAULT "",
            is_active INTEGER DEFAULT 1,
            created_at TEXT NOT NULL
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS user_permissions (
            user_id INTEGER PRIMARY KEY,
            can_send_messages INTEGER NOT NULL DEFAULT 1,
            message_scope TEXT NOT NULL DEFAULT "roles",
            allowed_message_roles TEXT NOT NULL DEFAULT "[]",
            allowed_message_users TEXT NOT NULL DEFAULT "[]",
            can_manage_permissions INTEGER NOT NULL DEFAULT 0,
            can_view_cases INTEGER NOT NULL DEFAULT 1,
            can_edit_cases INTEGER NOT NULL DEFAULT 1,
            can_delete_cases INTEGER NOT NULL DEFAULT 0,
            can_export_reports INTEGER NOT NULL DEFAULT 1,
            can_manage_users INTEGER NOT NULL DEFAULT 0,
            updated_at TEXT NOT NULL,
            FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sender_id INTEGER NOT NULL,
            receiver_id INTEGER NOT NULL,
            subject TEXT DEFAULT "",
            body TEXT NOT NULL,
            created_at TEXT NOT NULL,
            is_read INTEGER DEFAULT 0,
            FOREIGN KEY (sender_id) REFERENCES users (id),
            FOREIGN KEY (receiver_id) REFERENCES users (id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS permission_audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            actor_id INTEGER NOT NULL,
            target_scope_desc TEXT NOT NULL,
            details_json TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY (actor_id) REFERENCES users (id)
        )
        """
    )

    seed_users = [
        ("admin", "مدير النظام (الأدمن)", "Admin@123", "admin", "01000000000"),
        ("member1", "أ/ عضو قانوني أول", "Member@123", "member", "01011111111"),
        ("manager1", "أ/ مدير الإدارة", "Manager@123", "manager", "01022222222"),
        ("gm1", "أ/ المدير العام", "Gm@123", "general_manager", "01033333333"),
        ("leader1", "قيادي مركز رئيسي 1", "Leader@123", "leader1", "01044444444"),
        ("leader2", "قيادي مركز رئيسي 2", "Leader@123", "leader2", "01055555555"),
        ("leader3", "قيادي مركز رئيسي 3", "Leader@123", "leader3", "01066666666"),
    ]

    now = datetime.now().isoformat(timespec="seconds")
    for username, full_name, raw_pw, role, phone in seed_users:
        cur.execute("SELECT id FROM users WHERE username = ?", (username,))
        row = cur.fetchone()
        if not row:
            cur.execute(
                """
                INSERT INTO users (username, full_name, password_hash, role, phone, is_active, created_at)
                VALUES (?, ?, ?, ?, ?, 1, ?)
                """,
                (username, full_name, hash_password(raw_pw), role, phone, now),
            )
            user_id = cur.lastrowid
        else:
            user_id = row["id"]

        cur.execute("SELECT user_id FROM user_permissions WHERE user_id = ?", (user_id,))
        if not cur.fetchone():
            defaults = DEFAULT_ROLE_PERMISSIONS.get(role, DEFAULT_ROLE_PERMISSIONS["member"])
            cur.execute(
                """
                INSERT INTO user_permissions (
                    user_id, can_send_messages, message_scope, allowed_message_roles,
                    allowed_message_users, can_manage_permissions, can_view_cases,
                    can_edit_cases, can_delete_cases, can_export_reports, can_manage_users, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    user_id,
                    int(defaults["can_send_messages"]),
                    defaults["message_scope"],
                    json.dumps(defaults["allowed_message_roles"], ensure_ascii=False),
                    json.dumps([], ensure_ascii=False),
                    int(defaults["can_manage_permissions"]),
                    int(defaults["can_view_cases"]),
                    int(defaults["can_edit_cases"]),
                    int(defaults["can_delete_cases"]),
                    int(defaults["can_export_reports"]),
                    int(defaults["can_manage_users"]),
                    now,
                ),
            )

    conn.commit()
    conn.close()


def get_all_users(include_inactive: bool = True):
    conn = get_connection()
    cur = conn.cursor()
    if include_inactive:
        cur.execute("SELECT * FROM users ORDER BY id ASC")
    else:
        cur.execute("SELECT * FROM users WHERE is_active = 1 ORDER BY id ASC")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def get_user_by_id(user_id: int):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM users WHERE id = ?", (user_id,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def get_user_permissions(user_id: int) -> dict:
    user = get_user_by_id(user_id)
    role = user["role"] if user else "member"
    defaults = DEFAULT_ROLE_PERMISSIONS.get(role, DEFAULT_ROLE_PERMISSIONS["member"]).copy()
    defaults["allowed_message_roles"] = list(defaults.get("allowed_message_roles", []))
    defaults["allowed_message_users"] = []

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM user_permissions WHERE user_id = ?", (user_id,))
    row = cur.fetchone()
    conn.close()

    if not row:
        return defaults

    try:
        allowed_roles = json.loads(row["allowed_message_roles"] or "[]")
    except Exception:
        allowed_roles = defaults["allowed_message_roles"]

    try:
        allowed_users = [int(x) for x in json.loads(row["allowed_message_users"] or "[]")]
    except Exception:
        allowed_users = []

    return {
        "can_send_messages": bool(row["can_send_messages"]),
        "message_scope": row["message_scope"] or "roles",
        "allowed_message_roles": allowed_roles,
        "allowed_message_users": allowed_users,
        "can_manage_permissions": bool(row["can_manage_permissions"]),
        "can_view_cases": bool(row["can_view_cases"]),
        "can_edit_cases": bool(row["can_edit_cases"]),
        "can_delete_cases": bool(row["can_delete_cases"]),
        "can_export_reports": bool(row["can_export_reports"]),
        "can_manage_users": bool(row["can_manage_users"]),
        "updated_at": row["updated_at"],
    }


def upsert_user_permissions(user_id: int, perms: dict):
    now = datetime.now().isoformat(timespec="seconds")
    conn = get_connection()
    cur = conn.cursor()

    allowed_roles_json = json.dumps(list(perms.get("allowed_message_roles", [])), ensure_ascii=False)
    allowed_users_json = json.dumps([int(u) for u in perms.get("allowed_message_users", [])], ensure_ascii=False)

    cur.execute(
        """
        INSERT INTO user_permissions (
            user_id, can_send_messages, message_scope, allowed_message_roles,
            allowed_message_users, can_manage_permissions, can_view_cases,
            can_edit_cases, can_delete_cases, can_export_reports, can_manage_users, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(user_id) DO UPDATE SET
            can_send_messages = excluded.can_send_messages,
            message_scope = excluded.message_scope,
            allowed_message_roles = excluded.allowed_message_roles,
            allowed_message_users = excluded.allowed_message_users,
            can_manage_permissions = excluded.can_manage_permissions,
            can_view_cases = excluded.can_view_cases,
            can_edit_cases = excluded.can_edit_cases,
            can_delete_cases = excluded.can_delete_cases,
            can_export_reports = excluded.can_export_reports,
            can_manage_users = excluded.can_manage_users,
            updated_at = excluded.updated_at
        """,
        (
            user_id,
            int(bool(perms.get("can_send_messages", False))),
            perms.get("message_scope", "roles"),
            allowed_roles_json,
            allowed_users_json,
            int(bool(perms.get("can_manage_permissions", False))),
            int(bool(perms.get("can_view_cases", True))),
            int(bool(perms.get("can_edit_cases", False))),
            int(bool(perms.get("can_delete_cases", False))),
            int(bool(perms.get("can_export_reports", False))),
            int(bool(perms.get("can_manage_users", False))),
            now,
        ),
    )
    conn.commit()
    conn.close()


def apply_permissions_bulk(target_user_ids: list[int], perms: dict, actor_id: int, scope_desc: str):
    for uid in target_user_ids:
        user = get_user_by_id(uid)
        if not user:
            continue
        effective_perms = perms.copy()
        if user["role"] == "admin":
            effective_perms["can_manage_permissions"] = True
            effective_perms["can_send_messages"] = True
            effective_perms["message_scope"] = "all_users"
        upsert_user_permissions(uid, effective_perms)

    now = datetime.now().isoformat(timespec="seconds")
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO permission_audit_log (actor_id, target_scope_desc, details_json, created_at)
        VALUES (?, ?, ?, ?)
        """,
        (
            actor_id,
            scope_desc,
            json.dumps({"target_count": len(target_user_ids), "permissions": perms}, ensure_ascii=False),
            now,
        ),
    )
    conn.commit()
    conn.close()


def reset_user_to_role_defaults(user_id: int, actor_id: int):
    user = get_user_by_id(user_id)
    if not user:
        return
    defaults = DEFAULT_ROLE_PERMISSIONS.get(user["role"], DEFAULT_ROLE_PERMISSIONS["member"]).copy()
    defaults["allowed_message_users"] = []
    apply_permissions_bulk(
        [user_id],
        defaults,
        actor_id=actor_id,
        scope_desc=f"استعادة الصلاحيات الافتراضية للعضو: {user['full_name']} ({user['username']})",
    )


def has_permission(user_id: int, perm_key: str) -> bool:
    user = get_user_by_id(user_id)
    if not user or not user["is_active"]:
        return False
    if user["role"] == "admin" and perm_key == "can_manage_permissions":
        return True
    perms = get_user_permissions(user_id)
    return bool(perms.get(perm_key, False))


def can_message_db(sender_id: int, target_id: int) -> bool:
    if sender_id == target_id:
        return False
    sender = get_user_by_id(sender_id)
    target = get_user_by_id(target_id)
    if not sender or not target:
        return False
    if not sender["is_active"] or not target["is_active"]:
        return False
    if sender["role"] == "admin":
        return True

    perms = get_user_permissions(sender_id)
    if not perms.get("can_send_messages", False):
        return False

    scope = perms.get("message_scope", "roles")
    if scope == "all_users":
        return True
    if scope == "roles":
        return target["role"] in set(perms.get("allowed_message_roles", []))
    if scope == "custom_users":
        return int(target_id) in set(int(x) for x in perms.get("allowed_message_users", []))
    return False


def get_allowed_recipients(sender_id: int) -> list[dict]:
    users = get_all_users(include_inactive=False)
    return [u for u in users if u["id"] != sender_id and can_message_db(sender_id, u["id"])]


def send_message(sender_id: int, receiver_id: int, subject: str, body: str) -> tuple[bool, str]:
    if not can_message_db(sender_id, receiver_id):
        return False, "غير مسموح لك بمراسلة هذا العضو وفق الصلاحيات الحالية."
    if not body or not body.strip():
        return False, "برجاء كتابة نص الرسالة قبل الإرسال."

    now = datetime.now().isoformat(timespec="seconds")
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO messages (sender_id, receiver_id, subject, body, created_at, is_read)
        VALUES (?, ?, ?, ?, ?, 0)
        """,
        (sender_id, receiver_id, subject.strip(), body.strip(), now),
    )
    conn.commit()
    conn.close()
    return True, "تم إرسال الرسالة بنجاح."


def get_user_messages(user_id: int) -> list[dict]:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            m.*,
            s.full_name AS sender_name,
            s.role AS sender_role,
            r.full_name AS receiver_name,
            r.role AS receiver_role
        FROM messages m
        JOIN users s ON m.sender_id = s.id
        JOIN users r ON m.receiver_id = r.id
        WHERE m.sender_id = ? OR m.receiver_id = ?
        ORDER BY m.id DESC
        """,
        (user_id, user_id),
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def get_audit_logs(limit: int = 50) -> list[dict]:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT a.*, u.full_name AS actor_name
        FROM permission_audit_log a
        JOIN users u ON a.actor_id = u.id
        ORDER BY a.id DESC
        LIMIT ?
        """,
        (limit,),
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def authenticate_user(username: str, password: str):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM users WHERE username = ? AND is_active = 1", (username.strip(),))
    row = cur.fetchone()
    conn.close()
    if not row:
        return None
    if row["password_hash"] != hash_password(password):
        return None
    return dict(row)


# ===== دمج نظام الصلاحيات والمراسلات المتقدم مع بيانات العضويات والرسائل =====
PERMS_ROLE_MAP = {
    "admin": "admin",
    "member": "member",
    "manager": "manager",
    "general_manager": "general_manager",
    "leader1": "leader1",
    "leader2": "leader2",
    "leader3": "leader3",
    "الادمن": "admin",
    "عضو قانونى": "member",
    "مدير إدارة": "manager",
    "مدير عام": "general_manager",
    "قيادي مركز رئيسى 1": "leader1",
    "قيادي مركز رئيسى 2": "leader2",
    "قيادي مركز رئيسى 3": "leader3",
}

def _normalize_perm_role(role_val):
    s = str(role_val or "").strip()
    return PERMS_ROLE_MAP.get(s, PERMS_ROLE_MAP.get(s.lower(), "member"))

def _ensure_perms_db_ready():
    try:
        init_db()
    except Exception:
        pass

def get_member_phone(u):
    if not isinstance(u, dict):
        return ""
    return str(u.get("phone") or u.get("mobile") or u.get("رقم_الموبايل") or "").strip()

def sync_session_user_to_perms_db(u):
    if not isinstance(u, dict):
        return None
    _ensure_perms_db_ready()
    username = str(u.get("username") or "").strip()
    full_name = str(u.get("display_name") or username or "عضو").strip()
    role_code = _normalize_perm_role(u.get("role") or u.get("job_title"))
    phone_val = get_member_phone(u)
    now = datetime.now().isoformat(timespec="seconds")

    conn = get_connection()
    cur = conn.cursor()
    if username:
        cur.execute("SELECT * FROM users WHERE username = ?", (username,))
        row = cur.fetchone()
    else:
        row = None

    if not row:
        base_uname = username or f"user_{u.get('id', 'member')}"
        cur.execute("SELECT * FROM users WHERE username = ?", (base_uname,))
        row = cur.fetchone()
        if not row:
            cur.execute(
                """
                INSERT INTO users (username, full_name, password_hash, role, phone, photo_path, is_active, created_at)
                VALUES (?, ?, ?, ?, ?, ?, 1, ?)
                """,
                (
                    base_uname,
                    full_name,
                    hash_password(str(u.get("password") or "synced_user")),
                    role_code,
                    phone_val,
                    str(u.get("profile_photo_type") or ""),
                    now,
                ),
            )
            perm_user_id = cur.lastrowid
        else:
            perm_user_id = row["id"]
            cur.execute(
                "UPDATE users SET full_name = ?, role = ?, phone = ?, is_active = 1 WHERE id = ?",
                (full_name, role_code, phone_val, perm_user_id),
            )
    else:
        perm_user_id = row["id"]
        cur.execute(
            "UPDATE users SET full_name = ?, role = ?, phone = ?, is_active = 1 WHERE id = ?",
            (full_name, role_code, phone_val, perm_user_id),
        )

    cur.execute("SELECT user_id FROM user_permissions WHERE user_id = ?", (perm_user_id,))
    if not cur.fetchone():
        defaults = DEFAULT_ROLE_PERMISSIONS.get(role_code, DEFAULT_ROLE_PERMISSIONS["member"]).copy()
        cur.execute(
            """
            INSERT INTO user_permissions (
                user_id, can_send_messages, message_scope, allowed_message_roles,
                allowed_message_users, can_manage_permissions, can_view_cases,
                can_edit_cases, can_delete_cases, can_export_reports, can_manage_users, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                perm_user_id,
                int(defaults.get("can_send_messages", True)),
                defaults.get("message_scope", "roles"),
                json.dumps(defaults.get("allowed_message_roles", []), ensure_ascii=False),
                json.dumps([], ensure_ascii=False),
                int(defaults.get("can_manage_permissions", role_code == "admin")),
                int(defaults.get("can_view_cases", True)),
                int(defaults.get("can_edit_cases", True)),
                int(defaults.get("can_delete_cases", False)),
                int(defaults.get("can_export_reports", True)),
                int(defaults.get("can_manage_users", role_code == "admin")),
                now,
            ),
        )
    conn.commit()
    conn.close()
    return perm_user_id

def sync_all_json_users_to_perms_db():
    _ensure_perms_db_ready()
    try:
        json_users = load_users()
    except Exception:
        json_users = []
    for u in json_users:
        if isinstance(u, dict) and u.get("status") != "deleted":
            sync_session_user_to_perms_db(u)

def _resolve_perm_user_id_from_session_or_dict(user_or_role):
    _ensure_perms_db_ready()
    if isinstance(user_or_role, dict):
        return sync_session_user_to_perms_db(user_or_role)
    if isinstance(user_or_role, str):
        role_code = _normalize_perm_role(user_or_role)
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("SELECT id FROM users WHERE role = ? AND is_active = 1 ORDER BY id ASC LIMIT 1", (role_code,))
        row = cur.fetchone()
        conn.close()
        return row["id"] if row else None
    cur_user = st.session_state.get("user")
    if isinstance(cur_user, dict):
        return sync_session_user_to_perms_db(cur_user)
    return None

def can_message(sender_role_or_user, target_role_or_user):
    _ensure_perms_db_ready()
    sender_id = _resolve_perm_user_id_from_session_or_dict(sender_role_or_user)
    target_id = _resolve_perm_user_id_from_session_or_dict(target_role_or_user)
    if sender_id and target_id:
        return bool(can_message_db(sender_id, target_id))
    sr = _normalize_perm_role(sender_role_or_user.get("role") if isinstance(sender_role_or_user, dict) else sender_role_or_user)
    tr = _normalize_perm_role(target_role_or_user.get("role") if isinstance(target_role_or_user, dict) else target_role_or_user)
    if sr == "admin":
        return True
    defaults = DEFAULT_ROLE_PERMISSIONS.get(sr, DEFAULT_ROLE_PERMISSIONS["member"])
    if not defaults.get("can_send_messages", True):
        return False
    if defaults.get("message_scope") == "all_users":
        return True
    return tr in set(defaults.get("allowed_message_roles", []))

def current_user_has_permission(perm_key):
    cur_user = st.session_state.get("user")
    if not isinstance(cur_user, dict):
        return False
    if _normalize_perm_role(cur_user.get("role")) == "admin" and perm_key in ("can_manage_permissions", "can_manage_users"):
        return True
    uid = sync_session_user_to_perms_db(cur_user)
    if not uid:
        return True
    return bool(has_permission(uid, perm_key))


# ===== نظام تعيين وإدارة الأدمن المساعد بواسطة الأدمن 1 =====
def get_primary_admin_user():
    """يعيد سجل الأدمن الأساسي (الأدمن 1) سواء من قاعدة البيانات أو من users.json."""
    try:
        users = load_users()
        admins = [u for u in users if u.get("role") == "admin"]
        if admins:
            admins.sort(key=lambda x: int(x.get("id") or 999999))
            return admins[0]
    except Exception:
        pass
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("SELECT * FROM users WHERE role = 'admin' ORDER BY id ASC LIMIT 1")
        row = cur.fetchone()
        conn.close()
        if row:
            return dict(row)
    except Exception:
        pass
    return None

def is_primary_admin(user_or_id=None):
    """يتحقق ما إذا كان المستخدم هو الأدمن 1 (الأدمن الأساسي)."""
    primary = get_primary_admin_user()
    if not primary:
        return False
    p_id = primary.get("id")
    p_uname = str(primary.get("username") or "").strip().casefold()

    if user_or_id is None:
        u = st.session_state.get("user") or {}
    elif isinstance(user_or_id, dict):
        u = user_or_id
    else:
        u = {"id": user_or_id}

    if not u:
        return False
    if u.get("id") is not None and p_id is not None and str(u.get("id")) == str(p_id):
        return True
    uname = str(u.get("username") or "").strip().casefold()
    if uname and p_uname and uname == p_uname:
        return True
    if u.get("role") == "admin" and not u.get("is_assistant_admin"):
        if p_id is None and not p_uname:
            return True
    return False

def is_assistant_admin(user_or_dict=None):
    """يتحقق ما إذا كان المستخدم معيناً كأدمن مساعد."""
    if user_or_dict is None:
        u = st.session_state.get("user") or {}
    elif isinstance(user_or_dict, dict):
        u = user_or_dict
    else:
        users = _cached_users_snapshot()
        u = next((x for x in users if str(x.get("id")) == str(user_or_dict)), {})
    if not u:
        return False
    if is_primary_admin(u):
        return False
    return bool(u.get("is_assistant_admin")) or str(u.get("role") or "") == "assistant_admin"

def get_assistant_admin_title(user_or_dict=None):
    """يعيد المسمى الذي منحه الأدمن 1 للأدمن المساعد."""
    if user_or_dict is None:
        u = st.session_state.get("user") or {}
    elif isinstance(user_or_dict, dict):
        u = user_or_dict
    else:
        users = _cached_users_snapshot()
        u = next((x for x in users if str(x.get("id")) == str(user_or_dict)), {})
    if not u:
        return "أدمن مساعد"
    custom_title = str(u.get("assistant_admin_title") or u.get("job_title") or "").strip()
    return custom_title or "أدمن مساعد"

def assign_assistant_admin(target_user_id, custom_title, actor_user=None):
    """يعين عضواً كأدمن مساعد بالمسمى الذي يحدده الأدمن 1، ويمنحه كافة الصلاحيات ما عدا المساس بالأدمن 1 أو غلق الموقع."""
    actor = actor_user or st.session_state.get("user") or {}
    if not is_primary_admin(actor):
        return False, "تعيين أو إدارة الأدمن المساعد متاح للأدمن 1 فقط."
    clean_title = str(custom_title or "").strip()
    if not clean_title:
        return False, "يرجى إدخال المسمى المخصص لعضوية الأدمن المساعد."

    users = load_users()
    target = next((u for u in users if str(u.get("id")) == str(target_user_id)), None)
    if not target:
        return False, "لم يتم العثور على العضو المختار."
    if is_primary_admin(target):
        return False, "لا يمكن تعديل عضوية الأدمن 1."

    target["is_assistant_admin"] = True
    target["assistant_admin_title"] = clean_title
    target["job_title"] = clean_title
    target["status"] = "active"
    target["can_close_site"] = False
    target["can_modify_primary_admin"] = False
    register_job_title(clean_title)
    save_users(users)

    uid_db = sync_session_user_to_perms_db(target)
    if uid_db:
        full_assistant_perms = {
            "can_send_messages": True,
            "message_scope": "all_users",
            "allowed_message_roles": ALL_ROLES,
            "allowed_message_users": [],
            "can_manage_permissions": True,
            "can_view_cases": True,
            "can_edit_cases": True,
            "can_delete_cases": True,
            "can_export_reports": True,
            "can_manage_users": True,
        }
        upsert_user_permissions(uid_db, full_assistant_perms)
        actor_db_id = sync_session_user_to_perms_db(actor) or uid_db
        now = datetime.now().isoformat(timespec="seconds")
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(
            "UPDATE users SET full_name = ?, is_active = 1 WHERE id = ?",
            (f"{target.get('display_name') or target.get('username')} ({clean_title})", uid_db),
        )
        cur.execute(
            """
            INSERT INTO permission_audit_log (actor_id, target_scope_desc, details_json, created_at)
            VALUES (?, ?, ?, ?)
            """,
            (
                actor_db_id,
                f"تعيين أدمن مساعد: {target.get('display_name') or target.get('username')} بالمسمى ({clean_title})",
                json.dumps({"target_user_id": target_user_id, "assistant_title": clean_title}, ensure_ascii=False),
                now,
            ),
        )
        conn.commit()
        conn.close()

    cur_session_u = st.session_state.get("user") or {}
    if str(cur_session_u.get("id")) == str(target.get("id")):
        st.session_state.user = dict(target)

    return True, f"تم تعيين ({target.get('display_name') or target.get('username')}) كأدمن مساعد بالمسمى: {clean_title}"

def revoke_assistant_admin(target_user_id, actor_user=None):
    """يلغي صفة الأدمن المساعد عن العضو ويعيده كعضو عادي وتطبق عليه صلاحيات الأدمن 1 كباقي الأعضاء."""
    actor = actor_user or st.session_state.get("user") or {}
    if not is_primary_admin(actor):
        return False, "إلغاء صفة الأدمن المساعد متاح للأدمن 1 فقط."

    users = load_users()
    target = next((u for u in users if str(u.get("id")) == str(target_user_id)), None)
    if not target:
        return False, "لم يتم العثور على العضو."
    if is_primary_admin(target):
        return False, "لا يمكن المساس بعضوية الأدمن 1."

    default_job = "عضو قانونى"
    target["is_assistant_admin"] = False
    target["assistant_admin_title"] = ""
    target["job_title"] = default_job
    target["role"] = "member"
    save_users(users)

    uid_db = sync_session_user_to_perms_db(target)
    if uid_db:
        actor_db_id = sync_session_user_to_perms_db(actor) or uid_db
        reset_user_to_role_defaults(uid_db, actor_id=actor_db_id)

    cur_session_u = st.session_state.get("user") or {}
    if str(cur_session_u.get("id")) == str(target.get("id")):
        st.session_state.user = dict(target)
        st.session_state.role = "member"

    return True, f"تم إلغاء صفة الأدمن المساعد عن ({target.get('display_name') or target.get('username')}) وأصبح عضوًا عاديًا."

def render_assistant_admin_management_section():
    """واجهة مخصصة للأدمن 1 لتعيين وإدارة الأدمن المساعدين بالمسمى الوظيفي المخصص وكافة الصلاحيات."""
    cur_user = st.session_state.get("user") or {}
    if not is_primary_admin(cur_user):
        if is_assistant_admin(cur_user):
            st.info(f"أنت مسجل بصلاحيات الأدمن المساعد بالمسمى: **{get_assistant_admin_title(cur_user)}**. تعيين أو إزالة الأدمن المساعدين أو التعديل على عضوية الأدمن 1 مقتصر على الأدمن 1 فقط.")
        return

    st.markdown("---")
    st.markdown("<h3 style='color:#D4AF37;text-align:center;'>👑 تعيين وإدارة الأدمن المساعدين (للأدمن 1 فقط)</h3>", unsafe_allow_html=True)
    st.caption("يمكنك تعيين عضو أو أكثر كأدمن مساعد ومنحه كافة الصلاحيات مع إظهار عضويته بالمسمى الذي تحدده له، باستثناء حذف أو إيقاف أو تعديل عضوية الأدمن 1 أو غلق الموقع، مع احتفاظك بكافة الصلاحيات عليه.")

    sync_all_json_users_to_perms_db()
    users = load_users()
    primary_admin = get_primary_admin_user()
    primary_id = primary_admin.get("id") if primary_admin else None

    eligible_members = [
        u for u in users
        if u.get("status") != "deleted" and str(u.get("id")) != str(primary_id) and not is_primary_admin(u)
    ]

    if not eligible_members:
        st.info("لا يوجد أعضاء متاحون لتعيينهم كأدمن مساعد حاليًا.")
        return

    with st.form("assign_assistant_admin_form_v1"):
        member_map = {
            f"{u.get('display_name') or u.get('username')} — ({'أدمن مساعد: ' + get_assistant_admin_title(u) if is_assistant_admin(u) else (u.get('job_title') or role_label(u.get('role')))})": u
            for u in eligible_members
        }
        chosen_label = st.selectbox("اختر العضو لتعيينه كأدمن مساعد أو تحديث مسماه", list(member_map.keys()), key="asst_admin_target_sel")
        chosen_u = member_map[chosen_label]
        default_title_val = get_assistant_admin_title(chosen_u) if is_assistant_admin(chosen_u) else ""
        custom_title = st.text_input(
            "المسمى المخصص لعضوية الأدمن المساعد (يظهر به في العضويات والرسائل)",
            value=default_title_val,
            placeholder="مثال: نائب مدير الشؤون القانونية / أدمن مساعد للشؤون القانونية",
            key="asst_admin_title_inp",
        )
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            submit_assign = st.form_submit_button("✅ تعيين / حفظ عضوية الأدمن المساعد", type="primary", use_container_width=True)
        with col_a2:
            submit_revoke = st.form_submit_button("⛔ إلغاء صفة الأدمن المساعد", use_container_width=True)

    if submit_assign:
        ok, msg = assign_assistant_admin(chosen_u.get("id"), custom_title, actor_user=cur_user)
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

    if submit_revoke:
        ok, msg = revoke_assistant_admin(chosen_u.get("id"), actor_user=cur_user)
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

    assistants_list = [u for u in eligible_members if is_assistant_admin(u)]
    st.markdown("#### 📋 قائمة الأدمن المساعدين الحاليين")
    if not assistants_list:
        st.caption("لا يوجد أدمن مساعدون معينون حاليًا.")
    else:
        for au in assistants_list:
            au_id = au.get("id")
            au_name = au.get("display_name") or au.get("username") or ""
            au_title = get_assistant_admin_title(au)
            au_status = au.get("status", "active")
            with st.container(border=True):
                c1, c2, c3, c4 = st.columns([3, 1.5, 1.5, 1.5])
                with c1:
                    st.markdown(f"**{html.escape(str(au_name))}** — المسمى: <span style='color:#00E676;font-weight:900;'>{html.escape(str(au_title))}</span>", unsafe_allow_html=True)
                    st.caption(f"الحالة: {au_status} | المنطقة: {au.get('insurance_area') or 'غير محددة'}")
                with c2:
                    if st.button("🚫 إيقاف", key=f"stop_asst_{au_id}", use_container_width=True):
                        all_u = load_users()
                        for item in all_u:
                            if str(item.get("id")) == str(au_id) and not is_primary_admin(item):
                                item["status"] = "suspended_temp"
                        save_users(all_u)
                        st.success(f"تم إيقاف عضوية الأدمن المساعد: {au_name}")
                        st.rerun()
                with c3:
                    if st.button("✅ تفعيل", key=f"act_asst_{au_id}", use_container_width=True):
                        all_u = load_users()
                        for item in all_u:
                            if str(item.get("id")) == str(au_id) and not is_primary_admin(item):
                                item["status"] = "active"
                        save_users(all_u)
                        st.success(f"تم تفعيل عضوية الأدمن المساعد: {au_name}")
                        st.rerun()
                with c4:
                    if st.button("🗑️ حذف الصفة", key=f"rem_asst_{au_id}", use_container_width=True):
                        ok_r, msg_r = revoke_assistant_admin(au_id, actor_user=cur_user)
                        if ok_r:
                            st.success(msg_r)
                            st.rerun()
                        else:
                            st.error(msg_r)


def render_admin_permissions_page(actor_record):
    st.markdown("<h4 style='color:#D4AF37;'>🛡️ إدارة صلاحيات الأعضاء والمراسلات</h4>", unsafe_allow_html=True)
    users_db = get_all_users(include_inactive=True)
    if not users_db:
        st.info("لا يوجد مستخدمون مسجلون في قاعدة الصلاحيات.")
        return

    user_options = {
        f"{u.get('full_name') or u.get('username')} ({u.get('username')}) — {ROLES.get(u.get('role'), u.get('role'))}": int(u["id"])
        for u in users_db
    }
    selected_label = st.selectbox("اختر العضو لتعديل صلاحياته", list(user_options.keys()), key="perm_target_user_sel")
    target_uid = user_options[selected_label]
    current_perms = get_user_permissions(target_uid)

    with st.form(f"perms_form_{target_uid}"):
        p_send = st.checkbox(PERMISSION_LABELS["can_send_messages"], value=bool(current_perms.get("can_send_messages", True)))
        p_scope = st.selectbox(
            "نطاق المراسلة المسموح",
            list(SCOPE_LABELS.keys()),
            format_func=lambda k: SCOPE_LABELS.get(k, k),
            index=list(SCOPE_LABELS.keys()).index(current_perms.get("message_scope", "roles")) if current_perms.get("message_scope", "roles") in SCOPE_LABELS else 0,
        )
        p_roles = st.multiselect(
            "الأدوار المسموح بمراسلتها",
            ALL_ROLES,
            default=[r for r in current_perms.get("allowed_message_roles", []) if r in ALL_ROLES],
            format_func=lambda r: ROLES.get(r, r),
        )
        other_users = [u for u in users_db if int(u["id"]) != int(target_uid)]
        other_map = {int(u["id"]): f"{u.get('full_name') or u.get('username')} ({u.get('username')})" for u in other_users}
        p_users = st.multiselect(
            "أعضاء محددون بالاسم للمراسلة",
            list(other_map.keys()),
            default=[uid_x for uid_x in current_perms.get("allowed_message_users", []) if uid_x in other_map],
            format_func=lambda uid_x: other_map.get(uid_x, str(uid_x)),
        )
        c1, c2, c3 = st.columns(3)
        with c1:
            p_view = st.checkbox(PERMISSION_LABELS["can_view_cases"], value=bool(current_perms.get("can_view_cases", True)))
            p_edit = st.checkbox(PERMISSION_LABELS["can_edit_cases"], value=bool(current_perms.get("can_edit_cases", True)))
        with c2:
            p_del = st.checkbox(PERMISSION_LABELS["can_delete_cases"], value=bool(current_perms.get("can_delete_cases", False)))
            p_exp = st.checkbox(PERMISSION_LABELS["can_export_reports"], value=bool(current_perms.get("can_export_reports", True)))
        with c3:
            p_mgr_u = st.checkbox(PERMISSION_LABELS["can_manage_users"], value=bool(current_perms.get("can_manage_users", False)))
            p_mgr_p = st.checkbox(PERMISSION_LABELS["can_manage_permissions"], value=bool(current_perms.get("can_manage_permissions", False)))

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            saved = st.form_submit_button("💾 حفظ الصلاحيات", type="primary", use_container_width=True)
        with col_s2:
            reset_def = st.form_submit_button("↩️ استعادة الافتراضي", use_container_width=True)

    if saved:
        new_perms = {
            "can_send_messages": p_send,
            "message_scope": p_scope,
            "allowed_message_roles": p_roles,
            "allowed_message_users": p_users,
            "can_view_cases": p_view,
            "can_edit_cases": p_edit,
            "can_delete_cases": p_del,
            "can_export_reports": p_exp,
            "can_manage_users": p_mgr_u,
            "can_manage_permissions": p_mgr_p,
        }
        actor_id = int(actor_record.get("id") or 1)
        apply_permissions_bulk([target_uid], new_perms, actor_id=actor_id, scope_desc=f"تحديث صلاحيات العضو رقم {target_uid}")
        st.success("✅ تم حفظ الصلاحيات بنجاح.")
        st.rerun()

    if reset_def:
        actor_id = int(actor_record.get("id") or 1)
        reset_user_to_role_defaults(target_uid, actor_id=actor_id)
        st.success("✅ تمت استعادة الصلاحيات الافتراضية.")
        st.rerun()


def render_integrated_admin_permissions_panel():
    cur_user = st.session_state.get("user") or {}
    if _normalize_perm_role(cur_user.get("role")) != "admin" and not current_user_has_permission("can_manage_permissions"):
        st.warning("تعديل صلاحيات الأعضاء متاح لحساب الأدمن أو لمن منحه الأدمن صلاحية إدارة الصلاحيات.")
        return
    sync_all_json_users_to_perms_db()
    actor_perm_id = sync_session_user_to_perms_db(cur_user) or 1
    actor_record = get_user_by_id(actor_perm_id) or {
        "id": actor_perm_id,
        "full_name": cur_user.get("display_name") or cur_user.get("username") or "الادمن",
        "username": cur_user.get("username") or "admin",
        "role": "admin",
        "is_active": 1,
    }
    render_admin_permissions_page(actor_record)
    render_assistant_admin_management_section()


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
    if sr not in ("admin", "assistant_admin", "member", "manager", "general_manager", "leader1", "leader2", "leader3") and not is_assistant_admin(user):
        render_return_home_button("⬅️ العودة للرئيسية", key=f"msg_no_role_home_{sr}")
        return

    users = _cached_users_snapshot()
    messages = load_messages()
    st.markdown("<h2 style='color:#D4AF37;text-align:center;'>📢 الرسائل والبيانات</h2>", unsafe_allow_html=True)
    render_return_home_button("⬅️ العودة للرئيسية", key=f"msg_top_home_{sr}")

    target_options = [
        ("الادمن", "admin"), ("عضو قانونى", "member"), ("مدير إدارة", "manager"),
        ("مدير عام", "general_manager"), ("قيادي مركز رئيسى 1", "leader1"),
        ("قيادي مركز رئيسى 2", "leader2"), ("قيادي مركز رئيسى 3", "leader3")
    ]
    preselected = list(st.session_state.get("message_target_ids") or [])
    candidates = [u for u in users if u.get("status") == "active" and u.get("id") != user.get("id") and can_message(user, u)]
    allowed_roles = [r for _, r in target_options if any(u.get("role") == r for u in candidates)]

    if not candidates:
        st.info("لا يوجد مستلمون متاحون وفق صلاحيات الحساب")
        render_return_home_button("⬅️ العودة للرئيسية", key=f"msg_nocand_home_{sr}")
        return

    kind_choices = ["الجميع"] + [label for label, role in target_options if role in allowed_roles] + ["محددون"]
    default_kind_idx = len(kind_choices) - 1 if preselected else 0
    target_kind = st.selectbox(
        "نوع المستلمين",
        kind_choices,
        index=default_kind_idx,
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
            "sender_job_title": get_assistant_admin_title(user) if is_assistant_admin(user) else (user.get("job_title") or role_label(sr)),
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
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    render_return_home_button("⬅️ العودة للرئيسية", key=f"msg_bottom_home_{sr}")

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


def render_return_home_button(label="⬅️ العودة للرئيسية", key=None, reset_keys=None):
    """زر موحّد للعودة للرئيسية في أي صفحة أو نافذة أو رسالة للأدمن والأعضاء."""
    safe_key = key or f"home_btn_{secrets.token_hex(4)}"
    if st.button(label, key=safe_key, use_container_width=True):
        st.session_state.page = "الرئيسية"
        for k in (reset_keys or []):
            st.session_state.pop(k, None)
        st.rerun()

def get_members_reaching_age_60():
    """يعيد قائمة الأعضاء الذين بلغوا سن 60 سنة فأكثر ولم يتم مد عضويتهم أو إيقافهم لبلوغ السن."""
    users = _cached_users_snapshot()
    due_members = []
    for u in users:
        if not isinstance(u, dict):
            continue
        if u.get("role") == "admin" or u.get("status") in ("deleted", "suspended_age", "suspended_final"):
            continue
        age, extended_active = member_age_status(u)
        if age is not None and age >= 60 and not extended_active:
            due_members.append({
                "user": u,
                "age": age,
                "birth_date_fmt": fmt_date_ar(u.get("birth_date"), "غير مسجل"),
                "name": u.get("display_name") or u.get("username") or "عضو",
                "area": u.get("insurance_area") or "غير محددة",
                "job": u.get("job_title") or role_label(u.get("role")),
            })
    return due_members


def show_admin_age_60_alert():
    """إشعار واضح للأدمن (والأدمن المساعد) عند بلوغ أحد الأعضاء سن 60 سنة."""
    cur_user = st.session_state.get("user") or {}
    if cur_user.get("role") != "admin" and not is_assistant_admin(cur_user):
        return
    due_members = get_members_reaching_age_60()
    if not due_members:
        return

    items_html = ""
    for m in due_members:
        items_html += (
            f"<div style='background:#142038;border:1px solid #ff3b30;border-radius:12px;padding:10px 14px;margin-top:8px;text-align:right;direction:rtl;'>"
            f"<span style='color:#ff5252;font-weight:900;font-size:16px;'>👤 {html.escape(str(m['name']))}</span> "
            f"<span style='color:#ffd700;font-weight:900;'>— العمر: {m['age']} سنة</span> "
            f"<span style='color:#fff;font-weight:700;'>(تاريخ الميلاد: {html.escape(str(m['birth_date_fmt']))} | الوظيفة: {html.escape(str(m['job']))} | المنطقة: {html.escape(str(m['area']))})</span>"
            f"</div>"
        )

    st.markdown(
        f"""
        <div style='background:linear-gradient(135deg,#2c090c,#4a1016);border:2px solid #ff3b30;border-radius:18px;padding:16px;margin:10px 0 14px;box-shadow:0 10px 28px rgba(255,59,48,.28);text-align:center;'>
            <div style='color:#ff5252;font-size:21px;font-weight:900;'>🚨 إشعار للأدمن: بلوغ سن الستين (60 سنة) لأحد الأعضاء</div>
            <div style='color:#fff;font-size:15px;font-weight:800;margin-top:4px;'>يوجد {len(due_members)} عضو بلغ سن الستين فأكثر ويلزم مراجعة موقف عضويته:</div>
            {items_html}
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("⚖️ الانتقال لإدارة الأعضاء لاتخاذ إجراء بلوغ سن الستين", key="admin_age_60_goto_users_btn", use_container_width=True):
        st.session_state.page = "لوحة التحكم"
        st.rerun()


def show_top_notices():
    user = st.session_state.get("user") or {}
    uid = user.get("id")
    if not uid:
        return
    users = _cached_users_snapshot()
    by_id = {u.get("id"): u for u in users}
    notices = []
    incoming_msg_count = 0
    for m in load_messages():
        if uid in m.get("recipient_ids", []):
            incoming_msg_count += 1
            sender = by_id.get(m.get("sender_id"), {})
            notices.append({"kind":"message", "created_at":m.get("created_at", ""), "item":m, "sender":sender})
    for b in _visible_banners_for_user(user):
        sender = by_id.get(b.get("sender_id"), {})
        notices.append({"kind":"banner", "created_at":b.get("created_at", ""), "item":b, "sender":sender})

    if incoming_msg_count > 0:
        st.markdown(
            f"<div style='background:linear-gradient(90deg,#991b1b,#dc2626,#991b1b);color:#ffffff;border:2px solid #ff6b6b;border-radius:14px;padding:10px 14px;margin:0 0 10px;text-align:center;font-weight:900;font-size:17px;box-shadow:0 6px 18px rgba(220,38,38,.35);'>🔴 إشعار رسائل واردة ({incoming_msg_count})</div>",
            unsafe_allow_html=True,
        )

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
    is_incoming_msg = (n["kind"] == "message")
    title = "رسالة واردة" if is_incoming_msg else "بيان وارد"
    created = html.escape(fmt_date_ar(item.get("created_at", ""), str(item.get("created_at", ""))))
    sender_color = "#ff3b30" if is_incoming_msg else ("#00E676" if sender_role == "admin" else "#D4AF37")
    card_border = "#ff3b30" if is_incoming_msg else sender_color
    card_bg = "linear-gradient(135deg,#2b0b10,#46121a)" if is_incoming_msg else "linear-gradient(135deg,#101b31,#203b63)"
    title_color = "#ff4d4d" if is_incoming_msg else "#D4AF37"
    card = f"<div style='background:{card_bg};border:2px solid {card_border};border-radius:18px;padding:16px;margin:0 0 10px;box-shadow:0 10px 28px rgba(0,0,0,.35);text-align:center;'>"
    card += _photo_html(photo, photo_type, 70)
    card += f"<div style='color:{title_color};font-size:22px;font-weight:900;'>{'🔴' if is_incoming_msg else '📢'} {title}</div>"
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
        render_return_home_button("⬅️ العودة للرئيسية", key=f"open_msg_home_{item.get('id','x')}_{idx}")
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
    render_return_home_button("⬅️ العودة للرئيسية", key=f"after_notice_home_{idx}")

def role_order(role):
    return {"leader1":0,"leader2":1,"leader3":2,"general_manager":3,"manager":4,"member":5,"admin":-1}.get(role,99)

def show_member_directory():
    sync_all_json_users_to_perms_db()
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
                raw_name = str(u.get("display_name") or u.get("username") or "").strip()
                name = html.escape(raw_name)
                is_asst = is_assistant_admin(u)
                is_admin = is_primary_admin(u) or (u.get("role") == "admin" and not is_asst)
                custom_job = membership_job_title(u)
                if is_asst:
                    job_value = get_assistant_admin_title(u)
                elif is_admin:
                    job_value = custom_job or "الادمن"
                else:
                    job_value = custom_job or role_label(u.get("role"))
                job = html.escape(job_value)
                photo = _photo_html(u.get("profile_photo"), u.get("profile_photo_type") or "image/png", 44)
                border = "#00E676" if (is_admin or is_asst) else "#D4AF37"
                name_color = "#00E676" if (is_admin or is_asst) else "#fff"
                if is_admin:
                    admin_title_line = f"الادمن — {name}" if raw_name and raw_name.casefold() != "ادمن" and raw_name.casefold() != "الادمن" else "الادمن"
                    admin_badge = f"<div style='color:#00E676;font-size:10px;font-weight:900;line-height:1.2;'>{html.escape(admin_title_line)}</div>"
                elif is_asst:
                    admin_badge = "<div style='color:#00E676;font-size:10px;font-weight:900;'>" + html.escape(job_value) + "</div>"
                else:
                    admin_badge = ""
                phone_txt = html.escape(get_member_phone(u))
                phone_badge = f"<div style='color:#42d77d;font-size:10px;font-weight:800;line-height:1.2;margin-top:2px;'>📱 {phone_txt}</div>" if phone_txt else "<div style='color:#94a3b8;font-size:9px;line-height:1.2;margin-top:2px;'>📱 غير مسجل</div>"
                card = f"<div style='background:#142038;border:2px solid {border};border-radius:12px;padding:6px;text-align:center;min-height:134px;margin-bottom:5px;'>{photo}<div style='color:{name_color};font-size:12px;font-weight:900;line-height:1.2;'>{name}</div>{admin_badge}<div style='color:#D4AF37;font-size:10px;font-weight:800;line-height:1.2;'>{job}</div><div style='color:#cbd5e1;font-size:9px;line-height:1.2;'>{html.escape(area)}</div>{phone_badge}</div>"
                st.markdown(card, unsafe_allow_html=True)
                if u.get("id") != current.get("id") and can_message(current, u):
                    if st.button("✉️ إرسال", key=f"dir_msg_v5_{current.get('id')}_{u.get('id')}", use_container_width=True):
                        st.session_state.message_target_ids = [u.get("id")]
                        st.session_state.page = "الرسائل"
                        st.rerun()
                elif u.get("id") == current.get("id") and current.get("role") != "admin":
                    if st.button("⚙️ لوحة تحكمى", key=f"dir_my_ctrl_{current.get('id')}_{u.get('id')}", use_container_width=True):
                        st.session_state.page = "لوحة تحكمى"
                        st.rerun()

def global_people_search():
    users=_cached_users_snapshot()
    q=st.text_input("🔎 بحث سريع بالاسم", placeholder="اكتب اسم العضو أو مدير الإدارة أو المدير العام", key="global_people_search")
    if q.strip():
        matches=[u for u in users if q.strip().lower() in str(u.get("display_name") or u.get("username","")).lower()]
        if matches:
            for u in matches:
                _ph = get_member_phone(u) or "غير مسجل"
                if u.get("profile_photo"):
                    st.markdown(_photo_html(u.get("profile_photo"), u.get("profile_photo_type") or "image/png", 48), unsafe_allow_html=True)
                _disp_name = str(u.get("display_name") or u.get("username") or "").strip()
                _job_disp = membership_job_title(u) or role_label(u.get("role"))
                if u.get("role") == "admin" and not is_assistant_admin(u):
                    _role_prefix = f"الادمن ({_disp_name})" if _disp_name else "الادمن"
                    _job_str = f"{_role_prefix} — الوظيفة: {_job_disp}" if _job_disp and _job_disp != "الادمن" else _role_prefix
                else:
                    _job_str = f"{_disp_name} — {_job_disp}"
                st.markdown(f"**👤 {_job_str}** — المنطقة التأمينية: {u.get('insurance_area') or 'لا توجد منطقة'} — 📱 الموبايل: {_ph} — الحالة: {u.get('status','')}")
        else:
            st.info("لا توجد نتيجة بهذا الاسم")

def admin_control_page():
    cur_u = st.session_state.get("user") or {}
    if cur_u.get("role") != "admin" and not is_assistant_admin(cur_u):
        st.error("هذه الصفحة متاحة للادمن فقط")
        return
    if is_assistant_admin(cur_u):
        st.info(f"🛡️ صلاحيات الأدمن المساعد — المسمى الوظيفي: **{html.escape(get_assistant_admin_title(cur_u))}** (تملك كافة الصلاحيات باستثناء المساس بعضوية الأدمن 1 أو غلق الموقع).")

    st.markdown("<h1 style='color:#D4AF37;text-align:center;'>⚖️ لوحة التحكم</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:white;text-align:center;'>جميع أدوات إدارة الموقع في مكان واحد</p>", unsafe_allow_html=True)
    render_return_home_button("⬅️ العودة للرئيسية", key="admin_control_top_home")

    with st.expander("☁️ حالة الربط السحابي والمزامنة الفورية"):
        if cloud_status():
            st.success("☁️ متصل بإعدادات السحابة (Supabase) — الحفظ يتم محليًا ويُزامَن سحابيًا.")
        else:
            st.warning(f"💾 وضع العمل المحلي/الاحتياطي نشط ({_CLOUD_LAST_ERROR or 'أضف SUPABASE_URL و SUPABASE_KEY في Secrets لتفعيل الربط السحابي الكامل'}).")
        sc1, sc2 = st.columns(2)
        with sc1:
            if st.button("☁️ رفع ومزامنة البيانات المحلية للسحابة", key="cloud_sync_push_btn", use_container_width=True):
                ok_s, msg_s = sync_all_local_to_cloud()
                if ok_s:
                    st.success(msg_s)
                else:
                    st.warning(msg_s)
        with sc2:
            if st.button("⬇️ سحب وتحديث البيانات من السحابة", key="cloud_sync_pull_btn", use_container_width=True):
                ok_p, msg_p = pull_all_cloud_to_local()
                if ok_p:
                    st.success(msg_p)
                    st.rerun()
                else:
                    st.warning(msg_p)

    tab_users, tab_perms, tab_messages, tab_topbar, tab_search, tab_account = st.tabs([
        "👥 الأعضاء",
        "🛡️ صلاحيات الأعضاء",
        "📢 الرسائل والبيانات",
        "🪧 الشريط العلوي",
        "🔎 البحث والملف الشخصي",
        "🔐 حساب الادمن والإيميل",
    ])

    with tab_perms:
        render_integrated_admin_permissions_panel()

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
            col1, col2 = st.columns(2)
            with col1:
                duration_value = st.number_input("مدة الإبقاء", min_value=1, max_value=3650, value=7, step=1)
            with col2:
                duration_unit = st.selectbox("وحدة المدة", ["أيام", "ساعات", "دقائق"])
            audience = st.radio("يظهر لـ", ["كل الأعضاء", "أعضاء محددين"], horizontal=True, key="banner_audience_v2")
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
                        "attachments": [
                            {
                                "data": base64.b64encode(f.getvalue()).decode("utf-8"),
                                "name": f.name,
                                "type": f.type or "application/octet-stream",
                            }
                            for f in (banner_files or [])
                        ],
                    }
                    banners.append(banner)
                    save_banners(banners)
                    st.success("تم نشر الرسالة / البيان")
                    st.rerun()

        st.markdown("---")
        st.markdown(f"<h3 style='color:#D4AF37;'>الرسائل والبيانات الظاهرة أعلى الصفحة: {len(banners)}</h3>", unsafe_allow_html=True)

        for i, b in enumerate(banners):
            expire_raw = b.get("expire_at") or b.get("expire")
            expire_dt = parse_date_safe(expire_raw)
            try:
                expire_full_dt = datetime.fromisoformat(str(expire_raw))
            except Exception:
                expire_full_dt = None
            expired = bool(expire_full_dt and expire_full_dt <= datetime.now())
            audience_text = "الكل" if b.get("audience") in ("الكل", "كل الأعضاء") else "، ".join(b.get("visible_to", []))
            expire_display = fmt_date_ar(expire_raw, "غير محدد")
            with st.container(border=True):
                st.markdown(
                    f"<div class='banner-admin-card'><div style='color:#D4AF37;font-size:20px;font-weight:900;'>📢 {b.get('text','')}</div>"
                    f"<div style='color:white;margin-top:8px;'>الجمهور: {audience_text}</div>"
                    f"<div style='color:#ddd;'>تنتهي: {expire_display} | الحالة: {'منتهية' if expired else 'فعالة'}</div>"
                    + "".join(f"<div style='color:#fff;margin-top:5px;'>📎 {html.escape(str(a.get('name') or 'مرفق'))}</div>" for a in (b.get('attachments') or []))
                    + "</div>",
                    unsafe_allow_html=True,
                )
                e1, e2 = st.columns(2)
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
                        nt = st.text_area("نص الرسالة أو البيان", value=b.get("text", ""), key=f"et_{b.get('id',i)}")
                        nc = st.color_picker("اللون", value=b.get("color", "#D4AF37"), key=f"ec_{b.get('id',i)}")
                        nfiles = st.file_uploader("مرفقات البيان", type=None, accept_multiple_files=True, key=f"eimg_{b.get('id',i)}")
                        na = st.radio(
                            "المستلمون",
                            ["كل الأعضاء", "أعضاء محددين"],
                            index=0 if b.get("audience") in ("الكل", "كل الأعضاء") else 1,
                            horizontal=True,
                            key=f"ea_{b.get('id',i)}",
                        )
                        nvis = []
                        if na == "أعضاء محددين":
                            edit_candidates = [u for u in users if u.get("status") == "active"]
                            default_ids = [u.get("id") for u in edit_candidates if u.get("username") in b.get("visible_to", [])]
                            selected_edit_ids = _recipient_groups(edit_candidates, f"banner_edit_recipient_{b.get('id',i)}", default_ids)
                            nvis = [u.get("username") for u in edit_candidates if u.get("id") in selected_edit_ids]
                        new_expire = date_input_dmy(
                            "تاريخ الانتهاء",
                            value=expire_dt or datetime.now().date(),
                            max_value=date(2100, 12, 31),
                            key=f"ed_{b.get('id',i)}",
                        )
                        ec1, ec2 = st.columns(2)
                        with ec1:
                            if st.form_submit_button("💾 حفظ التعديل", use_container_width=True):
                                b["text"] = nt.strip()
                                b["color"] = nc
                                b["audience"] = "الكل" if na == "كل الأعضاء" else "اعضاء محددين"
                                b["visible_to"] = nvis
                                if nfiles:
                                    b["attachments"] = b.get("attachments") or []
                                    b["attachments"].extend([
                                        {
                                            "data": base64.b64encode(f.getvalue()).decode("utf-8"),
                                            "name": f.name,
                                            "type": f.type or "application/octet-stream",
                                        }
                                        for f in nfiles
                                    ])
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

    with tab_topbar:
        st.markdown("<h3 style='color:#D4AF37;text-align:center;'>🪧 التحكم الكامل في الشريط الظاهر أعلى البرنامج</h3>", unsafe_allow_html=True)
        st.caption("يمكنك من هنا إظهار الشريط أو إيقافه/إخفائه بالكامل، أو تعديل نصه وألوانه وحجمه وسرعته أو إلغاء نصه دون المساس بالكود.")
        cur_bar = load_top_bar_config()

        st.markdown("<div style='color:#D4AF37;font-weight:900;margin-bottom:6px;'>معاينة حية للشريط الحالي:</div>", unsafe_allow_html=True)
        if cur_bar.get("enabled", True) and str(cur_bar.get("text") or "").strip():
            render_top_bar()
        else:
            st.info("الشريط العلوي متوقف / مخفي حاليًا.")

        with st.form("admin_top_bar_control_form"):
            bar_enabled = st.checkbox(
                "تفعيل / إظهار الشريط العلوي أعلى البرنامج",
                value=bool(cur_bar.get("enabled", True)),
                key="bar_enabled_chk"
            )
            bar_text = st.text_area(
                "نص الشريط العلوي (اتركه فارغًا أو اضغط زر الإلغاء لإلغاء النص)",
                value=str(cur_bar.get("text") or ""),
                height=90,
                key="bar_text_inp"
            )
            bar_mode = st.selectbox(
                "نمط عرض الشريط",
                ["marquee", "static"],
                format_func=lambda m: "متحرك (شريط إخباري متحرك)" if m == "marquee" else "ثابت (نص ثابت واضح)",
                index=0 if cur_bar.get("mode", "marquee") == "marquee" else 1,
                key="bar_mode_sel"
            )
            c1, c2, c3 = st.columns(3)
            with c1:
                bar_bg_start = st.color_picker("لون بداية الخلفية", value=str(cur_bar.get("bg_start") or "#D4AF37"), key="bar_bg1")
            with c2:
                bar_bg_mid = st.color_picker("لون منتصف الخلفية", value=str(cur_bar.get("bg_mid") or "#FFD700"), key="bar_bg2")
            with c3:
                bar_bg_end = st.color_picker("لون نهاية الخلفية", value=str(cur_bar.get("bg_end") or "#D4AF37"), key="bar_bg3")

            c4, c5, c6 = st.columns(3)
            with c4:
                bar_text_color = st.color_picker("لون النص", value=str(cur_bar.get("text_color") or "#0A1428"), key="bar_fg")
            with c5:
                bar_font_size = st.slider("حجم الخط (px)", min_value=12, max_value=32, value=int(cur_bar.get("font_size", 16)), key="bar_fs")
            with c6:
                bar_speed = st.slider("مدة دورة الحركة (ثواني)", min_value=5, max_value=60, value=int(cur_bar.get("speed_seconds", 15)), key="bar_spd")

            if st.form_submit_button("💾 حفظ إعدادات الشريط العلوي", use_container_width=True, type="primary"):
                updated_cfg = {
                    "enabled": bool(bar_enabled),
                    "text": bar_text.strip(),
                    "mode": bar_mode,
                    "bg_start": bar_bg_start,
                    "bg_mid": bar_bg_mid,
                    "bg_end": bar_bg_end,
                    "text_color": bar_text_color,
                    "font_size": int(bar_font_size),
                    "speed_seconds": int(bar_speed),
                }
                if save_top_bar_config(updated_cfg):
                    st.success("✅ تم حفظ إعدادات الشريط العلوي وتطبيقها فورًا.")
                    st.rerun()

        b_col1, b_col2, b_col3 = st.columns(3)
        with b_col1:
            if st.button("⏸️ إيقاف / إخفاء الشريط", use_container_width=True, key="bar_stop_btn"):
                cur_bar["enabled"] = False
                if save_top_bar_config(cur_bar):
                    st.success("تم إيقاف الشريط العلوي.")
                    st.rerun()
        with b_col2:
            if st.button("▶️ تشغيل / إظهار الشريط", use_container_width=True, key="bar_enable_btn"):
                cur_bar["enabled"] = True
                if not str(cur_bar.get("text") or "").strip():
                    cur_bar["text"] = DEFAULT_TOP_BAR_CONFIG["text"]
                if save_top_bar_config(cur_bar):
                    st.success("تم تشغيل الشريط العلوي.")
                    st.rerun()
        with b_col3:
            if st.button("🗑️ إلغاء نص الشريط وإيقافه", use_container_width=True, key="bar_clear_btn"):
                cur_bar["text"] = ""
                cur_bar["enabled"] = False
                if save_top_bar_config(cur_bar):
                    st.success("تم إلغاء نص الشريط وإيقافه.")
                    st.rerun()

        render_admin_audio_manager(actor_username=str(cur_u.get("username") or "admin"))

    with tab_search:
        st.markdown("<h3 style='color:#D4AF37;'>🔎 البحث عن عضو أو مدير إدارة أو مدير عامأو قيادي</h3>", unsafe_allow_html=True)
        global_people_search()
        st.markdown("---")
        st.markdown("<h3 style='color:#D4AF37;'>👤 الملف الشخصي</h3>", unsafe_allow_html=True)
        admin_user = next((u for u in load_users() if u.get("role") == "admin"), st.session_state.user)
        _adm_name = str(admin_user.get("display_name") or admin_user.get("username") or "").strip()
        _adm_job = membership_job_title(admin_user) or role_label(admin_user.get("role"))
        _adm_phone = get_member_phone(admin_user) or "غير مسجل"
        if admin_user.get("profile_photo"):
            st.markdown(_photo_html(admin_user.get("profile_photo"), admin_user.get("profile_photo_type") or "image/png", 80), unsafe_allow_html=True)
        st.write(f"**الصفة والاسم:** الادمن — {_adm_name}" if _adm_name else "**الصفة:** الادمن")
        st.write(f"**الوظيفة:** {_adm_job}")
        st.write(f"**المنطقة التأمينية:** {admin_user.get('insurance_area') or 'غير مسجلة'}")
        st.write(f"**رقم الموبايل:** {_adm_phone}")
        st.write(f"**البريد الإلكتروني:** {admin_user.get('email') or 'غير مسجل'}")
        st.write(f"**الحالة:** {admin_user.get('status','')}")

    with tab_account:
        st.markdown("<h3 style='color:#D4AF37;'>👤 بيانات حساب الادمن</h3>", unsafe_allow_html=True)
        users = _cached_users_snapshot()
        admin_user = next((u for u in users if u.get("role") == "admin"), None)
        if admin_user:
            with st.form("admin_profile_edit_form_v4"):
                display_name = st.text_input("الاسم (يظهر بجانب الادمن للجميع)", value=admin_user.get("display_name") or admin_user.get("username") or "admin")
                job_title = st.text_input("الوظيفة", value=admin_user.get("job_title") or "")
                insurance_area = st.text_input("المنطقة التأمينية", value=admin_user.get("insurance_area") or "")
                admin_phone = st.text_input("رقم الموبايل", value=get_member_phone(admin_user), placeholder="01xxxxxxxxx", key="admin_profile_phone_inp")
                login_name = st.text_input("اسم الدخول", value=admin_user.get("username") or "admin")
                admin_photo = st.file_uploader("الصورة الشخصية", type=["jpg", "jpeg", "png"], key="admin_profile_photo_upload")
                if st.form_submit_button("💾 حفظ بيانات الادمن", use_container_width=True, type="primary"):
                    if not is_primary_admin(st.session_state.get("user")):
                        st.error("غير مصرح بتعديل أو إيقاف أو حذف عضوية الأدمن 1 إلا لحساب الأدمن 1 فقط")
                        st.stop()
                    login_name = login_name.strip() or "admin"
                    if login_name != admin_user.get("username") and any(u.get("username") == login_name and u.get("id") != admin_user.get("id") for u in users):
                        st.error("اسم الدخول مستخدم بالفعل")
                    else:
                        clean_adm_phone = admin_phone.strip()
                        admin_user["display_name"] = display_name.strip() or login_name
                        admin_user["username"] = login_name
                        if job_title.strip():
                            register_job_title(job_title.strip())
                        admin_user["job_title"] = job_title.strip()
                        admin_user["insurance_area"] = insurance_area.strip()
                        admin_user["phone"] = clean_adm_phone
                        admin_user["mobile"] = clean_adm_phone
                        if admin_photo:
                            admin_user["profile_photo"] = base64.b64encode(admin_photo.getvalue()).decode("utf-8")
                            admin_user["profile_photo_type"] = admin_photo.type
                        save_users(users)
                        st.session_state.user = admin_user
                        st.success("تم حفظ اسم الادمن والوظيفة والمنطقة التأمينية ورقم الموبايل والصورة وبيانات الدخول")
                        st.rerun()
        st.info("يظهر اسم الادمن بجوار كلمة الادمن، وكذلك وظيفته ورقم موبايله وصورته للجميع في العضويات والبحث.")
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
    st.markdown("<p style='text-align:center;color:#fff!important;font-weight:900;'>أدخل البريد الإلكتروني لاسترداد الحساب وكلمة السر التي تريد استخدامها لحسابك</p>", unsafe_allow_html=True)
    recovery_email_first = st.text_input("البريد الإلكتروني لاسترداد الحساب (اختياري)", placeholder="example@domain.com", key="first_recovery_email")
    new_pass = st.text_input("كلمة السر الجديدة", type="password", key="first_new_pass")
    confirm_pass = st.text_input("تأكيد كلمة السر", type="password", key="first_confirm_pass")
    profile_photo = st.file_uploader("الصورة الشخصية", type=["jpg","jpeg","png"], key="first_profile_photo")
    if st.button("حفظ كلمة السر والدخول", type="primary", use_container_width=True, key="save_first_password"):
        cleaned_email = recovery_email_first.strip()
        if not new_pass:
            st.error("اكتب كلمة السر أولاً")
            return
        if new_pass != confirm_pass:
            st.error("كلمتا السر غير متطابقتين")
            return
        if cleaned_email and ("@" not in cleaned_email or "." not in cleaned_email.split("@")[-1]):
            st.error("أدخل بريدًا إلكترونيًا صحيحًا لاسترداد الحساب")
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
        if cleaned_email:
            target["email"] = cleaned_email
            target["recovery_email"] = cleaned_email
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
    render_return_home_button("⬅️ العودة للرئيسية", key="personal_control_top_home")

    tab_profile, tab_permissions, tab_messages, tab_account = st.tabs(
        ["👤 الملف الشخصي", "🛡️ صلاحياتى", "📢 الرسائل والبيانات", "🔐 الحساب"]
    )

    with tab_profile:
        st.markdown("<h3 style='color:#D4AF37;'>بيانات العضوية</h3>", unsafe_allow_html=True)
        if user.get("profile_photo"):
            st.markdown(_photo_html(user.get("profile_photo"), user.get("profile_photo_type") or "image/png", 88), unsafe_allow_html=True)
        st.write(f"**الاسم:** {user.get('display_name') or user.get('username','')}")
        st.write(f"**الوظيفة:** {user.get('job_title') or role_label(role)}")
        st.write(f"**تاريخ الميلاد:** {fmt_date_ar(user.get('birth_date'), 'غير مسجل')}")
        st.write(f"**المنطقة التأمينية:** {user.get('insurance_area') or 'غير مسجلة'}")
        st.write(f"**رقم الموبايل:** {get_member_phone(user) or 'غير مسجل'}")
        st.write(f"**البريد الإلكتروني:** {user.get('email') or 'غير مسجل'}")
        st.write(f"**الحالة:** {user.get('status','')}")
        st.info("إيقاف حساب الدخول لا يحذف القضايا أو المستندات أو الرسائل المحفوظة.")
        st.markdown("---")
        st.markdown("<h4 style='color:#D4AF37;'>📱 إضافة أو تعديل رقم الموبايل والصورة الشخصية (تظهر للجميع)</h4>", unsafe_allow_html=True)
        with st.form("member_profile_phone_photo_form"):
            inp_phone = st.text_input("رقم الموبايل", value=get_member_phone(user), placeholder="01xxxxxxxxx")
            inp_photo = st.file_uploader("رفع أو تحديث الصورة الشخصية", type=["jpg", "jpeg", "png"], key="member_profile_photo_uploader")
            if st.form_submit_button("💾 حفظ رقم الموبايل والصورة", use_container_width=True, type="primary"):
                if save_member_phone_and_photo(user.get("id"), inp_phone, inp_photo):
                    st.success("تم حفظ رقم الموبايل والصورة الشخصية بنجاح وظهورهما للجميع.")
                    st.rerun()

    with tab_permissions:
        st.markdown("<h3 style='color:#D4AF37;'>الصلاحيات المقررة لحسابك</h3>", unsafe_allow_html=True)
        targets = [
            ("الادمن", "admin"), ("عضو قانونى", "member"), ("مدير إدارة", "manager"),
            ("مدير عام", "general_manager"), ("قيادي مركز رئيسى 1", "leader1"),
            ("قيادي مركز رئيسى 2", "leader2"), ("قيادي مركز رئيسى 3", "leader3"),
        ]
        allowed = [label for label, target_role in targets if can_message(role, target_role)]
        if allowed:
            st.success("يمكنك إرسال رسالة أو بيان إلى: " + "، ".join(allowed))
        else:
            st.info("لا توجد صلاحية إرسال رسائل لهذا الحساب.")

    with tab_messages:
        messaging_page()

    with tab_account:
        st.markdown("<h3 style='color:#D4AF37;'>🔐 إعدادات حسابك الشخصي</h3>", unsafe_allow_html=True)
        st.info("تعديل بيانات حساب الادمن متاح لحساب الادمن فقط. من هنا يمكنك إدارة رقم الموبايل والصورة الشخصية والبريد الإلكتروني وكلمة السر الخاصة بحسابك.")
        with st.form("member_account_phone_photo_form"):
            acc_phone = st.text_input("رقم الموبايل", value=get_member_phone(user), placeholder="01xxxxxxxxx")
            acc_photo = st.file_uploader("الصورة الشخصية", type=["jpg", "jpeg", "png"], key="member_account_photo_uploader")
            if st.form_submit_button("💾 حفظ رقم الموبايل والصورة", use_container_width=True, type="primary"):
                if save_member_phone_and_photo(user.get("id"), acc_phone, acc_photo):
                    st.success("تم حفظ رقم الموبايل والصورة الشخصية بنجاح.")
                    st.rerun()
        st.markdown("---")
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
# === نهاية الجزء 1 من 3 ===
# === بداية الجزء 2 من 3 ===
def edit_memberships_page():
    cur_u = st.session_state.get("user") or {}
    if cur_u.get("role") != "admin" and not is_assistant_admin(cur_u):
        st.error("هذه الصفحة متاحة للادمن فقط")
        return

    st.markdown("<h2 style='color:#D4AF37;text-align:center;'>✏️ تعديل بيانات العضويات</h2>", unsafe_allow_html=True)
    render_return_home_button("⬅️ العودة للرئيسية", key="edit_memberships_home")
    if st.button("⬅️ العودة لإدارة الأعضاء", key="back_to_manage_users", use_container_width=True):
        st.session_state.page = "لوحة التحكم"
        st.rerun()

    users = load_users()
    members = [u for u in users if u.get("role") != "admin" and u.get("status") != "deleted"]

    if not members:
        st.info("لا توجد عضويات لتعديلها حالياً.")
        return

    for idx, user in enumerate(members):
        uid = user.get("id")
        current_name = user.get("display_name") or user.get("username") or ""
        current_login = user.get("username") or ""
        current_area = user.get("insurance_area") or ""
        current_birth = parse_date_safe(user.get("birth_date")) or date(1990, 1, 1)
        current_job = membership_job_title(user) or role_label(user.get("role")) or "عضو قانونى"
        current_phone = get_member_phone(user)

        with st.container(border=True):
            if user.get("profile_photo"):
                st.markdown(_photo_html(user.get("profile_photo"), user.get("profile_photo_type") or "image/png", 76), unsafe_allow_html=True)
            st.markdown(f"### 👤 {current_name}")
            with st.form(f"edit_membership_form_{uid}_{idx}"):
                new_name = st.text_input("الاسم الظاهر للعضو", value=current_name, key=f"em_name_{uid}_{idx}")
                new_login = st.text_input("اسم الدخول", value=current_login, key=f"em_login_{uid}_{idx}")
                new_area = st.text_input("المنطقة التأمينية", value=current_area, key=f"em_area_{uid}_{idx}")
                new_phone = st.text_input("رقم الموبايل", value=current_phone, placeholder="01xxxxxxxxx", key=f"em_phone_{uid}_{idx}")
                new_birth = date_input_dmy(
                    "تاريخ الميلاد",
                    value=current_birth,
                    min_value=date(1900, 1, 1),
                    max_value=date(2100, 12, 31),
                    key=f"em_birth_{uid}_{idx}",
                )
                job_options = get_all_job_titles(include_users=True)
                if current_job and current_job not in job_options:
                    job_options.append(current_job)
                new_job = st.selectbox(
                    "الوظيفة",
                    job_options,
                    index=job_options.index(current_job) if current_job in job_options else 0,
                    key=f"em_job_{uid}_{idx}",
                )
                new_photo = st.file_uploader("الصورة الشخصية", type=["jpg", "jpeg", "png"], key=f"em_photo_{uid}_{idx}")

                if st.form_submit_button("💾 حفظ التعديلات", type="primary", use_container_width=True):
                    cleaned_name = new_name.strip()
                    cleaned_login = new_login.strip() or cleaned_name
                    cleaned_area = new_area.strip()
                    cleaned_phone = new_phone.strip()
                    if not cleaned_name:
                        st.error("اسم العضو مطلوب")
                    elif not cleaned_area:
                        st.error("المنطقة التأمينية مطلوبة")
                    elif any(
                        u.get("id") != uid and (
                            (u.get("username") or "").strip().casefold() == cleaned_login.casefold()
                            or (u.get("display_name") or "").strip().casefold() == cleaned_name.casefold()
                        )
                        for u in users
                    ):
                        st.error("يوجد عضو آخر بنفس اسم الدخول أو الاسم")
                    else:
                        role_val = {
                            "عضو قانونى": "member",
                            "مدير إدارة": "manager",
                            "مدير عام": "general_manager",
                            "مدير عام (قيادى مناطق)": "general_manager",
                            "قيادي مركز رئيسى 1": "leader1",
                            "قيادي مركز رئيسى 2": "leader2",
                            "قيادي مركز رئيسى 3": "leader3",
                        }.get(new_job, user.get("role") or "member")

                        user["display_name"] = cleaned_name
                        user["username"] = cleaned_login
                        user["insurance_area"] = cleaned_area
                        user["phone"] = cleaned_phone
                        user["mobile"] = cleaned_phone
                        user["birth_date"] = fmt_date_ar(new_birth)
                        register_job_title(new_job)
                        user["job_title"] = new_job
                        user["role"] = role_val
                        if new_photo:
                            user["profile_photo"] = base64.b64encode(new_photo.getvalue()).decode("utf-8")
                            user["profile_photo_type"] = new_photo.type

                        save_users(users)
                        st.success(f"تم حفظ تعديلات العضوية بنجاح: {cleaned_name}")
                        st.rerun()

# ===== القضايا والجلسات والمكتبة القانونية =====
DATA_FILE = os.path.join(LOCAL_DATA_DIR, "cases_data.json")
TOKENS_FILE = os.path.join(LOCAL_DATA_DIR, "tokens.json")
CASE_TYPES = ["استئناف", "نقض", "مذكرات دفاع", "محاكم إدارية", "قضاء إداري", "إدارية عليا", "فتاوى", "تحقيقات"]
CASE_STATUSES = ["متداولة", "حجز للحكم", "مؤجلة", "منظورة", "منتهية بحكم"]

def load_data():
    default = {"cases": [], "library": []}
    data = load_json_persistent("data", DATA_FILE, default)
    if not isinstance(data, dict):
        data = dict(default)
    data.setdefault("cases", [])
    data.setdefault("library", [])
    return data

def save_data(data):
    return save_json_persistent("data", DATA_FILE, data)

def load_tokens():
    tokens = load_json_persistent("tokens", TOKENS_FILE, {})
    return tokens if isinstance(tokens, dict) else {}

def save_tokens(tokens):
    return save_json_persistent("tokens", TOKENS_FILE, tokens)

def ensure_case_defaults(case):
    case.setdefault("id", secrets.token_hex(6))
    case.setdefault("رقم_القضية", "")
    case.setdefault("النوع", CASE_TYPES[0])
    case.setdefault("الخصوم", "")
    case.setdefault("الدائرة", "")
    case.setdefault("المحكمة", "")
    case.setdefault("تاريخ_جلسة", "")
    case.setdefault("الحالة", CASE_STATUSES[0])
    case.setdefault("الحكم_أو_القرار", "")
    case.setdefault("ملخص_الموضوع", "")
    case.setdefault("ملاحظات", "")
    case.setdefault("جلسات", [])
    case.setdefault("مستندات", [])
    case.setdefault("user_id", st.session_state.get("user", {}).get("id"))
    case.setdefault("assigned_to", st.session_state.get("user", {}).get("username", ""))
    return case

def current_user_can_edit_case(case):
    user = st.session_state.get("user") or {}
    if user.get("role") == "admin" or is_assistant_admin(user):
        return True
    if not current_user_has_permission("can_edit_cases"):
        return False
    return case.get("user_id") == user.get("id") or case.get("assigned_to") == user.get("username")

def get_visible_cases():
    data = load_data()
    cases = [ensure_case_defaults(dict(c)) for c in data.get("cases", [])]
    user = st.session_state.get("user") or {}
    role = user.get("role")
    if role == "admin" or is_assistant_admin(user):
        return cases
    if not current_user_has_permission("can_view_cases"):
        return []
    return [c for c in cases if c.get("user_id") == user.get("id") or c.get("assigned_to") == user.get("username")]

def render_cases_page():
    user = st.session_state.get("user") or {}
    if user.get("role") != "admin" and not is_assistant_admin(user) and not current_user_has_permission("can_view_cases"):
        st.error("لا تملك صلاحية عرض القضايا والملفات وفق الصلاحيات المقررة لحسابك.")
        render_return_home_button("⬅️ العودة للرئيسية", key="cases_no_perm_home")
        return

    st.markdown("<h2 style='color:#D4AF37;text-align:center;'>📁 إدارة القضايا</h2>", unsafe_allow_html=True)
    render_return_home_button("⬅️ العودة للرئيسية", key="cases_top_home")

    users = _cached_users_snapshot()
    member_names = [u.get("username") for u in users if u.get("username") and u.get("status") != "deleted"]
    default_owner = user.get("username") if user.get("username") in member_names else (member_names[0] if member_names else "")

    with st.expander("➕ إضافة قضية جديدة", expanded=not get_visible_cases()):
        if user.get("role") != "admin" and not is_assistant_admin(user) and not current_user_has_permission("can_edit_cases"):
            st.warning("حسابك يملك صلاحية العرض فقط ولا يملك صلاحية إضافة القضايا.")
        else:
            with st.form("add_case_form"):
                c1, c2, c3 = st.columns(3)
                with c1:
                    case_num = st.text_input("رقم القضية")
                    case_type = st.selectbox("النوع", CASE_TYPES)
                    status = st.selectbox("الحالة", CASE_STATUSES)
                with c2:
                    opponents = st.text_input("الخصوم")
                    court = st.text_input("المحكمة / الجهة")
                    circuit = st.text_input("الدائرة")
                with c3:
                    session_date = date_input_dmy("تاريخ جلسة", value=datetime.now().date(), max_value=date(2100, 12, 31), key="new_case_session_date")
                    assigned_to = st.selectbox("المسند إليه", member_names, index=member_names.index(default_owner) if default_owner in member_names else 0)
                    ruling = st.text_input("الحكم أو القرار")
                summary = st.text_area("ملخص الموضوع", height=90)
                notes = st.text_area("ملاحظات", height=70)
                docs = st.file_uploader("مستندات القضية", accept_multiple_files=True, key="new_case_docs")
                if st.form_submit_button("💾 حفظ القضية", type="primary", use_container_width=True):
                    if not case_num.strip():
                        st.error("أدخل رقم القضية أولاً")
                    else:
                        data = load_data()
                        uploaded_docs = []
                        for f in docs or []:
                            raw_b = f.getvalue()
                            cloud_p, _ = cloud_upload_file(raw_b, f.name, f.type or "application/octet-stream")
                            uploaded_docs.append({
                                "name": f.name,
                                "type": f.type or "application/octet-stream",
                                "data": base64.b64encode(raw_b).decode("utf-8"),
                                "cloud_path": cloud_p or "",
                            })
                        owner_user = next((u for u in users if u.get("username") == assigned_to), user)
                        new_case = ensure_case_defaults({
                            "id": secrets.token_hex(6),
                            "رقم_القضية": case_num.strip(),
                            "النوع": case_type,
                            "الخصوم": opponents.strip(),
                            "المحكمة": court.strip(),
                            "الدائرة": circuit.strip(),
                            "تاريخ_جلسة": fmt_date_ar(session_date),
                            "الحالة": status,
                            "الحكم_أو_القرار": ruling.strip(),
                            "ملخص_الموضوع": summary.strip(),
                            "ملاحظات": notes.strip(),
                            "user_id": owner_user.get("id"),
                            "assigned_to": assigned_to,
                            "جلسات": [{"تاريخ": fmt_date_ar(session_date), "قرار": ruling.strip() or status, "ملاحظات": notes.strip()}] if session_date else [],
                            "مستندات": uploaded_docs,
                        })
                        data.setdefault("cases", []).append(new_case)
                        save_data(data)
                        st.success("تم حفظ القضية بنجاح")
                        st.rerun()

    st.markdown("---")
    f1, f2, f3 = st.columns([2, 1.2, 1.2])
    with f1:
        q = st.text_input("🔎 بحث في القضايا", placeholder="رقم القضية أو الخصوم أو المحكمة أو الموضوع")
    with f2:
        type_filter = st.selectbox("النوع", ["الكل"] + CASE_TYPES)
    with f3:
        status_filter = st.selectbox("الحالة", ["الكل"] + CASE_STATUSES)

    cases = get_visible_cases()
    if q.strip():
        ql = q.strip().lower()
        cases = [c for c in cases if ql in " ".join(str(c.get(k, "")) for k in ("رقم_القضية", "النوع", "الخصوم", "المحكمة", "الدائرة", "ملخص_الموضوع", "ملاحظات", "assigned_to")).lower()]
    if type_filter != "الكل":
        cases = [c for c in cases if c.get("النوع") == type_filter]
    if status_filter != "الكل":
        cases = [c for c in cases if c.get("الحالة") == status_filter]

    st.markdown(f"### عدد القضايا الظاهرة: {len(cases)}")
    for c in cases:
        cid = c["id"]
        with st.container(border=True):
            st.markdown(f"### ⚖️ قضية رقم: {c.get('رقم_القضية','')} — {c.get('النوع','')}")
            st.write(f"**الخصوم:** {c.get('الخصوم','') or '—'} | **المحكمة:** {c.get('المحكمة','') or '—'} | **الدائرة:** {c.get('الدائرة','') or '—'}")
            st.write(f"**تاريخ الجلسة:** {fmt_date_ar(c.get('تاريخ_جلسة'), '—')} | **الحالة:** {c.get('الحالة','') or '—'} | **المسند إليه:** {c.get('assigned_to','') or '—'}")
            if c.get("الحكم_أو_القرار"):
                st.write(f"**الحكم أو القرار:** {c.get('الحكم_أو_القرار')}")
            if c.get("ملخص_الموضوع"):
                st.write(f"**ملخص الموضوع:** {c.get('ملخص_الموضوع')}")
            if c.get("ملاحظات"):
                st.write(f"**ملاحظات:** {c.get('ملاحظات')}")

            for di, doc in enumerate(c.get("مستندات") or []):
                try:
                    st.download_button(
                        f"📎 تحميل مستند: {doc.get('name','مستند')}",
                        data=base64.b64decode(doc.get("data") or ""),
                        file_name=doc.get("name") or "مستند",
                        mime=doc.get("type") or "application/octet-stream",
                        key=f"dl_case_doc_{cid}_{di}",
                        use_container_width=True,
                    )
                except Exception:
                    pass

            if current_user_can_edit_case(c):
                with st.expander("✏️ تعديل القضية / إضافة جلسة أو مستند"):
                    with st.form(f"edit_case_{cid}"):
                        ec1, ec2, ec3 = st.columns(3)
                        with ec1:
                            e_num = st.text_input("رقم القضية", value=c.get("رقم_القضية", ""))
                            e_type = st.selectbox("النوع", CASE_TYPES, index=CASE_TYPES.index(c.get("النوع")) if c.get("النوع") in CASE_TYPES else 0)
                            e_status = st.selectbox("الحالة", CASE_STATUSES, index=CASE_STATUSES.index(c.get("الحالة")) if c.get("الحالة") in CASE_STATUSES else 0)
                        with ec2:
                            e_opp = st.text_input("الخصوم", value=c.get("الخصوم", ""))
                            e_court = st.text_input("المحكمة", value=c.get("المحكمة", ""))
                            e_circ = st.text_input("الدائرة", value=c.get("الدائرة", ""))
                        with ec3:
                            e_date = date_input_dmy("تاريخ جلسة", value=parse_date_safe(c.get("تاريخ_جلسة")) or datetime.now().date(), max_value=date(2100, 12, 31), key=f"ed_case_dt_{cid}")
                            e_assign = st.selectbox("المسند إليه", member_names, index=member_names.index(c.get("assigned_to")) if c.get("assigned_to") in member_names else 0)
                            e_ruling = st.text_input("الحكم أو القرار", value=c.get("الحكم_أو_القرار", ""))
                        e_sum = st.text_area("ملخص الموضوع", value=c.get("ملخص_الموضوع", ""))
                        e_notes = st.text_area("ملاحظات", value=c.get("ملاحظات", ""))
                        add_session_flag = st.checkbox("إضافة هذه الجلسة إلى سجل جلسات القضية", value=False)
                        new_docs = st.file_uploader("إضافة مستندات جديدة", accept_multiple_files=True, key=f"add_docs_{cid}")
                        can_del_case = (user.get("role") == "admin" or is_assistant_admin(user)) or current_user_has_permission("can_delete_cases")
                        del_case = st.checkbox("حذف هذه القضية", value=False, disabled=not can_del_case)
                        if st.form_submit_button("💾 حفظ التعديلات", type="primary", use_container_width=True):
                            all_data = load_data()
                            for idx, item in enumerate(all_data.get("cases", [])):
                                if item.get("id") == cid:
                                    if del_case:
                                        if not can_del_case:
                                            st.error("لا تملك صلاحية حذف القضايا.")
                                            st.stop()
                                        for d_item in item.get("مستندات") or []:
                                            if isinstance(d_item, dict) and d_item.get("cloud_path"):
                                                cloud_delete_file(d_item["cloud_path"])
                                        cloud_delete_record("cases", cid)
                                        all_data["cases"].pop(idx)
                                    else:
                                        item["رقم_القضية"] = e_num.strip()
                                        item["النوع"] = e_type
                                        item["الحالة"] = e_status
                                        item["الخصوم"] = e_opp.strip()
                                        item["المحكمة"] = e_court.strip()
                                        item["الدائرة"] = e_circ.strip()
                                        item["تاريخ_جلسة"] = fmt_date_ar(e_date)
                                        item["assigned_to"] = e_assign
                                        item["الحكم_أو_القرار"] = e_ruling.strip()
                                        item["ملخص_الموضوع"] = e_sum.strip()
                                        item["ملاحظات"] = e_notes.strip()
                                        if add_session_flag:
                                            item.setdefault("جلسات", []).append({
                                                "تاريخ": fmt_date_ar(e_date),
                                                "قرار": e_ruling.strip() or e_status,
                                                "ملاحظات": e_notes.strip(),
                                            })
                                        for f in new_docs or []:
                                            raw_b = f.getvalue()
                                            cloud_p, _ = cloud_upload_file(raw_b, f.name, f.type or "application/octet-stream")
                                            item.setdefault("مستندات", []).append({
                                                "name": f.name,
                                                "type": f.type or "application/octet-stream",
                                                "data": base64.b64encode(raw_b).decode("utf-8"),
                                                "cloud_path": cloud_p or "",
                                            })
                                    break
                            save_data(all_data)
                            st.success("تم تحديث بيانات القضية")
                            st.rerun()

def render_sessions_page():
    user = st.session_state.get("user") or {}
    if user.get("role") != "admin" and not is_assistant_admin(user) and not current_user_has_permission("can_view_cases"):
        st.error("لا تملك صلاحية عرض الجلسات وفق الصلاحيات المقررة لحسابك.")
        render_return_home_button("⬅️ العودة للرئيسية", key="sessions_no_perm_home")
        return

    st.markdown("<h2 style='color:#D4AF37;text-align:center;'>📅 جدول الجلسات</h2>", unsafe_allow_html=True)
    render_return_home_button("⬅️ العودة للرئيسية", key="sessions_top_home")

    cases = get_visible_cases()
    rows = []
    for c in cases:
        for d in _case_session_dates(c):
            rows.append({
                "تاريخ الجلسة": fmt_date_ar(d),
                "_sort": d.strftime("%Y-%m-%d"),
                "رقم القضية": c.get("رقم_القضية", ""),
                "النوع": c.get("النوع", ""),
                "الخصوم": c.get("الخصوم", ""),
                "المحكمة": c.get("المحكمة", ""),
                "الدائرة": c.get("الدائرة", ""),
                "الحالة": c.get("الحالة", ""),
                "المسند إليه": c.get("assigned_to", ""),
                "الحكم أو القرار": c.get("الحكم_أو_القرار", ""),
            })
    rows.sort(key=lambda r: r["_sort"])

    col1, col2 = st.columns(2)
    with col1:
        day_filter = date_input_dmy("تصفية بتاريخ محدد", value=datetime.now().date(), max_value=date(2100, 12, 31), key="sessions_filter_date")
    with col2:
        only_day = st.checkbox("عرض جلسات هذا التاريخ فقط", value=False)

    if only_day:
        target_str = fmt_date_ar(day_filter)
        rows = [r for r in rows if r["تاريخ الجلسة"] == target_str]

    if not rows:
        st.info("لا توجد جلسات مسجلة مطابقة للبحث.")
        return

    df = pd.DataFrame([{k: v for k, v in r.items() if k != "_sort"} for r in rows])
    st.dataframe(df, use_container_width=True, hide_index=True)

def render_library_page():
    user = st.session_state.get("user") or {}
    if user.get("role") != "admin" and not is_assistant_admin(user) and not current_user_has_permission("can_view_cases"):
        st.error("لا تملك صلاحية عرض المكتبة القانونية وفق الصلاحيات المقررة لحسابك.")
        render_return_home_button("⬅️ العودة للرئيسية", key="library_no_perm_home")
        return

    st.markdown("<h2 style='color:#D4AF37;text-align:center;'>📚 المكتبة القانونية</h2>", unsafe_allow_html=True)
    render_return_home_button("⬅️ العودة للرئيسية", key="library_top_home")

    data = load_data()
    items = data.get("library", [])

    with st.expander("➕ إضافة مرجع أو مستند قانوني"):
        if user.get("role") != "admin" and not is_assistant_admin(user) and not current_user_has_permission("can_edit_cases"):
            st.warning("حسابك يملك صلاحية العرض فقط ولا يملك صلاحية إضافة مراجع جديدة.")
        else:
            with st.form("add_library_item"):
                title = st.text_input("عنوان المرجع أو المستند")
                category = st.selectbox("التصنيف", ["قوانين ولوائح", "أحكام ومبادئ قضائية", "نماذج ومذكرات", "فتاوى وآراء قانونية", "أخرى"])
                notes = st.text_area("ملخص أو ملاحظات", height=80)
                file = st.file_uploader("رفع الملف", key="lib_file_upload")
                if st.form_submit_button("💾 حفظ في المكتبة", type="primary", use_container_width=True):
                    if not title.strip():
                        st.error("أدخل عنوان المرجع أولاً")
                    else:
                        raw_b = file.getvalue() if file else b""
                        cloud_p = ""
                        if raw_b and file:
                            cloud_p, _ = cloud_upload_file(raw_b, file.name, file.type or "application/octet-stream")
                        items.append({
                            "id": secrets.token_hex(6),
                            "title": title.strip(),
                            "category": category,
                            "notes": notes.strip(),
                            "file_name": file.name if file else "",
                            "file_type": file.type if file else "application/octet-stream",
                            "file_data": base64.b64encode(raw_b).decode("utf-8") if raw_b else "",
                            "cloud_path": cloud_p or "",
                            "user_id": user.get("id"),
                            "added_by": user.get("username", ""),
                            "created_at": fmt_date_ar(datetime.now()),
                        })
                        data["library"] = items
                        save_data(data)
                        st.success("تمت إضافة المرجع إلى المكتبة القانونية")
                        st.rerun()

    q = st.text_input("🔎 بحث في المكتبة القانونية", placeholder="ابحث بالعنوان أو التصنيف أو الملاحظات")
    filtered = items
    if q.strip():
        ql = q.strip().lower()
        filtered = [x for x in filtered if ql in f"{x.get('title','')} {x.get('category','')} {x.get('notes','')}".lower()]

    for it in reversed(filtered):
        with st.container(border=True):
            st.markdown(f"### 📖 {it.get('title','')}")
            st.write(f"**التصنيف:** {it.get('category','')} | **بواسطة:** {it.get('added_by','')} | **التاريخ:** {fmt_date_ar(it.get('created_at'), '')}")
            if it.get("notes"):
                st.write(it.get("notes"))
            c1, c2 = st.columns([3, 1])
            with c1:
                if it.get("file_data"):
                    try:
                        st.download_button(
                            f"⬇️ تحميل {it.get('file_name') or 'الملف'}",
                            data=base64.b64decode(it["file_data"]),
                            file_name=it.get("file_name") or "document",
                            mime=it.get("file_type") or "application/octet-stream",
                            key=f"dl_lib_{it['id']}",
                            use_container_width=True,
                        )
                    except Exception:
                        pass
            with c2:
                can_del_lib = (user.get("role") == "admin" or is_assistant_admin(user)) or current_user_has_permission("can_delete_cases")
                if can_del_lib and st.button("🗑️ حذف", key=f"del_lib_{it['id']}", use_container_width=True):
                    if it.get("cloud_path"):
                        cloud_delete_file(it["cloud_path"])
                    cloud_delete_record("library", it["id"])
                    data["library"] = [x for x in data.get("library", []) if x.get("id") != it["id"]]
                    save_data(data)
                    st.rerun()
# === نهاية الجزء 2 من 3 ===
# === بداية الجزء 3 من 3 ===
def save_member_phone_and_photo(user_id, phone_value, photo_file=None):
    users = load_users()
    updated = False
    clean_phone = str(phone_value or "").strip()
    for u in users:
        if u.get("id") == user_id:
            u["phone"] = clean_phone
            u["mobile"] = clean_phone
            if photo_file is not None:
                u["profile_photo"] = base64.b64encode(photo_file.getvalue()).decode("utf-8")
                u["profile_photo_type"] = photo_file.type or "image/png"
            updated = True
            if st.session_state.get("user", {}).get("id") == user_id:
                st.session_state.user = dict(u)
            break
    if updated:
        return save_users(users)
    return False

# ===== التقارير والتصدير (Word / PDF / Excel) =====
def _find_arabic_font():
    candidates = [
        str(BASE_DIR / "Cairo-Regular.ttf"),
        str(BASE_DIR / "fonts" / "Cairo-Regular.ttf"),
        str(Path(LOCAL_DATA_DIR) / "Cairo-Regular.ttf"),
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return None

def _set_cell_rtl_and_shading(cell, fill_hex=None, bold=False, font_size=11):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(OxmlElement("w:tcW"))
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if fill_hex:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill_hex.replace("#", ""))
        tcPr.append(shd)
    for p in cell.paragraphs:
        p.paragraph_format.bidi = True
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        for r in p.runs:
            r.font.name = "Cairo"
            r.font.size = Pt(font_size)
            r.bold = bold

def build_cases_docx(cases, title="تقرير القضايا والجلسات"):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.6)
    section.right_margin = Inches(0.6)

    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Cairo"
    normal_style.font.size = Pt(11)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.bidi = True
    r_t = p_title.add_run(title)
    r_t.font.name = "Cairo"
    r_t.font.size = Pt(16)
    r_t.bold = True

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.bidi = True
    r_s = p_sub.add_run(f"تاريخ التقرير: {fmt_date_ar(datetime.now())} — إجمالي القضايا: {len(cases)}")
    r_s.font.name = "Cairo"
    r_s.font.size = Pt(10)

    headers = ["رقم القضية", "النوع", "الخصوم", "المحكمة / الدائرة", "تاريخ الجلسة", "الحالة", "الحكم أو القرار"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_cells = table.rows[0].cells
    for idx, h_text in enumerate(headers):
        hdr_cells[idx].text = h_text
        _set_cell_rtl_and_shading(hdr_cells[idx], fill_hex="162642", bold=True, font_size=10)
        for p in hdr_cells[idx].paragraphs:
            for r in p.runs:
                r.font.color.rgb = None

    for c in cases:
        row_cells = table.add_row().cells
        court_circ = " / ".join([x for x in [c.get("المحكمة", ""), c.get("الدائرة", "")] if x])
        vals = [
            str(c.get("رقم_القضية", "")),
            str(c.get("النوع", "")),
            str(c.get("الخصوم", "")),
            court_circ,
            fmt_date_ar(c.get("تاريخ_جلسة"), ""),
            str(c.get("الحالة", "")),
            str(c.get("الحكم_أو_القرار", "") or c.get("ملخص_الموضوع", "")),
        ]
        for idx, val in enumerate(vals):
            row_cells[idx].text = val
            _set_cell_rtl_and_shading(row_cells[idx], fill_hex=None, bold=(idx == 0), font_size=10)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

def build_cases_pdf(cases, title="تقرير القضايا والجلسات"):
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()

    font_path = _find_arabic_font()
    if font_path:
        pdf.add_font("ArFont", "", font_path, uni=True)
        pdf.set_font("ArFont", "", 14)
    else:
        pdf.set_font("Helvetica", "B", 14)

    pdf.cell(0, 10, fix_arabic(title), ln=True, align="C")
    pdf.set_font("ArFont" if font_path else "Helvetica", "", 10)
    pdf.cell(0, 7, fix_arabic(f"تاريخ التقرير: {fmt_date_ar(datetime.now())} | العدد: {len(cases)}"), ln=True, align="C")
    pdf.ln(3)

    col_widths = [32, 28, 52, 52, 30, 28, 70]
    headers = ["رقم القضية", "النوع", "الخصوم", "المحكمة / الدائرة", "تاريخ الجلسة", "الحالة", "الحكم / القرار"]

    pdf.set_fill_color(22, 38, 66)
    pdf.set_text_color(255, 255, 255)
    for w, h in zip(col_widths, headers):
        pdf.cell(w, 9, fix_arabic(h), border=1, align="C", fill=True)
    pdf.ln()

    pdf.set_text_color(0, 0, 0)
    for c in cases:
        court_circ = " / ".join([x for x in [c.get("المحكمة", ""), c.get("الدائرة", "")] if x])
        row_vals = [
            str(c.get("رقم_القضية", ""))[:22],
            str(c.get("النوع", ""))[:18],
            str(c.get("الخصوم", ""))[:36],
            court_circ[:36],
            fmt_date_ar(c.get("تاريخ_جلسة"), "")[:14],
            str(c.get("الحالة", ""))[:18],
            str(c.get("الحكم_أو_القرار", "") or c.get("ملخص_الموضوع", ""))[:46],
        ]
        for w, val in zip(col_widths, row_vals):
            pdf.cell(w, 8, fix_arabic(val), border=1, align="C")
        pdf.ln()

    out = pdf.output(dest="S")
    if isinstance(out, (bytes, bytearray)):
        return bytes(out)
    return str(out).encode("latin1", errors="ignore")

def build_cases_excel(cases):
    rows = []
    for c in cases:
        rows.append({
            "رقم القضية": c.get("رقم_القضية", ""),
            "النوع": c.get("النوع", ""),
            "الخصوم": c.get("الخصوم", ""),
            "المحكمة": c.get("المحكمة", ""),
            "الدائرة": c.get("الدائرة", ""),
            "تاريخ الجلسة": fmt_date_ar(c.get("تاريخ_جلسة"), ""),
            "الحالة": c.get("الحالة", ""),
            "المسند إليه": c.get("assigned_to", ""),
            "الحكم أو القرار": c.get("الحكم_أو_القرار", ""),
            "ملخص الموضوع": c.get("ملخص_الموضوع", ""),
            "ملاحظات": c.get("ملاحظات", ""),
        })
    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="القضايا")
        ws = writer.sheets["القضايا"]
        ws.sheet_view.rightToLeft = True
        header_fill = PatternFill(start_color="162642", end_color="162642", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
    buf.seek(0)
    return buf.getvalue()

def render_reports_page():
    user = st.session_state.get("user") or {}
    if user.get("role") != "admin" and not is_assistant_admin(user) and not current_user_has_permission("can_export_reports"):
        st.error("لا تملك صلاحية تصدير التقارير وفق الصلاحيات المقررة لحسابك.")
        render_return_home_button("⬅️ العودة للرئيسية", key="reports_no_perm_home")
        return

    st.markdown("<h2 style='color:#D4AF37;text-align:center;'>📊 التقارير والتصدير</h2>", unsafe_allow_html=True)
    render_return_home_button("⬅️ العودة للرئيسية", key="reports_top_home")

    cases = get_visible_cases()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("إجمالي القضايا", len(cases))
    c2.metric("قضايا متداولة", sum(1 for c in cases if c.get("الحالة") == "متداولة"))
    c3.metric("حجز للحكم", sum(1 for c in cases if c.get("الحالة") == "حجز للحكم"))
    c4.metric("منتهية بحكم", sum(1 for c in cases if c.get("الحالة") == "منتهية بحكم"))

    st.markdown("---")
    f1, f2 = st.columns(2)
    with f1:
        rep_type = st.selectbox("تصفية حسب النوع", ["الكل"] + CASE_TYPES, key="rep_type_sel")
    with f2:
        rep_status = st.selectbox("تصفية حسب الحالة", ["الكل"] + CASE_STATUSES, key="rep_status_sel")

    filtered = cases
    if rep_type != "الكل":
        filtered = [c for c in filtered if c.get("النوع") == rep_type]
    if rep_status != "الكل":
        filtered = [c for c in filtered if c.get("الحالة") == rep_status]

    if filtered:
        preview_df = pd.DataFrame([
            {
                "رقم القضية": c.get("رقم_القضية", ""),
                "النوع": c.get("النوع", ""),
                "الخصوم": c.get("الخصوم", ""),
                "المحكمة": c.get("المحكمة", ""),
                "تاريخ الجلسة": fmt_date_ar(c.get("تاريخ_جلسة"), ""),
                "الحالة": c.get("الحالة", ""),
                "الحكم أو القرار": c.get("الحكم_أو_القرار", ""),
            }
            for c in filtered
        ])
        st.dataframe(preview_df, use_container_width=True, hide_index=True)
    else:
        st.info("لا توجد قضايا مطابقة للتصفية الحالية.")

    st.markdown("### ⬇️ تحميل التقرير")
    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button(
            "📄 تصدير Word (.docx)",
            data=build_cases_docx(filtered),
            file_name="تقرير_القضايا.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="dl_rep_docx",
        )
    with d2:
        st.download_button(
            "📕 تصدير PDF (.pdf)",
            data=build_cases_pdf(filtered),
            file_name="تقرير_القضايا.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="dl_rep_pdf",
        )
    with d3:
        st.download_button(
            "📗 تصدير Excel (.xlsx)",
            data=build_cases_excel(filtered),
            file_name="تقرير_القضايا.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_rep_xlsx",
        )

# ===== الصفحة الرئيسية =====
def home_page():
    render_top_bar()
    show_admin_age_60_alert()
    show_top_notices()
    st.markdown("<h1 style='color:#D4AF37;text-align:center;margin:10px 0 4px;'>⚖️ منظومة إدارة القضايا</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#e2e8f0;text-align:center;margin-bottom:18px;font-weight:700;'>إدارة القضايا والجلسات والمكتبة القانونية والتقارير والرسائل</p>", unsafe_allow_html=True)

    cards = [
        ("📁", "القضايا", "القضايا"),
        ("📅", "الجلسات", "الجلسات"),
        ("📚", "المكتبة القانونية", "المكتبة القانونية"),
        ("📊", "التقارير", "التقارير"),
        ("📢", "الرسائل والبيانات", "الرسائل"),
    ]
    user = st.session_state.get("user") or {}
    if user.get("role") == "admin" or is_assistant_admin(user):
        cards.append(("⚖️", "لوحة التحكم", "لوحة التحكم"))
    elif user.get("id"):
        cards.append(("⚙️", "لوحة تحكمى", "لوحة تحكمى"))

    cols = st.columns(3)
    for i, (icon, label, target_page) in enumerate(cards):
        with cols[i % 3]:
            st.markdown(
                f"""<div class="judicial-card"><div class="judicial-card-icon">{icon}</div><div class="judicial-card-title">{html.escape(label)}</div></div>""",
                unsafe_allow_html=True,
            )
            if st.button(f"فتح {label}", key=f"home_open_{target_page}_{i}", use_container_width=True):
                st.session_state.page = target_page
                st.rerun()

    st.markdown("---")
    global_people_search()
    show_member_directory()

# ===== التوجيه وتشغيل التطبيق =====
if "page" not in st.session_state:
    st.session_state.page = "login"
if "user" not in st.session_state:
    st.session_state.user = None
if "role" not in st.session_state:
    st.session_state.role = None

if not st.session_state.get("user"):
    login_page()
else:
    render_site_audio_player()
    page = st.session_state.get("page", "الرئيسية")
    if page == "set_password":
        set_password_page()
    elif page == "extract_member":
        extract_member_page()
    elif page == "تعديل_بيانات_العضويات":
        edit_memberships_page()
    elif page == "لوحة التحكم":
        admin_control_page()
    elif page == "لوحة تحكمى":
        personal_control_page()
    elif page == "الرسائل":
        render_top_bar()
        messaging_page()
    elif page == "القضايا":
        render_top_bar()
        render_cases_page()
    elif page == "الجلسات":
        render_top_bar()
        render_sessions_page()
    elif page == "المكتبة القانونية":
        render_top_bar()
        render_library_page()
    elif page == "التقارير":
        render_top_bar()
        render_reports_page()
    else:
        home_page()
# === نهاية الجزء 3 من 3 ===
