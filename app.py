import streamlit as st
import requests
import time
import json
import os
from datetime import datetime, timedelta, date
from collections import defaultdict
from enum import Enum
from typing import Optional
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Биллибоба",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
/* ── Hide sidebar ───────────────────────────────────────────────────────── */
[data-testid="collapsedControl"] { display: none; }
[data-testid="stSidebar"]        { display: none; }

/* ── Layout ─────────────────────────────────────────────────────────────── */
.block-container {
    padding-top: 1.4rem !important;
    padding-left: 2.5rem !important;
    padding-right: 2.5rem !important;
    max-width: 1440px !important;
}

/* ── Typography ─────────────────────────────────────────────────────────── */
h1, h2, h3 { letter-spacing: -0.3px; }
h2 { font-size: 1.6rem !important; font-weight: 700 !important; }

/* ── Tabs ────────────────────────────────────────────────────────────────── */
[data-testid="stTabs"] [role="tablist"] {
    gap: 4px;
    border-bottom: 2px solid rgba(255,255,255,0.07);
    padding-bottom: 0;
}
[data-testid="stTabs"] [role="tab"] {
    font-size: 0.88rem;
    font-weight: 500;
    padding: 8px 18px;
    border-radius: 8px 8px 0 0;
    color: rgba(255,255,255,0.55);
    transition: color 0.2s, background 0.2s;
}
[data-testid="stTabs"] [role="tab"]:hover {
    color: rgba(255,255,255,0.85);
    background: rgba(255,255,255,0.04);
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    font-weight: 700;
    color: #a78bfa;
}

/* ── Metric cards ────────────────────────────────────────────────────────── */
[data-testid="stMetric"] {
    background: linear-gradient(135deg, #1e1e2e 0%, #252540 100%);
    border-radius: 14px;
    padding: 18px 22px;
    border: 1px solid rgba(255,255,255,0.07);
    box-shadow: 0 4px 16px rgba(0,0,0,0.25);
    transition: transform 0.15s, box-shadow 0.15s;
}
[data-testid="stMetric"]:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 24px rgba(0,0,0,0.35);
}
[data-testid="stMetricLabel"] {
    font-size: 0.75rem !important;
    opacity: 0.6;
    text-transform: uppercase;
    letter-spacing: 0.7px;
    font-weight: 600;
}
[data-testid="stMetricValue"] {
    font-size: 2.2rem !important;
    font-weight: 800 !important;
    color: #e2e8f0;
}

/* ── Buttons ─────────────────────────────────────────────────────────────── */
button[kind="primary"], [data-testid="stButton"] > button[kind="primary"] {
    background: linear-gradient(135deg, #7c3aed 0%, #5b21b6 100%) !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    letter-spacing: 0.2px;
    box-shadow: 0 4px 14px rgba(124,58,237,0.35) !important;
    transition: all 0.2s !important;
}
button[kind="primary"]:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(124,58,237,0.5) !important;
}
button[kind="secondary"], [data-testid="stButton"] > button[kind="secondary"] {
    border-radius: 10px !important;
    font-weight: 600 !important;
    border: 1px solid rgba(255,255,255,0.15) !important;
}

/* ── Cards / containers ──────────────────────────────────────────────────── */
[data-testid="stVerticalBlockBorderWrapper"] > div {
    border-radius: 14px !important;
    border: 1px solid rgba(255,255,255,0.08) !important;
    background: rgba(255,255,255,0.02) !important;
    box-shadow: 0 2px 10px rgba(0,0,0,0.15);
}

/* ── Expanders ───────────────────────────────────────────────────────────── */
[data-testid="stExpander"] {
    border-radius: 10px !important;
    border: 1px solid rgba(255,255,255,0.08) !important;
}

/* ── Info / success / error boxes ───────────────────────────────────────── */
[data-testid="stAlert"] {
    border-radius: 10px !important;
    border-left-width: 4px !important;
}

/* ── DataFrames ──────────────────────────────────────────────────────────── */
[data-testid="stDataFrame"] {
    border-radius: 12px !important;
    overflow: hidden;
    border: 1px solid rgba(255,255,255,0.07) !important;
}

/* ── Divider ─────────────────────────────────────────────────────────────── */
hr { border-color: rgba(255,255,255,0.07) !important; margin: 1rem 0 !important; }

/* ── Inputs ──────────────────────────────────────────────────────────────── */
[data-testid="stTextInput"] > div > div,
[data-testid="stDateInput"]  > div > div {
    border-radius: 8px !important;
}

/* ── Status pills ────────────────────────────────────────────────────────── */
.pill {
    display: inline-block;
    padding: 3px 11px;
    border-radius: 20px;
    font-size: 0.78rem;
    font-weight: 600;
    line-height: 1.6;
}
.pill-open    { background:rgba(239,68,68,.15);  color:#f87171; }
.pill-sent    { background:rgba(59,130,246,.15); color:#60a5fa; }
.pill-pass    { background:rgba(245,158,11,.15); color:#fbbf24; }
.pill-handled { background:rgba(16,185,129,.15); color:#34d399; }
.pill-skipped { background:rgba(156,163,175,.15);color:#9ca3af; }
.pill-intern  { background:rgba(251,191,36,.15); color:#fbbf24; }

/* ── Info bar (period) ───────────────────────────────────────────────────── */
.info-bar {
    background: linear-gradient(90deg,rgba(124,58,237,.12),rgba(91,33,182,.06));
    border: 1px solid rgba(124,58,237,.25);
    border-radius: 10px;
    padding: 10px 18px;
    font-size: 0.88rem;
    color: rgba(255,255,255,0.75);
    margin-bottom: 0.8rem;
}
.info-bar b { color: #c4b5fd; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
#  TYPES
# ─────────────────────────────────────────────────────────────────────────────

class ErrorType(Enum):
    NO_REPORT          = "no_report"
    NO_ABSENCE_COMMENT = "no_abs_comment"


# ─────────────────────────────────────────────────────────────────────────────
#  PERSISTENT SETTINGS (saved to config.json)
# ─────────────────────────────────────────────────────────────────────────────

CONFIG_FILE   = os.path.join(os.path.dirname(__file__), "config.json")
HISTORY_FILE  = os.path.join(os.path.dirname(__file__), "history.json")


# ─────────────────────────────────────────────────────────────────────────────
#  HISTORY — Google Sheets backend (falls back to local JSON for dev)
# ─────────────────────────────────────────────────────────────────────────────

_HIST_HEADERS = [
    "id", "checked_at", "period_from", "period_to", "date", "teacher",
    "student_tag", "error_type", "error_description", "count", "students",
    "status", "reviewer", "reviewer_comment", "updated_at",
]

def _hist_row_to_dict(row: list) -> dict:
    if len(row) < len(_HIST_HEADERS):
        row = list(row) + [""] * (len(_HIST_HEADERS) - len(row))
    r = dict(zip(_HIST_HEADERS, row))
    try:    r["students"] = json.loads(r.get("students") or "[]")
    except Exception: r["students"] = []
    try:    r["count"] = int(r["count"]) if r["count"] else 0
    except Exception: r["count"] = 0
    return r

def _hist_dict_to_row(rec: dict) -> list:
    return [
        rec.get("id", ""),
        rec.get("checked_at", ""),
        rec.get("period_from", ""),
        rec.get("period_to", ""),
        rec.get("date", ""),
        rec.get("teacher", ""),
        rec.get("student_tag", "Все"),
        rec.get("error_type", ""),
        rec.get("error_description", ""),
        str(rec.get("count", 0)),
        json.dumps(rec.get("students", []), ensure_ascii=False),
        rec.get("status", "open"),
        rec.get("reviewer", ""),
        rec.get("reviewer_comment", ""),
        rec.get("updated_at", ""),
    ]

@st.cache_resource(show_spinner=False)
def _get_gspread_client():
    """Shared authenticated gspread client (cached as resource)."""
    try:
        import gspread as _gs
        from google.oauth2.service_account import Credentials as _C
        _scopes = ["https://www.googleapis.com/auth/spreadsheets",
                   "https://www.googleapis.com/auth/drive"]
        try:
            sa = dict(st.secrets["gcp_service_account"])
        except Exception:
            _cj = load_config().get("creds_json", "")
            if not _cj:
                return None
            sa = json.loads(_cj)
        return _gs.authorize(_C.from_service_account_info(sa, scopes=_scopes))
    except Exception:
        return None

@st.cache_resource(show_spinner=False)
def _get_history_worksheet(sheet_id: str):
    """Cached worksheet object — re-created only when sheet_id changes."""
    gc = _get_gspread_client()
    if not gc or not sheet_id:
        return None
    try:
        return gc.open_by_key(sheet_id).sheet1
    except Exception:
        return None

def _get_history_sheet():
    """Return (gspread Worksheet | None, sheet_id | None)."""
    cfg_ = load_config()
    sid  = cfg_.get("history_sheet_id", "") or st.secrets.get("history_sheet_id", "")
    if not sid:
        return None, None
    ws = _get_history_worksheet(sid)
    return (ws, sid) if ws is not None else (None, None)

# ── Supabase (primary history backend) ───────────────────────────────────────

def _read_secret(key: str) -> str:
    """Read a secret safely from st.secrets or config."""
    try:
        val = st.secrets[key]
        return str(val) if val else ""
    except Exception:
        pass
    return load_config().get(key, "")

@st.cache_resource(show_spinner=False)
def _get_supabase_client():
    """Cached Supabase client. Returns (client, error_str)."""
    try:
        from supabase import create_client as _sb_create
    except ImportError as e:
        return None, f"Пакет supabase не установлен: {e}"
    try:
        _url = _read_secret("supabase_url")
        _key = _read_secret("supabase_key")
        if not _url:
            # show what keys ARE visible for debugging
            try:
                _visible = list(st.secrets.keys())
            except Exception:
                _visible = []
            return None, f"supabase_url не найден. Видимые ключи в Secrets: {_visible}"
        if not _key:
            return None, "supabase_key не найден в Secrets"
        client = _sb_create(_url, _key)
        return client, None
    except Exception as e:
        return None, str(e)

def _hist_dict_to_sb_row(rec: dict) -> dict:
    """Convert history dict to Supabase row (students → JSON string)."""
    row = dict(rec)
    if isinstance(row.get("students"), list):
        row["students"] = json.dumps(row["students"], ensure_ascii=False)
    row["count"] = int(row.get("count", 0))
    return row

def _sb_row_to_hist_dict(row: dict) -> dict:
    """Convert Supabase row back to history dict (students JSON → list)."""
    row = dict(row)
    if isinstance(row.get("students"), str):
        try:    row["students"] = json.loads(row["students"])
        except: row["students"] = []
    elif not isinstance(row.get("students"), list):
        row["students"] = []
    row["count"] = int(row.get("count", 0))
    return row

def mirror_to_sheets(records: list):
    """Write full history to Google Sheets (mirror/backup)."""
    ws, _ = _get_history_sheet()
    if ws is None:
        return False, "Sheets не подключён"
    try:
        all_values = [_HIST_HEADERS] + [_hist_dict_to_row(r) for r in records]
        ws.clear()
        if all_values:
            ws.update("A1", all_values)
        return True, f"Синхронизировано {len(records)} записей"
    except Exception as e:
        return False, str(e)

def create_history_sheet() -> tuple:
    """Create a new Google Sheet for history. Returns (sheet_id, url) or raises."""
    gc = _get_gspread_client()
    if not gc:
        raise RuntimeError("Нет подключения к Google (проверь credentials в настройках)")
    sh = gc.create("📚 Проверка отчётов — История")
    ws = sh.sheet1
    ws.update("A1", [_HIST_HEADERS])
    ws.freeze(rows=1)
    # Make it accessible to anyone with the link (viewer)
    gc.insert_permission(sh.id, None, perm_type="anyone", role="reader")
    return sh.id, sh.url

def load_history() -> list:
    # 1. Supabase (primary)
    sb, _ = _get_supabase_client()
    if sb is not None:
        try:
            res = sb.table("history").select("*").execute()
            if res.data is not None:
                return [_sb_row_to_hist_dict(r) for r in res.data]
        except Exception:
            pass
    # 2. Google Sheets (fallback)
    ws, _ = _get_history_sheet()
    if ws is not None:
        try:
            rows = ws.get_all_values()
            if rows and len(rows) >= 2:
                return [_hist_row_to_dict(r) for r in rows[1:] if any(r)]
        except Exception:
            pass
    # 3. Local JSON (last resort)
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def save_history(records: list):
    # 1. Supabase (primary) — upsert by id
    sb, _ = _get_supabase_client()
    if sb is not None:
        try:
            rows = [_hist_dict_to_sb_row(r) for r in records]
            if rows:
                sb.table("history").upsert(rows, on_conflict="id").execute()
            return
        except Exception:
            pass
    # 2. Google Sheets (fallback)
    ws, _ = _get_history_sheet()
    if ws is not None:
        try:
            all_values = [_HIST_HEADERS] + [_hist_dict_to_row(r) for r in records]
            ws.clear()
            if all_values:
                ws.update("A1", all_values)
            return
        except Exception:
            pass
    # 3. Local JSON (last resort)
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2, default=str)


def upsert_history(all_errors: list, period_from: str, period_to: str, reviewer: str):
    """Add/update history records. Global dedup by (teacher, date, error_type) — no duplicates across periods."""
    import uuid as _uuid
    records = load_history()
    now = datetime.now().isoformat(timespec="seconds")

    # Global index — one record per (teacher, date, error_type) regardless of check period
    existing = {(r["teacher"], r["date"], r["error_type"]): r for r in records}

    result = {}   # will become the new full history

    # Process current errors
    current_keys = set()
    for e in all_errors:
        key = (e["teacher"], e["date"], e["error_type"])
        current_keys.add(key)
        if key in existing:
            rec = existing[key].copy()
            rec.update({"count": e["count"], "error_description": e["error_description"], "updated_at": now})
            if e.get("students"):   # refresh student list on re-check
                rec["students"] = e["students"]
        else:
            rec = {
                "id":                str(_uuid.uuid4())[:8],
                "checked_at":        now,
                "period_from":       period_from,
                "period_to":         period_to,
                "date":              e["date"],
                "teacher":           e["teacher"],
                "student_tag":       "Все",
                "error_type":        e["error_type"],
                "error_description": e["error_description"],
                "count":             e["count"],
                "students":          e.get("students", []),
                "status":            "open",
                "reviewer":          reviewer,
                "reviewer_comment":  "",
                "updated_at":        now,
            }
        result[key] = rec

    # Carry over all records outside current check; auto-close those inside that disappeared
    for key, rec in existing.items():
        if key in result:
            continue   # already handled above
        if period_from <= rec["date"] <= period_to:
            # Was in range but error gone — teacher fixed it
            rec = rec.copy()
            if rec["status"] in ("open", "message_sent", "pass_set"):
                rec.update({"status": "handled", "updated_at": now})
        result[key] = rec

    save_history(list(result.values()))


def update_history_sent(period_from: str, period_to: str, teachers_sent: set, reviewer: str):
    """Mark no_report records as message_sent for teachers within the checked date range."""
    records = load_history()
    now = datetime.now().isoformat(timespec="seconds")
    for rec in records:
        if (period_from <= rec["date"] <= period_to
                and rec["error_type"] == "no_report"
                and rec["teacher"] in teachers_sent
                and rec["status"] == "open"):
            rec.update({"status": "message_sent", "reviewer": reviewer, "updated_at": now})
    save_history(records)


def update_history_passes(period_from: str, period_to: str, reviewer: str):
    """Mark open/sent no_report records as pass_set within the checked date range."""
    records = load_history()
    now = datetime.now().isoformat(timespec="seconds")
    for rec in records:
        if (period_from <= rec["date"] <= period_to
                and rec["error_type"] == "no_report"
                and rec["status"] in ("open", "message_sent")):
            rec.update({"status": "pass_set", "reviewer": reviewer, "updated_at": now})
    save_history(records)


def update_history_record(record_id: str, status: str, comment: str, reviewer: str):
    """Update a single record's status and comment."""
    records = load_history()
    now = datetime.now().isoformat(timespec="seconds")
    for rec in records:
        if rec["id"] == record_id:
            rec.update({"status": status, "reviewer_comment": comment, "reviewer": reviewer, "updated_at": now})
            break
    save_history(records)


# ─────────────────────────────────────────────────────────────────────────────
#  TEACHER INFO — Google Sheet (publicly shared, no auth needed)
# ─────────────────────────────────────────────────────────────────────────────

_TEACHER_SHEET_ID  = "1ipNiGu1cjGfkhWEmsJ0NG4K4icyYyIDFvpwRFZdOklQ"
_TEACHER_SHEET_GID = "1826758403"

@st.cache_data(ttl=3600, show_spinner=False)
def load_teacher_info(creds_json_str: str = "") -> tuple:
    """
    Returns (dict, str) — teacher info dict and status message.
    dict: { "ФИО": {"is_intern": bool, "subject": str, "fired": bool} }
    """
    import io

    def _parse_rows(rows: list) -> dict:
        result = {}
        for row in rows:
            name = str(row.get("ФИО", "")).strip()
            if not name or name.lower() == "nan":
                continue
            _raw_status = str(row.get("Статус набора", "")).strip()
            result[name] = {
                "is_intern": str(row.get("стажер", "")).strip().lower() == "стажер",
                "subject":   str(row.get("Предмет", "")).strip(),
                "status":    _raw_status,
                "fired":     _raw_status.lower() == "уволен",
            }
        return result

    # — Via gspread (service account) ————————————————————————————————————————
    import gspread as _gspread
    from google.oauth2.service_account import Credentials as _Creds

    _SA_SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    def _open_sheet(sa_info: dict):
        creds = _Creds.from_service_account_info(sa_info, scopes=_SA_SCOPES)
        gc    = _gspread.authorize(creds)
        sh    = gc.open_by_key(_TEACHER_SHEET_ID)
        gid   = int(_TEACHER_SHEET_GID)
        for w in sh.worksheets():
            if w.id == gid:
                return w
        return None

    def _ws_to_records(ws):
        """Read worksheet tolerating empty/duplicate column headers.
        Returns (records, found_headers, total_rows)."""
        all_vals = ws.get_all_values()
        if not all_vals:
            return [], [], 0
        # Find the header row: first row that contains "ФИО"
        header_idx = 0
        for i, row in enumerate(all_vals):
            if any("ФИО" in str(cell) for cell in row):
                header_idx = i
                break
        headers = [h.strip() for h in all_vals[header_idx]]
        records = []
        for row in all_vals[header_idx + 1:]:
            row_dict = {}
            for i, val in enumerate(row):
                if i < len(headers) and headers[i]:
                    row_dict[headers[i]] = val
            records.append(row_dict)
        return records, headers, len(all_vals)

    pub_err = ""

    # 1) st.secrets["gcp_service_account"] — native TOML dict (best on cloud)
    try:
        ws = _open_sheet(dict(st.secrets["gcp_service_account"]))
        if ws:
            records, found_headers, total_rows = _ws_to_records(ws)
            data = _parse_rows(records)
            if not data:
                # Return debug info about what was found
                headers_preview = ", ".join(found_headers[:6])
                return {}, (f"⚠️ Загружено (Secrets) 0 преп. "
                            f"(строк: {total_rows}, колонки: {headers_preview})")
            return data, f"✅ Загружено (Secrets): {len(data)} преп."
    except Exception as e:
        pub_err = f"secrets: {e}"

    # 2) creds_json_str — JSON string from config/uploader
    if creds_json_str:
        try:
            sa_info = json.loads(creds_json_str)
            ws = _open_sheet(sa_info)
            if ws:
                records, found_headers, total_rows = _ws_to_records(ws)
                data = _parse_rows(records)
                return data, f"✅ Загружено (JSON): {len(data)} преп."
        except Exception as e:
            pub_err += f" | json: {e}"

    # — Fallback: public CSV ——————————————————————————————————————————————————
    url = (
        f"https://docs.google.com/spreadsheets/d/{_TEACHER_SHEET_ID}"
        f"/export?format=csv&gid={_TEACHER_SHEET_GID}"
    )
    try:
        r = requests.get(url, timeout=15, allow_redirects=True)
        r.raise_for_status()
        df = pd.read_csv(io.StringIO(r.content.decode("utf-8", errors="replace")))
        data = _parse_rows(df.to_dict("records"))
        return data, f"✅ Загружено через публичный URL: {len(data)} преп."
    except Exception as e:
        msg = f"❌ Service Account: {pub_err} | CSV: {e}" if pub_err else f"❌ {e}"
        return {}, msg


def intern_badge(teacher: str, teacher_info: dict) -> str:
    """Return HTML pill badge if teacher is an intern, else empty string."""
    info = teacher_info.get(teacher, {})
    if info.get("is_intern"):
        return '<span class="pill pill-intern">🎓 Стажёр</span>'
    return ""


def load_config() -> dict:
    # 1) Local config.json takes priority (development / self-hosted)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    # 2) Streamlit secrets (Streamlit Community Cloud deployment)
    try:
        return {k: v for k, v in st.secrets.items()}
    except Exception:
        return {}

def save_config(cfg: dict):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

cfg = load_config()


# ─────────────────────────────────────────────────────────────────────────────
#  SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

def _init():
    defaults = {
        "loaded":          False,
        "all_errors":      [],
        "teacher_errors":  {},   # teacher -> {ErrorType -> {"dates": set, "count": int}}
        "messages":        {},   # teacher -> str  (editable)
        "selected":        {},   # teacher -> bool
        "send_results":    {},   # teacher -> {"ok", "reason", "detail"}
        "sending_done":    False,
        "passes_to_set":   [],   # list of {edUnitId, studentClientId, date} — for SetStudentPasses
        "passes_done":     False,
        "passes_results":  {},   # summary after setting passes
        "DATE_FROM":       "",
        "DATE_TO":         "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init()


# ─────────────────────────────────────────────────────────────────────────────
#  HOLLIHOP API
# ─────────────────────────────────────────────────────────────────────────────

def api_call(base_url, api_key, method, params=None, max_retries=3):
    if params is None:
        params = {}
    params["authkey"] = api_key
    params["culture"]  = "ru-RU"
    url = f"{base_url}/{method}"
    for attempt in range(max_retries):
        try:
            r = requests.get(url, params=params, timeout=60)
            if r.status_code == 200:
                return r.json()
            elif r.status_code == 500:
                err = r.json().get("Error", "")
                if "limit" in str(err).lower():
                    time.sleep(30)
                    continue
            return None
        except requests.exceptions.Timeout:
            time.sleep(5)
        except Exception:
            return None
    return None


def api_paginated(base_url, api_key, method, result_key, params=None, take=1000):
    if params is None:
        params = {}
    params["take"] = take
    all_items, skip = [], 0
    while True:
        params["skip"] = skip
        data = api_call(base_url, api_key, method, params.copy())
        if not data:
            break
        items = data.get(result_key, [])
        if not items:
            break
        all_items.extend(items)
        if len(items) < take:
            break
        skip += take
        time.sleep(0.1)
    return all_items


def load_test_results(base_url, api_key, date_from, date_to):
    all_results, skip = [], 0
    while True:
        r = requests.get(f"{base_url}/GetEdUnitTestResults", params={
            "authkey": api_key, "dateFrom": date_from, "dateTo": date_to,
            "take": 1000, "skip": skip,
        }, timeout=60)
        batch = r.json().get("EdUnitTestResults", [])
        if not batch:
            break
        all_results.extend(batch)
        skip += len(batch)
        if len(batch) < 1000:
            break
    return all_results


# ─────────────────────────────────────────────────────────────────────────────
#  MESSAGE TEMPLATES
# ─────────────────────────────────────────────────────────────────────────────

MSG_GREETING = "Добрый вечер!"

MSG_NO_REPORT = (
    "По результатам проверки обнаружено, что за следующие даты "
    "отсутствует отчёт по занятию:\n\n"
    "{dates_list}\n\n"
    "По данным занятиям автоматически выставлен пропуск.\n\n"
    "Если занятие было проведено — необходимо:\n"
    "1. Убрать пропуск в системе\n"
    "2. Заполнить отчёт:\n"
    "   — оценка\n"
    "   — описание урока\n"
    "   — комментарий\n"
    "   — домашнее задание\n\n"
    "⚠️ Без выполнения этих шагов оплата за занятие начислена не будет."
)

MSG_FOOTER = "\n\nСпасибо за внимание к этому вопросу!"


def fmt_date(s: str) -> str:
    try:
        return datetime.strptime(s, "%Y-%m-%d").strftime("%d.%m.%Y")
    except Exception:
        return s


def build_message(teacher: str, ei: dict) -> str:
    sections = []
    if "no_report" in ei:
        dates = sorted(ei["no_report"]["dates"])
        dl = "\n".join(f"• {fmt_date(d)}" for d in dates)
        sections.append(MSG_NO_REPORT.format(dates_list=dl))
    return MSG_GREETING + "\n\n" + "\n\n".join(sections) + MSG_FOOTER


# ─────────────────────────────────────────────────────────────────────────────
#  PACHCA API
# ─────────────────────────────────────────────────────────────────────────────

class PachcaAPI:
    def __init__(self, token: str):
        self.token = token
        self.url   = "https://api.pachca.com/api/shared/v1"
        self.hdrs  = {
            "Authorization":  f"Bearer {token}",
            "Content-Type":   "application/json; charset=utf-8",
        }
        self._cache = {}   # last_name_lower -> user   (unique surnames)
        self._full  = {}   # (last, first)   -> user
        self._all   = []

    def check_token(self) -> bool:
        try:
            r = requests.get(f"{self.url}/profile", headers=self.hdrs, timeout=10)
            return r.status_code != 401
        except Exception:
            return False

    def load_users(self, write_fn=None):
        users, cursor, page = [], None, 1
        while True:
            params = {"limit": 50}
            if cursor:
                params["cursor"] = cursor
            if write_fn:
                write_fn(f"Страница {page}…")
            r = requests.get(f"{self.url}/users", headers=self.hdrs, params=params, timeout=30)
            if r.status_code != 200:
                break
            data  = r.json()
            batch = data.get("data", [])
            users.extend(batch)
            if not batch:
                break
            nxt = data.get("meta", {}).get("paginate", {}).get("next_page")
            if not nxt:
                break
            cursor, page = nxt, page + 1
        self._all = users
        # Build search index
        ln_groups = defaultdict(list)
        for u in users:
            ln = (u.get("last_name")  or "").strip().lower()
            fn = (u.get("first_name") or "").strip().lower()
            if not ln:
                continue
            ln_groups[ln].append(u)
            if fn:
                self._full[(ln, fn)] = u
        for ln, grp in ln_groups.items():
            if len(grp) == 1:
                self._cache[ln] = grp[0]

    def find_user(self, label: str) -> Optional[dict]:
        label = label.strip()
        if not label:
            return None
        parts = label.split()
        ln    = parts[0].lower()
        if len(parts) == 1:
            return self._cache.get(ln)
        fn_part = parts[1].lower()
        exact = self._full.get((ln, fn_part))
        if exact:
            return exact
        initial = fn_part.rstrip(".")
        if len(initial) == 1:
            candidates = [u for (l, f), u in self._full.items()
                          if l == ln and f.startswith(initial)]
            if len(candidates) == 1:
                return candidates[0]
            return None   # ambiguous
        return self._cache.get(ln)

    def send_dm(self, user_id: int, content: str) -> dict:
        payload = {"message": {"entity_type": "user", "entity_id": user_id, "content": content}}
        r = requests.post(f"{self.url}/messages", headers=self.hdrs, json=payload, timeout=30)
        r.raise_for_status()
        return r.json().get("data", {})

    def create_task(self, user_id: int, content: str, due_days: int) -> dict:
        due_at = (datetime.now() + timedelta(days=due_days)).strftime("%Y-%m-%dT23:59:59.000+0300")
        payload = {"task": {"kind": "reminder", "content": content, "due_at": due_at,
                            "priority": 2, "performer_ids": [user_id]}}
        r = requests.post(f"{self.url}/tasks", headers=self.hdrs, json=payload, timeout=30)
        r.raise_for_status()
        return r.json().get("data", {})


# ─────────────────────────────────────────────────────────────────────────────
#  GOOGLE SHEETS
# ─────────────────────────────────────────────────────────────────────────────

def _gs_client(creds_json: str):
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    info  = json.loads(creds_json)
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)


def write_to_sheets(sheet_id, sheet_name, errors, date_from, date_to, reviewer, creds_json):
    try:
        gc = _gs_client(creds_json)
        sh = gc.open_by_key(sheet_id)
        try:
            ws = sh.worksheet(sheet_name)
        except Exception:
            ws = sh.add_worksheet(title=sheet_name, rows=1000, cols=10)
        existing = ws.get_all_values()
        next_row = len(existing) + 1
        if next_row == 1:
            ws.update("A1:G1", [[
                f"{date_from}—{date_to}", "Имя препода", "Тег ученика",
                "Описание ошибки", "Статус ошибки", "Комментарий проверяющего", "Кол-во",
            ]])
            next_row = 2
        rows = [[e["date"], e["teacher"], "Все", e["error_description"], "", "", e["count"]]
                for e in sorted(errors, key=lambda x: x["teacher"])]
        if rows:
            ws.update(f"A{next_row}:G{next_row + len(rows) - 1}", rows)
        return True, f"Записано {len(rows)} строк"
    except Exception as e:
        return False, str(e)


def set_student_passes(base_url, api_key, passes, absence_comment, dry_run=False, write_fn=None, batch_size=100):
    """
    Ставит пропуск (Pass=True) с комментарием для переданных занятий.
    passes: list of {edUnitId, studentClientId, date}
    Возвращает (ok_count, err_count, errors_list)
    """
    if not passes:
        return 0, 0, []

    url        = f"{base_url}/SetStudentPasses"
    ok_count   = 0
    err_count  = 0
    errors_out = []

    # Отправляем батчами по batch_size записей
    for i in range(0, len(passes), batch_size):
        batch = []
        for p in passes[i : i + batch_size]:
            existing = (p.get("existing_desc") or "").strip()
            combined = f"{existing}\n{absence_comment}".strip() if existing else absence_comment
            batch.append({
                "edUnitId":        p["edUnitId"],
                "studentClientId": p["studentClientId"],
                "date":            p["date"],
                "pass":            True,
                "description":     combined,
            })
        if write_fn:
            write_fn(f"Пакет {i // batch_size + 1}: {len(batch)} записей…")

        if dry_run:
            ok_count += len(batch)
            continue

        try:
            r = requests.post(
                url,
                params={"authkey": api_key},
                json=batch,
                timeout=60,
            )
            if r.status_code in (200, 204):
                ok_count += len(batch)
            else:
                err_count += len(batch)
                errors_out.append(f"HTTP {r.status_code}: {r.text[:200]}")
                if write_fn:
                    write_fn(f"   ⚠️ Ошибка: HTTP {r.status_code}")
        except Exception as e:
            err_count += len(batch)
            errors_out.append(str(e))
            if write_fn:
                write_fn(f"   ❌ Исключение: {e}")

        time.sleep(0.3)

    return ok_count, err_count, errors_out


def update_sheet_statuses(sheet_id, sheet_name, sent_teachers, date_from, date_to, reviewer, creds_json):
    try:
        gc = _gs_client(creds_json)
        ws = gc.open_by_key(sheet_id).worksheet(sheet_name)
        all_rows = ws.get_all_values()
        updated  = 0
        for i, row in enumerate(all_rows, start=1):
            if i == 1:
                continue
            if len(row) >= 5:
                teacher  = row[1].strip()
                status   = row[4].strip()
                date_val = row[0].strip()
                if date_from <= date_val <= date_to and not status and teacher in sent_teachers:
                    ws.update_cell(i, 5, f"Обработано ({reviewer})")
                    updated += 1
                    time.sleep(0.3)
        return True, f"Обновлено {updated} строк"
    except Exception as e:
        return False, str(e)


# Load teacher info BEFORE the header (settings popover references it)
_teacher_info, _teacher_info_status = load_teacher_info(creds_json_str=cfg.get("creds_json") or "")

# ─────────────────────────────────────────────────────────────────────────────
#  WELCOME SCREEN
# ─────────────────────────────────────────────────────────────────────────────

def _show_welcome():
    import random as _rng

    hour = datetime.now().hour
    if 0 <= hour < 6:
        greeting, emoji, period = "Доброй ночи", "🌙", "night"
        bg       = "radial-gradient(ellipse at 50% 0%, #2d1b69 0%, #0d0628 50%, #060614 100%)"
        accent   = "#a78bfa"
        txt_sub  = "rgba(200,185,255,0.7)"
        card_bg  = "rgba(255,255,255,0.04)"
        card_bdr = "rgba(167,139,250,0.2)"
    elif 6 <= hour < 12:
        greeting, emoji, period = "Доброе утро", "🌅", "morning"
        bg       = "linear-gradient(160deg, #431407 0%, #9a3412 20%, #ea580c 55%, #fcd34d 100%)"
        accent   = "#fb923c"
        txt_sub  = "rgba(255,237,213,0.75)"
        card_bg  = "rgba(255,255,255,0.10)"
        card_bdr = "rgba(251,191,36,0.3)"
    elif 12 <= hour < 18:
        greeting, emoji, period = "Добрый день", "☀️", "day"
        bg       = "linear-gradient(160deg, #0c4a6e 0%, #0369a1 35%, #0ea5e9 70%, #7dd3fc 100%)"
        accent   = "#38bdf8"
        txt_sub  = "rgba(224,242,254,0.75)"
        card_bg  = "rgba(255,255,255,0.12)"
        card_bdr = "rgba(56,189,248,0.35)"
    else:
        greeting, emoji, period = "Добрый вечер", "🌆", "evening"
        bg       = "linear-gradient(160deg, #1e1b4b 0%, #4c1d95 35%, #7e22ce 60%, #9d174d 100%)"
        accent   = "#c084fc"
        txt_sub  = "rgba(233,213,255,0.75)"
        card_bg  = "rgba(255,255,255,0.06)"
        card_bdr = "rgba(192,132,252,0.25)"

    # Generate decorative particles with fixed seed
    _rng.seed(77)
    particles_css = ""
    particles_html = ""
    for i in range(18):
        x   = _rng.randint(2, 98)
        y   = _rng.randint(2, 98)
        sz  = _rng.randint(3, 14)
        dur = round(_rng.uniform(2.5, 6.5), 1)
        dly = round(_rng.uniform(0, 3), 1)
        op  = round(_rng.uniform(0.08, 0.22), 2)
        particles_css += f"""
        .wb-p{i} {{
            position: fixed; left:{x}vw; top:{y}vh;
            width:{sz}px; height:{sz}px; border-radius:50%;
            background: rgba(255,255,255,{op});
            animation: wb-float {dur}s ease-in-out {dly}s infinite alternate;
            pointer-events: none;
        }}"""
        particles_html += f'<div class="wb-p{i}"></div>'

    # Night: generate a tiny star-field using box-shadow trick
    night_stars = ""
    if period == "night":
        _rng.seed(42)
        _s = ", ".join(
            f"{_rng.randint(0,1900)}px {_rng.randint(0,1000)}px "
            f"rgba(255,255,255,{_rng.choice([0.3,0.5,0.7,0.9,1.0])})"
            for _ in range(300)
        )
        night_stars = f"""
        .wb-stars {{ position:fixed; top:0; left:0; width:1px; height:1px;
                     background:transparent; box-shadow:{_s};
                     border-radius:50%; animation:wb-twinkle 5s ease-in-out infinite alternate;
                     pointer-events:none; z-index:0; }}
        @keyframes wb-twinkle {{ 0%{{opacity:0.4}} 100%{{opacity:1}} }}
        """

    st.markdown(f"""
    <style>
    /* ── Welcome: full-page takeover ───────────────────────────────────────── */
    html, body, .stApp {{ background: {bg} !important; }}
    .stApp > header, [data-testid="stToolbar"],
    [data-testid="collapsedControl"], [data-testid="stSidebar"] {{ display:none!important; }}

    /* Center the entire Streamlit main column */
    section[data-testid="stMain"] {{
        padding: 0 !important;
    }}
    .block-container, [data-testid="stMainBlockContainer"] {{
        padding: 0 !important; margin: 0 auto !important;
        max-width: 100% !important; min-height: 100vh !important;
        display: flex !important; flex-direction: column !important;
        align-items: center !important; justify-content: center !important;
        gap: 20px !important;
    }}
    /* Remove stray element-container padding */
    .element-container {{ padding: 0 !important; margin: 0 !important; }}
    [data-testid="stHorizontalBlock"] {{
        padding: 0 !important; margin: 0 !important;
    }}

    {night_stars}
    {particles_css}

    @keyframes wb-float {{
        0%   {{ transform: translateY(0px)  scale(1);    }}
        100% {{ transform: translateY(-28px) scale(1.12); }}
    }}

    /* glow orb */
    .wb-orb {{
        position: fixed; top:50%; left:50%;
        transform: translate(-50%,-50%);
        width: 700px; height: 700px; border-radius: 50%;
        background: radial-gradient(circle, {accent}28 0%, transparent 65%);
        animation: wb-orb-pulse 7s ease-in-out infinite;
        pointer-events: none; z-index: 0;
    }}
    @keyframes wb-orb-pulse {{
        0%,100% {{ transform:translate(-50%,-50%) scale(1);   opacity:.7; }}
        50%      {{ transform:translate(-50%,-50%) scale(1.25); opacity:1; }}
    }}

    /* card — flows naturally in centered column */
    .wb-card {{
        position: relative; z-index: 10;
        text-align: center;
        padding: 3rem 3.5rem 2.5rem;
        background: {card_bg};
        backdrop-filter: blur(48px); -webkit-backdrop-filter: blur(48px);
        border: 1px solid {card_bdr};
        border-radius: 36px;
        box-shadow: 0 40px 100px rgba(0,0,0,0.45),
                    inset 0 1px 0 rgba(255,255,255,0.15);
        animation: wb-card-in 0.9s cubic-bezier(0.34,1.56,0.64,1) both;
        max-width: 480px; width: 86vw;
    }}
    @keyframes wb-card-in {{
        0%   {{ transform: translateY(50px) scale(.93); opacity:0; }}
        100% {{ transform: translateY(0)    scale(1);   opacity:1; }}
    }}

    .wb-emoji   {{ font-size:3.6rem; display:block; margin-bottom:.4rem;
                   animation: wb-spin-in 1s ease .15s both; }}
    @keyframes wb-spin-in {{
        0%   {{ transform:scale(0) rotate(-200deg); opacity:0; }}
        100% {{ transform:scale(1) rotate(0deg);    opacity:1; }}
    }}

    .wb-brand   {{ font-size:3rem; font-weight:900; color:#fff;
                   letter-spacing:-1.5px; margin:0 0 .3rem;
                   text-shadow: 0 4px 24px rgba(0,0,0,.35);
                   animation: wb-fade-up .7s ease .3s both; }}

    .wb-greet   {{ font-size:1.15rem; color:{txt_sub}; font-weight:300;
                   letter-spacing:.4px; margin:0;
                   animation: wb-fade-up .7s ease .45s both; }}

    @keyframes wb-fade-up {{
        0%   {{ transform:translateY(18px); opacity:0; }}
        100% {{ transform:translateY(0);    opacity:1; }}
    }}

    /* label between card and buttons */
    .wb-label-outer {{
        font-size: .76rem; color: rgba(255,255,255,.5);
        text-transform: uppercase; letter-spacing: 2.5px; font-weight: 600;
        text-align: center;
        animation: wb-fade-up .7s ease .55s both;
        z-index: 10; position: relative;
    }}

    /* buttons */
    [data-testid="stHorizontalBlock"] .stButton button {{
        padding: 1.1rem 1.6rem !important;
        font-size: 1.05rem !important; font-weight: 700 !important;
        border-radius: 16px !important;
        border: 1px solid rgba(255,255,255,.22) !important;
        background: rgba(255,255,255,.09) !important;
        color: #fff !important;
        backdrop-filter: blur(12px);
        transition: all .3s cubic-bezier(.34,1.56,.64,1) !important;
        box-shadow: 0 8px 28px rgba(0,0,0,.25) !important;
        letter-spacing: .3px !important;
        width: 100% !important;
        animation: wb-fade-up .7s ease .65s both;
    }}
    [data-testid="stHorizontalBlock"] .stButton button:hover {{
        background: rgba(255,255,255,.2)  !important;
        border-color: rgba(255,255,255,.55) !important;
        transform: translateY(-5px) scale(1.04) !important;
        box-shadow: 0 22px 56px rgba(0,0,0,.35) !important;
    }}
    [data-testid="stHorizontalBlock"] .stButton button:active {{
        transform: translateY(0) scale(.97) !important;
    }}
    [data-testid="stHorizontalBlock"] {{
        width: 460px !important; max-width: 86vw !important;
        z-index: 10 !important;
    }}
    </style>

    <div class="wb-orb"></div>
    {'<div class="wb-stars"></div>' if period == "night" else ""}
    {particles_html}

    <div class="wb-card">
        <span class="wb-emoji">{emoji}</span>
        <h1 class="wb-brand">Биллибоба</h1>
        <p class="wb-greet">{greeting}&nbsp; 👋</p>
    </div>
    <div class="wb-label-outer">Под чьим именем хотите войти?</div>
    """, unsafe_allow_html=True)

    _u1, _u2 = st.columns(2, gap="small")
    if _u1.button("🌸  Дарья", key="wb_darya", use_container_width=True):
        st.session_state["billyboba_user"] = "Дарья"
        st.rerun()
    if _u2.button("⚡  Артём", key="wb_artem", use_container_width=True):
        st.session_state["billyboba_user"] = "Артём"
        st.rerun()

    st.stop()


# ── Show welcome if not logged in ────────────────────────────────────────────
if not st.session_state.get("billyboba_user"):
    _show_welcome()

# ─────────────────────────────────────────────────────────────────────────────
#  HEADER — title + dates + settings popover
# ─────────────────────────────────────────────────────────────────────────────

# Sensitive credentials — read from Secrets only (not editable in UI)
subdomain    = _read_secret("subdomain")    or cfg.get("subdomain", "matrix")
api_key      = _read_secret("api_key")      or cfg.get("api_key", "")
pachca_token = _read_secret("pachca_token") or cfg.get("pachca_token", "")

_title_col, _dates_col, _gear_col = st.columns([4, 2.5, 0.7])

with _title_col:
    st.markdown("## 📋 Проверка отчётов HolliHop")

with _dates_col:
    _dc1, _dc2 = st.columns(2)
    with _dc1:
        date_from_val = st.date_input("С", value=date.today() - timedelta(days=7))
    with _dc2:
        date_to_val = st.date_input("По", value=date.today())

with _gear_col:
    st.markdown('<div style="height:1.75rem"></div>', unsafe_allow_html=True)
    with st.popover("⚙️", use_container_width=True):
        st.markdown("### ⚙️ Настройки")

        st.markdown("**🏫 HolliHop**")
        st.caption(f"Субдомен: `{subdomain}.t8s.ru`  •  API ключ: {'✅' if api_key else '⚠️ не задан'}")
        st.caption("Задаются в Streamlit Secrets: `subdomain`, `api_key`")

        st.markdown("**📨 Пачка**")
        st.caption(f"Токен Пачки: {'✅ задан' if pachca_token else '⚠️ не задан'}")
        st.caption("Задаётся в Streamlit Secrets: `pachca_token`")
        create_reminder = st.checkbox("Создавать напоминания", value=cfg.get("create_reminder", False), key="cfg_reminder")
        reminder_days   = st.number_input("Дней до дедлайна", min_value=1, max_value=14,
                                          value=int(cfg.get("reminder_days", 1)), key="cfg_rdays")

        st.markdown("**👤 Проверяющий**")
        _default_reviewer = st.session_state.get("billyboba_user") or cfg.get("reviewer_name", "Артём")
        reviewer_name = st.text_input("Имя", value=_default_reviewer, key="cfg_reviewer")

        with st.expander("📊 Google Таблица (необязательно)"):
            sheet_id   = st.text_input("ID таблицы",  value=cfg.get("sheet_id", ""),    key="cfg_sheet_id")
            sheet_name = st.text_input("Имя листа",   value=cfg.get("sheet_name", "Лист1"), key="cfg_sheet_name")
            creds_file = st.file_uploader("Service Account JSON", type=["json"], key="cfg_creds")
            creds_json = creds_file.read().decode() if creds_file else cfg.get("creds_json", None)

        st.divider()
        st.markdown("**🗄️ Supabase (основное хранилище)**")
        _sb, _sb_err = _get_supabase_client()
        if _sb is not None:
            st.caption("✅ Supabase подключён")
            # Export JSON download
            _all_hist_export = load_history()
            if _all_hist_export:
                _export_json = json.dumps(_all_hist_export, ensure_ascii=False, indent=2, default=str)
                st.download_button(
                    "⬇️ Скачать всю историю (JSON)",
                    data=_export_json.encode("utf-8"),
                    file_name=f"history_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                    mime="application/json",
                    use_container_width=True,
                    key="export_json_btn",
                )
            # Sync to Sheets button
            if st.button("🔄 Синхронизировать в Sheets", key="sync_to_sheets_btn",
                         use_container_width=True):
                with st.spinner("Синхронизирую…"):
                    _ok, _msg = mirror_to_sheets(load_history())
                (st.success if _ok else st.error)(f"{'✅' if _ok else '❌'} {_msg}")
        else:
            st.caption(f"⚠️ Supabase не подключён: `{_sb_err}`")
            with st.expander("Как подключить Supabase"):
                st.markdown("""
1. Зайди на [supabase.com](https://supabase.com) → создай проект (бесплатно)
2. В проекте: **SQL Editor** → выполни:
```sql
create table history (
  id text primary key,
  checked_at text, period_from text, period_to text,
  date text, teacher text, student_tag text default 'Все',
  error_type text, error_description text,
  count integer default 0, students text default '[]',
  status text default 'open', reviewer text,
  reviewer_comment text, updated_at text
);
alter table history enable row level security;
create policy "allow_all" on history for all using (true) with check (true);
```
3. **Settings → API** → скопируй `URL` и `anon public key`
4. Добавь в Streamlit Secrets:
```toml
supabase_url = "https://xxxx.supabase.co"
supabase_key = "eyJ..."
```
5. Нажми **Migrate** ниже для переноса данных из Sheets
""")
            # One-time migration button
            if st.button("📥 Migrate: перенести данные в Supabase",
                         key="migrate_to_sb", use_container_width=True):
                st.info("Настрой Supabase в Secrets и перезапусти приложение, затем нажми снова.")

        st.divider()
        st.markdown("**📗 Зеркало (Google Sheets)**")
        _hist_ws, _hist_sid = _get_history_sheet()
        if _hist_sid:
            st.caption(f"✅ Подключено: `{_hist_sid[:20]}…`")
            _hist_url = f"https://docs.google.com/spreadsheets/d/{_hist_sid}"
            st.caption(f"[Открыть таблицу]({_hist_url})")
        else:
            st.caption("⚠️ Sheets не настроен (резервная копия недоступна)")
            if st.button("➕ Создать таблицу-зеркало", key="create_hist_sheet",
                         use_container_width=True):
                with st.spinner("Создаю таблицу…"):
                    try:
                        _new_sid, _new_url = create_history_sheet()
                        st.success("✅ Таблица создана!")
                        st.code(f'history_sheet_id = "{_new_sid}"', language="toml")
                        st.caption(f"[Открыть таблицу]({_new_url})")
                    except Exception as _ce:
                        st.error(f"Ошибка: {_ce}")

        st.divider()
        st.markdown("**📋 Данные о преподавателях**")
        st.caption(_teacher_info_status if _teacher_info_status else "⚠️ Не загружено")
        if _teacher_info:
            _interns = [n for n, v in _teacher_info.items() if v.get("is_intern")]
            st.caption(f"Стажёров: {len(_interns)} из {len(_teacher_info)}")
            with st.expander("Примеры имён из таблицы"):
                for n in list(_teacher_info.keys())[:8]:
                    st.caption(n)
        if st.button("🔄 Обновить сейчас", key="refresh_teacher_info", use_container_width=True):
            load_teacher_info.clear()
            st.rerun()

        st.divider()
        _is_cloud = not os.path.exists(CONFIG_FILE)
        if _is_cloud:
            st.caption("☁️ Настройки хранятся в Streamlit Secrets")
        else:
            if st.button("💾 Сохранить", use_container_width=True, key="cfg_save"):
                save_config({
                    "create_reminder": create_reminder,
                    "reminder_days": reminder_days, "sheet_id": sheet_id,
                    "sheet_name": sheet_name, "reviewer_name": reviewer_name,
                    "creds_json": creds_json,
                })
                st.success("✅ Сохранено!")

DATE_FROM = date_from_val.strftime("%Y-%m-%d")
DATE_TO   = date_to_val.strftime("%Y-%m-%d")
BASE_URL  = f"https://{subdomain}.t8s.ru/Api/V2"

tab4, tab5, tab1, tab2, tab3, tab6 = st.tabs(["📚 История", "📊 Статистика", "📥 Загрузить данные", "✉️ Сообщения", "📤 Отправить", "📢 Рассылка"])


# Load history once per render — reused across Tab 4 and Tab 5
_all_history: list = load_history()

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 1 — LOAD & ANALYSE
# ─────────────────────────────────────────────────────────────────────────────

with tab1:
    col_btn, col_hint = st.columns([1, 3])
    with col_btn:
        load_btn = st.button("🔄 Загрузить данные", type="primary", use_container_width=True)
    with col_hint:
        st.markdown(
            f'<div class="info-bar">Период:&nbsp;<b>{fmt_date(DATE_FROM)} — {fmt_date(DATE_TO)}</b>'
            f'&nbsp;&nbsp;·&nbsp;&nbsp;HolliHop:&nbsp;<b>{subdomain}.t8s.ru</b></div>',
            unsafe_allow_html=True,
        )

    if load_btn:
        if not api_key:
            st.error("Укажи API ключ HolliHop в настройках слева!")
        elif DATE_FROM > DATE_TO:
            st.error("Дата «С» должна быть раньше даты «По»!")
        else:
            with st.status("Загружаю данные…", expanded=True) as load_status:

                # — Teachers —
                st.write("👨‍🏫 Преподаватели…")
                teachers_raw = api_paginated(BASE_URL, api_key, "GetTeachers", "Teachers")
                st.write(f"   → {len(teachers_raw)} преподавателей")

                # — Ed units —
                st.write("📚 Учебные единицы + расписание…")
                ed_units = api_paginated(BASE_URL, api_key, "GetEdUnits", "EdUnits", params={
                    "types": "Group,MiniGroup,Individual",
                    "dateFrom": DATE_FROM, "dateTo": DATE_TO, "queryDays": "true",
                })
                eu_map = {}
                for eu in ed_units:
                    eu_id, t_names = eu["Id"], []
                    for si in eu.get("ScheduleItems", []):
                        for tn in si.get("Teachers", []):
                            if tn not in t_names:
                                t_names.append(tn)
                    eu_map[eu_id] = {"name": eu.get("Name", ""), "teachers": t_names}
                st.write(f"   → {len(ed_units)} учебных единиц")

                # — Attendance —
                st.write("📋 Посещаемость учеников…")
                eu_students = api_paginated(BASE_URL, api_key, "GetEdUnitStudents", "EdUnitStudents", params={
                    "dateFrom": DATE_FROM, "dateTo": DATE_TO, "queryDays": "true",
                })
                st.write(f"   → {len(eu_students)} записей посещаемости")

                # — Test results —
                st.write("📝 Результаты тестов (отчёты)…")
                all_test_results = load_test_results(BASE_URL, api_key, DATE_FROM, DATE_TO)
                filled_reports = {(r["EdUnitId"], r["Date"], r["StudentClientId"]) for r in all_test_results}
                st.write(f"   → {len(all_test_results)} отчётов заполнено")

                # — Analysis —
                st.write("🔍 Ищу ошибки…")
                no_report  = defaultdict(lambda: defaultdict(int))
                no_comment = defaultdict(lambda: defaultdict(int))
                # Track student names per (teacher, date, error_type)
                _nr_students  = defaultdict(lambda: defaultdict(list))  # no_report
                _nc_students  = defaultdict(lambda: defaultdict(list))  # no_comment
                passes_to_set = []

                for eus in eu_students:
                    eu_id      = eus.get("EdUnitId")
                    student_id = eus.get("StudentClientId")
                    # Try common name fields in HolliHop API response
                    student_name = (eus.get("Name") or eus.get("StudentName")
                                    or eus.get("ClientName") or str(student_id))
                    teachers   = eu_map.get(eu_id, {}).get("teachers", ["Неизвестный"])
                    for day in eus.get("Days", []):
                        d = day.get("Date", "")
                        if d < DATE_FROM or d > DATE_TO:
                            continue
                        is_pass = day.get("Pass", False)
                        desc    = (day.get("Description") or "").strip()
                        if not is_pass and (eu_id, d, student_id) not in filled_reports:
                            for t in teachers:
                                no_report[t][d] += 1
                                _nr_students[t][d].append(student_name)
                            passes_to_set.append({
                                "edUnitId":        eu_id,
                                "studentClientId": student_id,
                                "date":            d,
                                "existing_desc":   desc,
                            })
                        if is_pass and not desc:
                            for t in teachers:
                                no_comment[t][d] += 1
                                _nc_students[t][d].append(student_name)

                # Build flat errors list + teacher_errors dict
                all_errors, teacher_errors = [], {}

                for teacher, date_counts in no_report.items():
                    teacher_errors.setdefault(teacher, {})["no_report"] = {
                        "dates": set(date_counts.keys()),
                        "count": sum(date_counts.values()),
                    }
                    for d in sorted(date_counts):
                        all_errors.append({
                            "date": d, "teacher": teacher,
                            "error_type": "no_report",
                            "error_description": "Нет отчёта",
                            "count": date_counts[d],
                            "students": sorted(set(_nr_students[teacher][d])),
                        })

                for teacher, date_counts in no_comment.items():
                    teacher_errors.setdefault(teacher, {})["no_abs_comment"] = {
                        "dates": set(date_counts.keys()),
                        "count": sum(date_counts.values()),
                    }
                    for d in sorted(date_counts):
                        all_errors.append({
                            "date": d, "teacher": teacher,
                            "error_type": "no_abs_comment",
                            "error_description": "Нет комментария к пропуску",
                            "count": date_counts[d],
                            "students": sorted(set(_nc_students[teacher][d])),
                        })

                # Pre-generate messages
                messages, selected = {}, {}
                for teacher, ei in teacher_errors.items():
                    if "no_report" in ei:
                        messages[teacher] = build_message(teacher, ei)
                        selected[teacher] = True

                # Clear stale checkbox widget keys from a previous load
                for _k in list(st.session_state.keys()):
                    if _k.startswith("chk_"):
                        del st.session_state[_k]

                st.session_state.update({
                    "loaded":          True,
                    "all_errors":      all_errors,
                    "teacher_errors":  teacher_errors,
                    "messages":        messages,
                    "selected":        selected,
                    "send_results":    {},
                    "sending_done":    False,
                    "passes_to_set":   passes_to_set,
                    "passes_done":     False,
                    "passes_results":  {},
                    "DATE_FROM":       DATE_FROM,
                    "DATE_TO":         DATE_TO,
                })

                # Save to history
                if all_errors:
                    upsert_history(all_errors, DATE_FROM, DATE_TO, reviewer_name)

                n_nr = sum(1 for e in all_errors if e["error_type"] == "no_report")
                n_nc = sum(1 for e in all_errors if e["error_type"] == "no_abs_comment")
                load_status.update(
                    label=f"✅ Готово — найдено {len(all_errors)} ошибок "
                          f"({n_nr} без отчёта · {n_nc} без комментария)",
                    state="complete", expanded=False,
                )

    # ── Results ────────────────────────────────────────────────────────────
    if st.session_state.loaded:
        errors = st.session_state.all_errors

        if not errors:
            st.success("🎉 Ошибок за выбранный период не найдено!")
        else:
            st.divider()

            # Metrics
            teachers_nr = {e["teacher"] for e in errors if e["error_type"] == "no_report"}
            teachers_nc = {e["teacher"] for e in errors if e["error_type"] == "no_abs_comment"}
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Всего ошибок",         len(errors))
            c2.metric("Нет отчёта",           f"{sum(e['count'] for e in errors if e['error_type']=='no_report')} занятий")
            c3.metric("Преподавателей (отчёт)", len(teachers_nr))
            c4.metric("Нет комментария",       f"{sum(e['count'] for e in errors if e['error_type']=='no_abs_comment')} пропусков")

            st.subheader("Найденные ошибки")

            # Filter
            filter_col, _ = st.columns([1, 2])
            with filter_col:
                f_type = st.selectbox(
                    "Фильтр",
                    ["Все ошибки", "Только: нет отчёта", "Только: нет комментария к пропуску"],
                )

            df = pd.DataFrame([{
                "Дата":              e["date"],
                "Преподаватель":     e["teacher"],
                "Стажёр":           _teacher_info.get(e["teacher"], {}).get("is_intern", False),
                "Ошибка":            e["error_description"],
                "Занятий/пропусков": e["count"],
            } for e in sorted(errors, key=lambda x: (x["teacher"], x["date"]))])

            if f_type == "Только: нет отчёта":
                df = df[df["Ошибка"] == "Нет отчёта"]
            elif f_type == "Только: нет комментария к пропуску":
                df = df[df["Ошибка"] == "Нет комментария к пропуску"]

            st.dataframe(df, use_container_width=True, hide_index=True)

            # Export
            col_csv, col_gs = st.columns([1, 2])
            with col_csv:
                csv = df.to_csv(index=False).encode("utf-8-sig")
                st.download_button("⬇️ Скачать CSV", csv, "errors.csv", "text/csv", use_container_width=True)
            with col_gs:
                if sheet_id and creds_json:
                    if st.button("📊 Записать в Google Таблицу", use_container_width=True):
                        with st.spinner("Записываю…"):
                            ok, msg = write_to_sheets(
                                sheet_id, sheet_name, st.session_state.all_errors,
                                st.session_state.DATE_FROM, st.session_state.DATE_TO,
                                reviewer_name, creds_json,
                            )
                        (st.success if ok else st.error)(f"{'✅' if ok else '❌'} {msg}")
                else:
                    st.caption("Загрузи Service Account JSON в настройках для записи в Google Таблицу")

            st.info("Перейди на вкладку **✉️ Сообщения** для предпросмотра и редактирования")


# ─────────────────────────────────────────────────────────────────────────────
#  TAB 2 — MESSAGE PREVIEW & EDIT
# ─────────────────────────────────────────────────────────────────────────────

with tab2:
    if not st.session_state.loaded:
        st.info("Сначала загрузи данные на вкладке **📥 Загрузить данные**")
    elif not st.session_state.teacher_errors:
        st.success("🎉 Нет ошибок — сообщения отправлять не нужно!")
    else:
        teacher_errors = st.session_state.teacher_errors
        messages       = st.session_state.messages
        selected       = st.session_state.selected

        # Teachers that need a message (have NO_REPORT)
        mailable = [t for t in messages]
        # Teachers with only missing comments (no message)
        comment_only = [
            t for t, ei in teacher_errors.items()
            if "no_abs_comment" in ei and "no_report" not in ei
        ]

        st.subheader("✉️ Предпросмотр и редактирование сообщений")
        st.caption(
            f"Ниже — преподаватели, которым **будут отправлены сообщения** (только за отсутствующие отчёты). "
            f"Сними флажок, чтобы исключить кого-то из рассылки."
        )

        # Select/deselect buttons
        b1, b2, _, stat = st.columns([1, 1, 2, 2])
        if b1.button("✅ Выбрать всех", use_container_width=True):
            for t in mailable:
                st.session_state.selected[t] = True
                st.session_state[f"chk_{t}"] = True   # sync widget key
            st.rerun()
        if b2.button("⬜ Снять всех", use_container_width=True):
            for t in mailable:
                st.session_state.selected[t] = False
                st.session_state[f"chk_{t}"] = False  # sync widget key
            st.rerun()

        n_sel = sum(1 for v in selected.values() if v)
        stat.metric("Выбрано к отправке", f"{n_sel} из {len(mailable)}")

        st.divider()

        # ── One card per teacher ──────────────────────────────────────────
        for teacher in sorted(mailable):
            ei        = teacher_errors[teacher]
            nr_info   = ei.get("no_report", {})
            dates_set = nr_info.get("dates", set())
            count     = nr_info.get("count", 0)
            dates_str = "  ·  ".join(fmt_date(d) for d in sorted(dates_set))

            # Init widget key once (first render only)
            chk_key = f"chk_{teacher}"
            if chk_key not in st.session_state:
                st.session_state[chk_key] = selected.get(teacher, True)

            with st.container(border=True):
                col_chk, col_info = st.columns([0.04, 0.96])
                with col_chk:
                    # Read is_selected directly from the checkbox return value
                    is_selected = st.checkbox(
                        "Включить",
                        key=chk_key,
                        label_visibility="collapsed",
                    )
                    # Keep selected dict in sync (no rerun needed)
                    st.session_state.selected[teacher] = is_selected

                with col_info:
                    # Header row
                    hcol1, hcol2 = st.columns([3, 1])
                    with hcol1:
                        icon = "🟢" if is_selected else "⚪"
                        _ibadge = intern_badge(teacher, _teacher_info)
                        st.markdown(
                            f"**{icon} {teacher}**&nbsp;&nbsp;{_ibadge}",
                            unsafe_allow_html=True,
                        )
                    with hcol2:
                        if is_selected:
                            st.markdown('<span class="pill pill-sent">✉️ Отправится</span>', unsafe_allow_html=True)
                        else:
                            st.markdown('<span class="pill pill-skipped">⛔ Пропущен</span>', unsafe_allow_html=True)

                    st.caption(f"📭 {count} занятий без отчёта · {dates_str}")

                    if is_selected:
                        with st.expander("✏️ Посмотреть / отредактировать сообщение"):
                            current_msg = st.session_state.messages.get(teacher, "")
                            new_msg = st.text_area(
                                "Текст сообщения",
                                value=current_msg,
                                height=210,
                                key=f"msg_{teacher}",
                                label_visibility="collapsed",
                            )
                            if new_msg != current_msg:
                                st.session_state.messages[teacher] = new_msg

                            btn_col, _ = st.columns([1, 3])
                            if btn_col.button("↺ Сбросить к шаблону", key=f"reset_{teacher}"):
                                st.session_state.messages[teacher] = build_message(teacher, ei)
                                st.rerun()

        # ── Comment-only teachers ─────────────────────────────────────────
        if comment_only:
            st.divider()
            with st.expander(
                f"⚠️ Нет комментария к пропуску — только в таблицу, без рассылки "
                f"({len(comment_only)} преп.)"
            ):
                for t in sorted(comment_only):
                    nc = teacher_errors[t]["no_abs_comment"]
                    dates_str = "  ·  ".join(fmt_date(d) for d in sorted(nc["dates"]))
                    st.markdown(f"• **{t}** — {nc['count']} пропусков · {dates_str}")

        st.divider()
        st.info("Когда всё проверено и отредактировано — переходи на вкладку **📤 Отправить**")


# ─────────────────────────────────────────────────────────────────────────────
#  TAB 3 — SEND
# ─────────────────────────────────────────────────────────────────────────────

with tab3:
    if not st.session_state.loaded:
        st.info("Сначала загрузи данные на вкладке **📥 Загрузить данные**")
    elif not st.session_state.messages:
        st.success("🎉 Нет сообщений для отправки!")
    else:
        selected          = st.session_state.selected
        messages          = st.session_state.messages
        selected_teachers = sorted(t for t, v in selected.items() if v)

        if not selected_teachers:
            st.warning(
                "Не выбрано ни одного преподавателя. "
                "Вернись на вкладку **✉️ Сообщения** и включи нужных получателей."
            )
        elif not pachca_token:
            st.error("Укажи токен Пачки в настройках слева!")
        else:
            st.subheader("📤 Готово к отправке")

            # Preview table
            df_send = pd.DataFrame([{
                "Преподаватель": t,
                "Занятий без отчёта": st.session_state.teacher_errors[t]
                    .get("no_report", {}).get("count", 0),
                "Даты": "  ·  ".join(
                    fmt_date(d) for d in sorted(
                        st.session_state.teacher_errors[t]
                        .get("no_report", {}).get("dates", [])
                    )
                ),
            } for t in selected_teachers])
            st.dataframe(df_send, use_container_width=True, hide_index=True)
            st.caption(f"Будет отправлено **{len(selected_teachers)}** сообщений")

            st.divider()

            # ── Auto-pass block ───────────────────────────────────────────
            passes_to_set = st.session_state.get("passes_to_set", [])
            with st.expander(
                f"🚫 Автоматически поставить пропуск в HolliHop "
                f"({len(passes_to_set)} занятий без отчёта)",
                expanded=True,
            ):
                if not passes_to_set:
                    st.info("Нет занятий без отчёта — выставлять нечего.")
                else:
                    st.caption(
                        f"Для **{len(passes_to_set)}** записей в HolliHop будет выставлен пропуск "
                        f"и добавлен комментарий."
                    )
                    absence_comment = st.text_input(
                        "Текст комментария к пропуску",
                        value="Автоматический пропуск",
                        key="absence_comment_input",
                        disabled=st.session_state.passes_done,
                    )
                    dry_run_passes = st.toggle(
                        "🔸 Тестовый режим (не менять в HolliHop)",
                        key="dry_run_passes",
                        value=False,
                        disabled=st.session_state.passes_done,
                    )
                    passes_btn = st.button(
                        "🔸 Проверить (тест)" if dry_run_passes else "🚫 Выставить пропуски в HolliHop",
                        type="secondary",
                        disabled=st.session_state.passes_done,
                        key="passes_btn",
                    )

                    if passes_btn:
                        with st.status("Выставляю пропуски…", expanded=True) as passes_status:
                            ok_n, err_n, errs = set_student_passes(
                                BASE_URL, api_key,
                                passes_to_set,
                                absence_comment,
                                dry_run=dry_run_passes,
                                write_fn=st.write,
                            )
                            st.session_state.passes_results = {
                                "ok": ok_n, "err": err_n,
                                "errors": errs, "dry_run": dry_run_passes,
                            }
                            if not dry_run_passes:
                                st.session_state.passes_done = True
                                # Update history: mark records as pass_set
                                update_history_passes(
                                    st.session_state.DATE_FROM, st.session_state.DATE_TO,
                                    reviewer_name,
                                )
                            label = (
                                f"{'Тест' if dry_run_passes else 'Готово'}: "
                                f"{ok_n} ✅  {err_n} ❌"
                            )
                            passes_status.update(
                                label=label,
                                state="complete" if err_n == 0 else "error",
                                expanded=False,
                            )

                    # Results
                    pr = st.session_state.get("passes_results", {})
                    if pr:
                        if pr.get("dry_run"):
                            st.info(f"🔸 Тест: будет обработано {pr['ok']} занятий")
                        elif pr.get("err", 0) == 0:
                            st.success(f"✅ Пропуски выставлены: {pr['ok']} занятий")
                        else:
                            st.warning(f"⚠️ Выставлено: {pr['ok']}  Ошибок: {pr['err']}")
                            for e in pr.get("errors", []):
                                st.caption(f"  {e}")

            st.divider()

            # Dry run toggle
            dry_run = st.toggle(
                "🔸 Тестовый режим (проверить без реальной отправки)",
                value=False,
                disabled=st.session_state.sending_done,
            )

            if dry_run:
                st.info("В тестовом режиме сообщения **не отправляются** — только проверяется, что все преподаватели найдены в Пачке.")

            # Send button
            send_label = "🔸 Проверить (тест)" if dry_run else "🚀 Отправить сообщения"
            send_btn   = st.button(
                send_label,
                type="primary",
                disabled=st.session_state.sending_done,
                use_container_width=False,
            )

            if send_btn:
                pachca = PachcaAPI(pachca_token)

                with st.status("Подготовка…", expanded=True) as send_status:

                    # Verify token
                    st.write("🔑 Проверка токена Пачки…")
                    if not pachca.check_token():
                        st.error("❌ Токен Пачки недействителен! Проверь настройки.")
                        send_status.update(label="Ошибка токена", state="error")
                        st.stop()
                    st.write("   ✓ Токен OK")

                    # Load Pachca users
                    st.write("📇 Загружаю пользователей Пачки…")
                    pachca.load_users(write_fn=lambda m: st.write(f"   {m}"))
                    st.write(f"   ✓ {len(pachca._all)} пользователей загружено")

                    st.divider()

                    # Send loop
                    send_results = {}
                    for teacher in selected_teachers:
                        user = pachca.find_user(teacher)
                        msg  = messages.get(teacher, "")

                        if not user:
                            st.write(f"⚠️ **{teacher}** — не найден в Пачке")
                            send_results[teacher] = {"ok": False, "reason": "not_found", "detail": "Не найден в Пачке"}
                            continue

                        uid       = user["id"]
                        full_name = f"{user.get('first_name','')} {user.get('last_name','')}".strip()

                        if dry_run:
                            st.write(f"🔸 **{teacher}** → {full_name} (ID {uid}): не отправлено (тест)")
                            send_results[teacher] = {"ok": True, "reason": "dry_run", "detail": "Тестовый режим"}
                        else:
                            try:
                                result = pachca.send_dm(uid, msg)
                                msg_id = result.get("id", "?")
                                st.write(f"✅ **{teacher}** → {full_name}: отправлено (msg ID {msg_id})")
                                send_results[teacher] = {"ok": True, "reason": "sent", "detail": f"msg ID {msg_id}"}

                                if create_reminder:
                                    try:
                                        reminder_content = (
                                            f"Заполнить отчёты за "
                                            + ", ".join(
                                                fmt_date(d) for d in sorted(
                                                    st.session_state.teacher_errors[teacher]
                                                    .get("no_report", {}).get("dates", [])
                                                )
                                            )
                                        )
                                        pachca.create_task(uid, reminder_content, reminder_days)
                                        st.write(f"   📌 Напоминание создано")
                                    except Exception as re:
                                        st.write(f"   ⚠️ Напоминание не создано: {re}")

                                time.sleep(0.2)

                            except Exception as e:
                                st.write(f"❌ **{teacher}**: ошибка — {e}")
                                send_results[teacher] = {"ok": False, "reason": "error", "detail": str(e)}

                    st.session_state.send_results = send_results
                    if not dry_run:
                        st.session_state.sending_done = True
                        # Update history: mark sent teachers
                        sent_set = {t for t, r in send_results.items() if r["reason"] == "sent"}
                        if sent_set:
                            update_history_sent(
                                st.session_state.DATE_FROM, st.session_state.DATE_TO,
                                sent_set, reviewer_name,
                            )

                    ok_n  = sum(1 for r in send_results.values() if r["ok"])
                    err_n = sum(1 for r in send_results.values() if not r["ok"])
                    label = (
                        f"{'Тест завершён' if dry_run else 'Готово'}: "
                        f"{ok_n} ✅  {err_n} ❌"
                    )
                    send_status.update(label=label, state="complete", expanded=False)

                    # Update Google Sheets statuses
                    if not dry_run and sheet_id and creds_json:
                        sent_set = {t for t, r in send_results.items() if r["reason"] == "sent"}
                        if sent_set:
                            st.write("📊 Обновляю статусы в Google Таблице…")
                            ok_gs, msg_gs = update_sheet_statuses(
                                sheet_id, sheet_name, sent_set,
                                st.session_state.DATE_FROM, st.session_state.DATE_TO,
                                reviewer_name, creds_json,
                            )
                            st.write(f"   {'✓' if ok_gs else '⚠️'} {msg_gs}")

            # ── Results table ─────────────────────────────────────────────
            if st.session_state.send_results:
                st.divider()
                st.subheader("📊 Результаты рассылки")

                res_df = pd.DataFrame([{
                    "Преподаватель": t,
                    "Статус":   "✅ Отправлено" if r["ok"] else "❌ Ошибка",
                    "Детали":   r["detail"],
                } for t, r in st.session_state.send_results.items()])
                st.dataframe(res_df, use_container_width=True, hide_index=True)

                if st.session_state.sending_done:
                    st.success("✅ Рассылка успешно завершена!")
                    if st.button("🔄 Начать новую проверку"):
                        for k, v in {
                            "loaded": False, "all_errors": [], "teacher_errors": {},
                            "messages": {}, "selected": {}, "send_results": {}, "sending_done": False,
                            "passes_to_set": [], "passes_done": False, "passes_results": {},
                        }.items():
                            st.session_state[k] = v
                        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
#  TAB 4 — ИСТОРИЯ
# ─────────────────────────────────────────────────────────────────────────────

_STATUS_META = {
    "open":         ("🔴 Открыто",                    "#e05252"),
    "message_sent": ("💬 Сообщение отправлено",        "#4a90d9"),
    "pass_set":     ("🚫 Пропуск выставлен",          "#d97f2e"),
    "handled":      ("✅ Обработано",                  "#3aa84b"),
    "skipped":      ("⚪ Пропущено",                  "#888888"),
    "resolved":     ("✅ Обработано",                  "#3aa84b"),   # legacy → same as handled
}
_STATUS_PILL = {
    "open":         "pill pill-open",
    "message_sent": "pill pill-sent",
    "pass_set":     "pill pill-pass",
    "handled":      "pill pill-handled",
    "skipped":      "pill pill-skipped",
    "resolved":     "pill pill-handled",   # legacy alias
}
_STATUS_FILTER_MAP = {
    "🔴 Открыто":                   "open",
    "💬 Сообщение отправлено":       "message_sent",
    "🚫 Пропуск выставлен":         "pass_set",
    "✅ Обработано":                 "handled",
    "⚪ Пропущено":                 "skipped",
}

with tab4:
    st.subheader("📚 История проверок")

    history_records = _all_history

    if not history_records:
        st.info("История пуста — запустите проверку, чтобы записи появились здесь.")
    else:
        # ── Фильтры ───────────────────────────────────────────────────────────
        fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 1])
        with fc1:
            all_hist_teachers = sorted({r["teacher"] for r in history_records})
            f_teacher = st.selectbox("Преподаватель", ["Все"] + all_hist_teachers, key="hf_teacher")
        with fc2:
            f_status = st.selectbox("Статус", ["Все"] + list(_STATUS_FILTER_MAP.keys()), key="hf_status")
        with fc3:
            f_error = st.selectbox(
                "Тип ошибки",
                ["Все", "Нет отчёта", "Нет комментария к пропуску"],
                key="hf_error",
            )
        with fc4:
            st.markdown('<div style="height:1.75rem"></div>', unsafe_allow_html=True)
            if st.button("🔄 Обновить", use_container_width=True, key="hist_refresh"):
                st.rerun()

        # ── Считаем открытые ──────────────────────────────────────────────────
        total_open = sum(1 for r in history_records if r["status"] == "open")
        if total_open:
            st.warning(f"⚠️ Открытых записей: **{total_open}** — требуют внимания")

        # ── Массовые действия ─────────────────────────────────────────────────
        _actionable_statuses = ("open", "message_sent", "pass_set")
        _closeable_statuses  = ("handled", "skipped", "resolved")
        _all_sel_ids         = {r["id"] for r in history_records}

        _n_sel_open   = sum(1 for r in history_records
                            if r["status"] in _actionable_statuses
                            and st.session_state.get(f"hsel_{r['id']}", False))
        _n_sel_closed = sum(1 for r in history_records
                            if r["status"] in _closeable_statuses
                            and st.session_state.get(f"hsel_{r['id']}", False))
        _n_sel_total  = _n_sel_open + _n_sel_closed

        with st.container():
            _ba1, _ba2, _ba3, _ba4, _ba5, _ba6 = st.columns([1.4, 1.4, 1.6, 1.6, 1.6, 2])

            if _ba1.button("☑️ Выбрать все", use_container_width=True, key="hist_sel_all"):
                for r in history_records:
                    st.session_state[f"hsel_{r['id']}"] = True
                st.rerun()
            if _ba2.button("⬜ Снять всё", use_container_width=True, key="hist_desel_all"):
                for r in history_records:
                    st.session_state[f"hsel_{r['id']}"] = False
                st.rerun()

            if _n_sel_open > 0:
                if _ba3.button(f"✅ Обработано ({_n_sel_open})", use_container_width=True,
                               key="hist_bulk_handled", type="primary"):
                    for r in history_records:
                        if st.session_state.get(f"hsel_{r['id']}", False) and r["status"] in _actionable_statuses:
                            update_history_record(r["id"], "handled", r.get("reviewer_comment", ""), reviewer_name)
                            st.session_state[f"hsel_{r['id']}"] = False
                    st.rerun()
                if _ba4.button(f"⚪ Пропустить ({_n_sel_open})", use_container_width=True,
                               key="hist_bulk_skip"):
                    for r in history_records:
                        if st.session_state.get(f"hsel_{r['id']}", False) and r["status"] in _actionable_statuses:
                            update_history_record(r["id"], "skipped", r.get("reviewer_comment", ""), reviewer_name)
                            st.session_state[f"hsel_{r['id']}"] = False
                    st.rerun()

            if _n_sel_closed > 0:
                if _ba5.button(f"↺ Переоткрыть ({_n_sel_closed})", use_container_width=True,
                               key="hist_bulk_reopen"):
                    for r in history_records:
                        if st.session_state.get(f"hsel_{r['id']}", False) and r["status"] in _closeable_statuses:
                            update_history_record(r["id"], "open", r.get("reviewer_comment", ""), reviewer_name)
                            st.session_state[f"hsel_{r['id']}"] = False
                    st.rerun()

            if _n_sel_total > 0:
                _ba6.markdown(
                    f"<span style='line-height:2.4rem;color:#a78bfa;font-weight:600'>"
                    f"Выбрано: {_n_sel_total}</span>",
                    unsafe_allow_html=True,
                )

        st.divider()

        # ── Фильтр по дате занятия ────────────────────────────────────────────
        _ERR_FILTER_MAP = {
            "Нет отчёта":                 "no_report",
            "Нет комментария к пропуску": "no_abs_comment",
        }
        _all_dates = sorted({r["date"] for r in history_records if r.get("date")})
        if _all_dates:
            _hd_min = datetime.strptime(_all_dates[0],  "%Y-%m-%d").date()
            _hd_max = datetime.strptime(_all_dates[-1], "%Y-%m-%d").date()
        else:
            _hd_min = _hd_max = date.today()

        fd1, fd2 = st.columns(2)
        with fd1:
            _h_date_from = st.date_input("Дата занятия с", value=_hd_min,
                                         min_value=_hd_min, max_value=_hd_max, key="hf_date_from")
        with fd2:
            _h_date_to   = st.date_input("по",             value=_hd_max,
                                         min_value=_hd_min, max_value=_hd_max, key="hf_date_to")
        _hdf = _h_date_from.strftime("%Y-%m-%d")
        _hdt = _h_date_to.strftime("%Y-%m-%d")

        # ── Применяем все фильтры одним проходом ──────────────────────────────
        _f_status_val = _STATUS_FILTER_MAP.get(f_status)
        _f_error_val  = _ERR_FILTER_MAP.get(f_error)
        all_filtered = [
            r for r in history_records
            if (f_teacher == "Все" or r["teacher"] == f_teacher)
            and (f_status  == "Все" or r["status"]     == _f_status_val)
            and (f_error   == "Все" or r["error_type"] == _f_error_val)
            and _hdf <= r.get("date", "") <= _hdt
        ]

        # ── Группируем по дате занятия (новые сверху) ─────────────────────────
        from collections import defaultdict as _hdd, Counter as _HCnt
        _by_date: dict = _hdd(list)
        for r in all_filtered:
            _by_date[r.get("date", "")].append(r)

        for lesson_date in sorted(_by_date.keys(), reverse=True):
            recs_for_date = _by_date[lesson_date]
            try:
                date_label = datetime.strptime(lesson_date, "%Y-%m-%d").strftime("%d.%m.%Y")
            except Exception:
                date_label = lesson_date

            _hcnt     = _HCnt(r["status"] for r in recs_for_date)
            open_n    = _hcnt.get("open", 0)
            wip_n     = sum(_hcnt.get(s, 0) for s in ("message_sent", "pass_set"))
            handled_n = sum(_hcnt.get(s, 0) for s in ("handled", "resolved"))
            header    = f"📅 {date_label}  ·  {len(recs_for_date)} записей"
            if open_n:
                header += f"  ·  🔴 {open_n} открытых"
            if wip_n:
                header += f"  ·  ⏳ {wip_n} в работе"
            if handled_n:
                header += f"  ·  ✅ {handled_n} исправлено"

            with st.expander(header, expanded=(open_n > 0)):
                _date_ids        = {r["id"] for r in recs_for_date}
                _date_action_ids = {r["id"] for r in recs_for_date if r["status"] in _actionable_statuses}
                _date_close_ids  = {r["id"] for r in recs_for_date if r["status"] in _closeable_statuses}
                _pb1, _pb2, _pb3, _pb4, _ = st.columns([1.2, 1.2, 1.6, 1.6, 4])
                if _pb1.button("☑️ Все", key=f"dsel_{lesson_date}", use_container_width=True):
                    for rid in _date_ids:
                        st.session_state[f"hsel_{rid}"] = True
                    st.rerun()
                if _pb2.button("⬜ Снять", key=f"ddes_{lesson_date}", use_container_width=True):
                    for rid in _date_ids:
                        st.session_state[f"hsel_{rid}"] = False
                    st.rerun()
                _d_sel_open   = sum(1 for rid in _date_action_ids if st.session_state.get(f"hsel_{rid}", False))
                _d_sel_closed = sum(1 for rid in _date_close_ids  if st.session_state.get(f"hsel_{rid}", False))
                if _d_sel_open > 0:
                    if _pb3.button(f"✅ Обработано ({_d_sel_open})", key=f"dh_{lesson_date}",
                                   use_container_width=True, type="primary"):
                        for r in recs_for_date:
                            if st.session_state.get(f"hsel_{r['id']}", False) and r["status"] in _actionable_statuses:
                                update_history_record(r["id"], "handled", r.get("reviewer_comment", ""), reviewer_name)
                                st.session_state[f"hsel_{r['id']}"] = False
                        st.rerun()
                if _d_sel_closed > 0:
                    if _pb4.button(f"↺ Переоткрыть ({_d_sel_closed})", key=f"dr_{lesson_date}",
                                   use_container_width=True):
                        for r in recs_for_date:
                            if st.session_state.get(f"hsel_{r['id']}", False) and r["status"] in _closeable_statuses:
                                update_history_record(r["id"], "open", r.get("reviewer_comment", ""), reviewer_name)
                                st.session_state[f"hsel_{r['id']}"] = False
                        st.rerun()
                st.divider()

                for rec in sorted(recs_for_date, key=lambda r: r["teacher"]):
                    s_label, s_color = _STATUS_META.get(rec["status"], ("❓", "#888"))

                    with st.container(border=True):
                        r_c0, r_c1, r_c2, r_c3 = st.columns([0.25, 3, 2, 3])
                        with r_c0:
                            st.checkbox(
                                "sel",
                                key=f"hsel_{rec['id']}",
                                label_visibility="collapsed",
                            )

                        with r_c1:
                            try:
                                d_str = datetime.strptime(rec["date"], "%Y-%m-%d").strftime("%d.%m.%Y")
                            except Exception:
                                d_str = rec["date"]
                            _ibadge = intern_badge(rec["teacher"], _teacher_info)
                            st.markdown(
                                f"**{rec['teacher']}**&nbsp;&nbsp;{_ibadge}",
                                unsafe_allow_html=True,
                            )
                            err_icon = "📋" if rec["error_type"] == "no_report" else "💬"
                            cnt_str = f" · {rec['count']} шт." if rec.get("count") else ""
                            st.caption(f"{err_icon} {rec['error_description']}{cnt_str}  ·  {d_str}")
                            _students = rec.get("students", [])
                            if _students:
                                with st.expander(f"👥 {len(_students)} учеников", expanded=False):
                                    for _sn in _students:
                                        st.caption(f"• {_sn}")

                        with r_c2:
                            pill_cls = _STATUS_PILL.get(rec["status"], "pill")
                            st.markdown(
                                f'<span class="{pill_cls}">{s_label}</span>',
                                unsafe_allow_html=True,
                            )
                            upd = rec.get("updated_at", "")
                            if upd:
                                try:
                                    upd_str = datetime.fromisoformat(upd).strftime("%d.%m %H:%M")
                                except Exception:
                                    upd_str = upd
                                st.caption(f"🕐 {upd_str}")
                            if rec.get("reviewer"):
                                st.caption(f"👤 {rec['reviewer']}")

                        with r_c3:
                            rid = rec["id"]
                            if rec["status"] in ("open", "message_sent", "pass_set"):
                                new_comment = st.text_input(
                                    "Комментарий",
                                    value=rec.get("reviewer_comment", ""),
                                    key=f"hc_{rid}",
                                    placeholder="Добавьте комментарий…",
                                    label_visibility="collapsed",
                                )
                                btn1, btn2 = st.columns(2)
                                if btn1.button("✅ Обработано", key=f"hh_{rid}", use_container_width=True):
                                    update_history_record(rid, "handled", new_comment, reviewer_name)
                                    st.rerun()
                                if btn2.button("⚪ Пропустить", key=f"hs_{rid}", use_container_width=True):
                                    update_history_record(rid, "skipped", new_comment, reviewer_name)
                                    st.rerun()
                            else:
                                comment = rec.get("reviewer_comment", "")
                                if comment:
                                    st.caption(f"💬 {comment}")
                                if st.button("↺ Переоткрыть", key=f"hr_{rid}", use_container_width=True):
                                    update_history_record(rid, "open", rec.get("reviewer_comment", ""), reviewer_name)
                                    st.rerun()

        # ── Экспорт всей истории ──────────────────────────────────────────────
        st.divider()
        if all_filtered:
            def _fmt_date(d):
                try: return datetime.strptime(d, "%Y-%m-%d").strftime("%d.%m.%Y")
                except: return d

            export_df = pd.DataFrame([{
                "Дата":           _fmt_date(r["date"]),
                "Преподаватель":  r["teacher"],
                "Тег ученика":    r.get("student_tag", "Все"),
                "Ошибка":         r["error_description"],
                "Кол-во":         r.get("count", ""),
                "Статус":         _STATUS_META.get(r["status"], ("?", ""))[0],
                "Проверяющий":    r.get("reviewer", ""),
                "Комментарий":    r.get("reviewer_comment", ""),
                "Обновлено":      r.get("updated_at", ""),
            } for r in sorted(all_filtered, key=lambda r: r.get("date",""), reverse=True)])
            csv_hist = export_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Скачать историю (CSV)",
                csv_hist,
                "history.csv",
                "text/csv",
                use_container_width=False,
            )


# ─────────────────────────────────────────────────────────────────────────────
#  TAB 5 — СТАТИСТИКА
# ─────────────────────────────────────────────────────────────────────────────

def _is_processed(r):
    # Only truly resolved: teacher wrote report (auto-handled on re-check) or manually closed
    return r["status"] in ("handled", "resolved")

def _is_in_progress(r):
    # We took action, but not yet confirmed fixed by re-check
    return r["status"] in ("message_sent", "pass_set")

def _is_open(r):
    return r["status"] == "open"

with tab5:
    _s_hdr, _s_refresh = st.columns([5, 1])
    _s_hdr.subheader("📊 Статистика по проверкам")
    if _s_refresh.button("🔄 Обновить", key="stat_refresh", use_container_width=True):
        st.rerun()

    stat_records = _all_history

    if not stat_records:
        st.info("История пуста — запустите проверку для получения статистики.")
    else:
        # ── Filters row ───────────────────────────────────────────────────────
        all_stat_periods = sorted(
            set((r["period_from"], r["period_to"]) for r in stat_records),
            key=lambda x: x[0], reverse=True,
        )

        _fmode_col, _fp1_col, _fp2_col, _ftype_col = st.columns([1.5, 1.5, 1.5, 1.5])
        with _fmode_col:
            filter_mode = st.radio(
                "Режим фильтра",
                ["Готовые периоды", "Произвольный диапазон"],
                horizontal=True,
                key="stat_filter_mode",
                label_visibility="collapsed",
            )

        if filter_mode == "Готовые периоды":
            period_labels = ["Все периоды"] + [
                f"{fmt_date(pf)} — {fmt_date(pt)}" for pf, pt in all_stat_periods
            ]
            with _fp1_col:
                sel_period = st.selectbox("Выберите период", period_labels, key="stat_period_sel",
                                          label_visibility="collapsed")
            if sel_period == "Все периоды":
                sr = stat_records
            else:
                pidx = period_labels.index(sel_period) - 1
                pf0, pt0 = all_stat_periods[pidx]
                sr = [r for r in stat_records if r["period_from"] == pf0 and r["period_to"] == pt0]
        else:
            _hist_dates = sorted(r["date"] for r in stat_records)
            _min_d = datetime.strptime(_hist_dates[0], "%Y-%m-%d").date() if _hist_dates else date.today() - timedelta(days=30)
            _max_d = datetime.strptime(_hist_dates[-1], "%Y-%m-%d").date() if _hist_dates else date.today()
            with _fp1_col:
                custom_from = st.date_input("С", value=_min_d, key="stat_custom_from")
            with _fp2_col:
                custom_to   = st.date_input("По", value=_max_d, key="stat_custom_to")
            _cf = custom_from.strftime("%Y-%m-%d")
            _ct = custom_to.strftime("%Y-%m-%d")
            sr = [r for r in stat_records if _cf <= r["date"] <= _ct]
            sel_period = "custom"

        # ── Error type filter ─────────────────────────────────────────────────
        with _ftype_col:
            _err_type_opt = st.selectbox(
                "Тип ошибки",
                ["Все типы", "📋 Нет отчёта", "💬 Нет комментария"],
                key="stat_err_type",
                label_visibility="collapsed",
            )
        if _err_type_opt == "📋 Нет отчёта":
            sr = [r for r in sr if r.get("error_type") == "no_report"]
        elif _err_type_opt == "💬 Нет комментария":
            sr = [r for r in sr if r.get("error_type") == "no_abs_comment"]

        # ── Single pass: group by teacher + count statuses ───────────────────
        from collections import Counter as _Counter, defaultdict as _dd
        _sr_by_teacher: dict = _dd(list)
        _sr_status_cnt = _Counter()
        _sr_teachers   = set()
        for _r in sr:
            _sr_by_teacher[_r["teacher"]].append(_r)
            _sr_status_cnt[_r["status"]] += 1
            _sr_teachers.add(_r["teacher"])

        _processed_statuses = {"handled", "resolved"}
        _in_progress_statuses = {"message_sent", "pass_set"}
        total_rec      = len(sr)
        processed_n    = sum(_sr_status_cnt.get(s, 0) for s in _processed_statuses)
        in_progress_n  = sum(_sr_status_cnt.get(s, 0) for s in _in_progress_statuses)
        open_n         = _sr_status_cnt.get("open", 0)
        teachers_n     = len(_sr_teachers)

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("📋 Всего ошибок",      total_rec)
        m2.metric("✅ Исправлено",         processed_n,
                  help="Ошибки, где преподаватель реально написал отчёт (подтверждено перепроверкой) или закрыты вручную")
        m3.metric("⏳ В работе",           in_progress_n,
                  help="Отправлено напоминание или выставлен пропуск — ждём, пока преподаватель исправит")
        m4.metric("🔴 Не обработано",      open_n,
                  help="Ошибки, по которым ещё не предпринято никаких действий")
        m5.metric("👩‍🏫 Преподавателей",    teachers_n)

        st.divider()

        # ── Two-column layout: teacher table + summary ────────────────────────
        col_tbl, col_sum = st.columns([3, 1])

        with col_tbl:
            st.markdown("**По преподавателям**")

            teacher_rows = []
            for teacher in sorted(_sr_by_teacher.keys()):
                t     = _sr_by_teacher[teacher]
                t_cnt = _Counter(r["status"] for r in t)
                done  = sum(t_cnt.get(s, 0) for s in _processed_statuses)
                wip   = sum(t_cnt.get(s, 0) for s in _in_progress_statuses)
                total = len(t)
                teacher_rows.append({
                    "Преподаватель":    teacher,
                    "✅ Исправлено":    done,
                    "⏳ В работе":      wip,
                    "🔴 Открыто":       t_cnt.get("open", 0),
                    "Всего":            total,
                    "% исправлено":     round(done / total * 100) if total else 0,
                })

            df_teachers = pd.DataFrame(teacher_rows)

            st.dataframe(
                df_teachers,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Преподаватель":    st.column_config.TextColumn(width="large"),
                    "✅ Исправлено":    st.column_config.NumberColumn(width="small"),
                    "⏳ В работе":      st.column_config.NumberColumn(width="small"),
                    "🔴 Открыто":       st.column_config.NumberColumn(width="small"),
                    "Всего":            st.column_config.NumberColumn(width="small"),
                    "% исправлено":     st.column_config.ProgressColumn(
                        "% исправлено",
                        help="Доля ошибок, где преподаватель реально написал отчёт",
                        format="%d%%",
                        min_value=0,
                        max_value=100,
                    ),
                },
            )

            csv_stat = df_teachers.to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ Скачать CSV", csv_stat, "stats.csv", "text/csv")

        with col_sum:
            st.markdown("**Сводка по статусам**")
            summary_rows = [
                {"Статус": "✅ Исправлено",            "Кол-во": _sr_status_cnt.get("handled", 0) + _sr_status_cnt.get("resolved", 0)},
                {"Статус": "💬 Напоминание отправлено", "Кол-во": _sr_status_cnt.get("message_sent", 0)},
                {"Статус": "🚫 Пропуск выставлен",     "Кол-во": _sr_status_cnt.get("pass_set", 0)},
                {"Статус": "🔴 Открыто",               "Кол-во": _sr_status_cnt.get("open", 0)},
                {"Статус": "⚪ Пропущено",              "Кол-во": _sr_status_cnt.get("skipped", 0)},
            ]
            st.dataframe(
                pd.DataFrame(summary_rows),
                use_container_width=True,
                hide_index=True,
            )

        # ── Per-period dynamics (shown when "All periods" selected) ───────────
        if sel_period in ("Все периоды", "custom") and len(all_stat_periods) > 1:
            st.divider()
            st.markdown("**📅 Динамика по периодам**")

            # Pre-group stat_records by period key — O(N) instead of O(N×P)
            _stat_by_period: dict = _dd(list)
            for _r in stat_records:
                _stat_by_period[(_r["period_from"], _r["period_to"])].append(_r)

            period_rows = []
            for pf, pt in all_stat_periods:
                p = _stat_by_period.get((pf, pt), [])
                p_done = sum(1 for r in p if _is_processed(r))
                period_rows.append({
                    "Период":           f"{fmt_date(pf)} — {fmt_date(pt)}",
                    "✅ Обработано":    p_done,
                    "🔴 Не обработано": sum(1 for r in p if _is_open(r)),
                    "Всего":            len(p),
                    "% обработано":     round(p_done / len(p) * 100) if p else 0,
                })

            st.dataframe(
                pd.DataFrame(period_rows),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Период":           st.column_config.TextColumn(width="medium"),
                    "✅ Обработано":    st.column_config.NumberColumn(width="small"),
                    "🔴 Не обработано": st.column_config.NumberColumn(width="small"),
                    "Всего":            st.column_config.NumberColumn(width="small"),
                    "% обработано":     st.column_config.ProgressColumn(
                        "% обработано",
                        format="%d%%",
                        min_value=0,
                        max_value=100,
                    ),
                },
            )

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 6 — РАССЫЛКА
# ─────────────────────────────────────────────────────────────────────────────

with tab6:
    st.subheader("📢 Рассылка преподавателям")

    _ti = _teacher_info  # already loaded above

    if not _ti:
        st.warning("Данные о преподавателях не загружены. Настройте Google Sheets в ⚙️.")
    else:
        # ── Фильтры — строка 1: три селекта ──────────────────────────────────
        bc1, bc2, bc3 = st.columns(3)

        with bc1:
            category_opt = st.selectbox(
                "Категория",
                ["Все", "Основной состав", "Стажёры"],
                key="bc_category",
            )
        with bc2:
            all_subjects = sorted({
                v["subject"] for v in _ti.values()
                if v.get("subject") and v["subject"] not in ("", "nan", "-")
            })
            subject_opt = st.selectbox(
                "Предмет",
                ["Все"] + all_subjects,
                key="bc_subject",
            )
        with bc3:
            all_statuses = sorted({
                v["status"] for v in _ti.values()
                if v.get("status") and v["status"] not in ("", "nan", "-")
            })
            status_opt = st.selectbox(
                "Статус набора",
                ["Все"] + all_statuses,
                key="bc_status",
            )

        # ── Отфильтрованный список ─────────────────────────────────────────
        # Read hide_fired from state before the checkbox widget is rendered
        hide_fired = st.session_state.get("bc_hide_fired", True)

        def _bc_filter(name, info):
            if hide_fired and info.get("fired", False):
                return False
            if status_opt != "Все" and info.get("status") != status_opt:
                return False
            if category_opt == "Стажёры" and not info.get("is_intern"):
                return False
            if category_opt == "Основной состав" and info.get("is_intern"):
                return False
            if subject_opt != "Все" and info.get("subject") != subject_opt:
                return False
            return True

        filtered_teachers = {
            name: info for name, info in _ti.items()
            if _bc_filter(name, info)
        }

        # Group by subject for display
        by_subject: dict = {}
        for name, info in filtered_teachers.items():
            subj = info.get("subject") or "Без предмета"
            by_subject.setdefault(subj, []).append((name, info))
        for subj in by_subject:
            by_subject[subj].sort(key=lambda x: x[0])

        total_filtered = len(filtered_teachers)

        # ── Строка 2: действия + чекбокс ─────────────────────────────────────
        st.divider()
        sel_c1, sel_c2, sel_c3, sel_c4 = st.columns([2, 2, 2, 2])
        if sel_c1.button("☑️ Выбрать всех", key="bc_sel_all", use_container_width=True):
            for nm in filtered_teachers:
                st.session_state[f"bc_chk_{nm}"] = True
            st.rerun()
        if sel_c2.button("🔲 Снять всех", key="bc_desel_all", use_container_width=True):
            for nm in filtered_teachers:
                st.session_state[f"bc_chk_{nm}"] = False
            st.rerun()

        n_bc_selected = sum(
            1 for nm in filtered_teachers
            if st.session_state.get(f"bc_chk_{nm}", False)
        )
        sel_c3.markdown(f"<div style='padding-top:0.6rem'><b>Выбрано: {n_bc_selected} из {total_filtered}</b></div>", unsafe_allow_html=True)
        with sel_c4:
            st.checkbox("Скрыть уволенных", value=True, key="bc_hide_fired")

        # ── Teacher list grouped by subject ──────────────────────────────
        for subj in sorted(by_subject.keys()):
            teachers_in_subj = by_subject[subj]
            with st.expander(f"**{subj}** · {len(teachers_in_subj)} чел.", expanded=(subject_opt != "Все")):
                for nm, info in teachers_in_subj:
                    chk_key = f"bc_chk_{nm}"
                    c1, c2 = st.columns([5, 1])
                    with c1:
                        is_intern = info.get("is_intern", False)
                        intern_badge = ' <span class="pill pill-pass">🎓 Стажёр</span>' if is_intern else ""
                        st.markdown(f"{nm}{intern_badge}", unsafe_allow_html=True)
                    with c2:
                        st.checkbox("", key=chk_key, label_visibility="collapsed")

        st.divider()

        # ── Message composer ──────────────────────────────────────────────
        st.markdown("**✏️ Текст сообщения**")
        bc_message = st.text_area(
            "Текст",
            placeholder="Введите текст рассылки...",
            height=150,
            key="bc_message_text",
            label_visibility="collapsed",
        )

        # ── Send button ───────────────────────────────────────────────────
        recipients = [nm for nm in filtered_teachers if st.session_state.get(f"bc_chk_{nm}", False)]

        if st.button(
            f"📤 Отправить {len(recipients)} преподавателям",
            key="bc_send_btn",
            disabled=not recipients or not bc_message.strip(),
            type="primary",
            use_container_width=False,
        ):
            _bc_pachca_token = cfg.get("pachca_token", "")
            if not _bc_pachca_token:
                st.error("Токен Pachca не настроен — откройте ⚙️.")
            else:
                pachca = PachcaAPI(_bc_pachca_token)
                with st.status("Загружаю список пользователей Pachca…", expanded=True) as bc_status:
                    pachca.load_users(write_fn=st.write)
                    ok_list, fail_list = [], []
                    for nm in recipients:
                        user = pachca.find_user(nm)
                        if not user:
                            fail_list.append((nm, "Не найден в Pachca"))
                            st.write(f"⚠️ {nm} — не найден")
                            continue
                        try:
                            pachca.send_dm(user["id"], bc_message.strip())
                            ok_list.append(nm)
                            st.write(f"✅ {nm}")
                        except Exception as _e:
                            fail_list.append((nm, str(_e)))
                            st.write(f"❌ {nm} — {_e}")

                    if ok_list:
                        bc_status.update(label=f"✅ Отправлено: {len(ok_list)}, ошибок: {len(fail_list)}", state="complete")
                    else:
                        bc_status.update(label="❌ Не удалось отправить ни одного сообщения", state="error")

                if fail_list:
                    st.warning("Не найдены в Pachca:\n" + "\n".join(f"• {nm}: {reason}" for nm, reason in fail_list))
