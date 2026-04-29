import streamlit as st
import requests
import re
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

_last_save_error: str = ""   # module-level for debug display

def save_history(records: list):
    global _last_save_error
    _last_save_error = ""
    saved = False

    # 1. Supabase (primary) — always try
    sb, _ = _get_supabase_client()
    if sb is not None:
        try:
            rows = [_hist_dict_to_sb_row(r) for r in records]
            if rows:
                sb.table("history").upsert(rows).execute()
            saved = True
        except Exception as _e:
            _last_save_error = str(_e)

    # 2. Google Sheets (mirror) — always try in parallel with Supabase
    ws, _ = _get_history_sheet()
    if ws is not None:
        try:
            all_values = [_HIST_HEADERS] + [_hist_dict_to_row(r) for r in records]
            ws.clear()
            if all_values:
                ws.update("A1", all_values)
            saved = True
        except Exception:
            pass

    # 3. Local JSON (last resort — only if both cloud backends failed)
    if not saved:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2, default=str)


@st.dialog("⚠️ Некорректный отчёт")
def _flag_report_dialog(rec: dict):
    st.markdown(f"**Преподаватель:** {rec['teacher']}")
    st.markdown(f"**Ученик:** {rec['student']} &nbsp;·&nbsp; **Дата:** {rec['date']} &nbsp;·&nbsp; **Предмет:** {rec['subject']}")
    _preview = rec["comment"][:300] + ("…" if len(rec["comment"]) > 300 else "")
    st.caption(_preview)
    st.divider()
    note = st.text_area(
        "Что не так с отчётом?",
        placeholder="Например: нет конкретики, оценка не обоснована, скопировано у других учеников…",
        key="flag_note",
    )
    col1, col2 = st.columns(2)
    if col1.button("Добавить в историю", type="primary", use_container_width=True):
        import uuid as _uuid
        now = datetime.now().isoformat()
        _reviewer = cfg.get("reviewer_name", "Артём")
        new_rec = {
            "id":                str(_uuid.uuid4())[:8],
            "checked_at":        now,
            "period_from":       rec["date"],
            "period_to":         rec["date"],
            "date":              rec["date"],
            "teacher":           rec["teacher"],
            "student_tag":       rec["student"],
            "error_type":        "bad_report",
            "error_description": note.strip() if note.strip() else "Некорректный отчёт",
            "count":             1,
            "students":          [rec["student"]],
            "status":            "open",
            "reviewer":          _reviewer,
            "reviewer_comment":  note.strip(),
            "updated_at":        now,
        }
        all_hist = load_history()
        all_hist.append(new_rec)
        save_history(all_hist)
        st.success("✅ Добавлено в историю!")
        st.rerun()
    if col2.button("Отмена", use_container_width=True):
        st.rerun()


@st.dialog("📋 Жалоба на отчёты преподавателя")
def _flag_teacher_dialog(teacher: str, records: list):
    st.markdown(f"**Преподаватель:** {teacher}")
    _dates = sorted({r["date"] for r in records})
    try:
        _dates_fmt = ", ".join(datetime.strptime(d, "%Y-%m-%d").strftime("%d.%m") for d in _dates)
    except Exception:
        _dates_fmt = ", ".join(_dates)
    st.caption(f"Занятий: {len(_dates)} · Учеников: {len(records)} · Даты: {_dates_fmt}")
    st.divider()

    reason = st.radio(
        "Причина",
        ["Неинформативный отчёт", "Некорректный отчёт", "Другое"],
        key="flag_teacher_reason",
        horizontal=True,
    )
    note = st.text_area(
        "Комментарий (необязательно)" if reason != "Другое" else "Комментарий",
        placeholder="Уточни что именно не так…",
        key="flag_teacher_note",
    )

    col1, col2 = st.columns(2)
    if col1.button("Добавить в историю", type="primary", use_container_width=True):
        import uuid as _uuid
        now = datetime.now().isoformat()
        _reviewer = cfg.get("reviewer_name", "Артём")
        description = reason
        if note.strip():
            description += f": {note.strip()}"

        new_recs = []
        for _d in _dates:
            _day_recs     = [r for r in records if r["date"] == _d]
            _day_students = [r["student"] for r in _day_recs]
            # Собираем тексты отчётов по студентам для этого дня
            _reports_text = "\n".join(
                f"• {r['student']}: {r['comment']}" if r.get("comment") else f"• {r['student']}: —"
                for r in _day_recs
            )
            new_recs.append({
                "id":                str(_uuid.uuid4())[:8],
                "checked_at":        now,
                "period_from":       _d,
                "period_to":         _d,
                "date":              _d,
                "teacher":           teacher,
                "student_tag":       ", ".join(_day_students) if _day_students else "Все",
                "error_type":        "bad_report",
                "error_description": description,
                "count":             len(_day_students),
                "students":          _day_students,
                "status":            "open",
                "reviewer":          _reviewer,
                "reviewer_comment":  note.strip(),
                "updated_at":        now,
            })

        all_hist = load_history()
        all_hist.extend(new_recs)
        save_history(all_hist)
        st.success(f"✅ Добавлено {len(new_recs)} записей в историю!")
        st.rerun()
    if col2.button("Отмена", use_container_width=True):
        st.rerun()


def upsert_history(all_errors: list, period_from: str, period_to: str, reviewer: str,
                   resolved_keys: set = None):
    """Add/update history records. Global dedup by (teacher, date, error_type).

    resolved_keys: set of (teacher, date, error_type) confirmed resolved by a real fix
    (e.g. report actually written). Used to decide whether pass_set → handled.
    If None, only open/message_sent are auto-closed.
    """
    import uuid as _uuid
    records = load_history()
    now = datetime.now().isoformat(timespec="seconds")
    if resolved_keys is None:
        resolved_keys = set()

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
            rec = rec.copy()
            # A record is genuinely resolved when:
            # - error_type is no_abs_comment (teacher added absence comment — real fix), OR
            # - key is in resolved_keys (report was actually written, confirmed by API)
            # If neither — lesson was likely cancelled (отмена): keep current status visible.
            _genuine = (rec.get("error_type") == "no_abs_comment") or (key in resolved_keys)
            if rec["status"] in ("open", "message_sent", "pass_set"):
                if _genuine:
                    rec.update({"status": "handled", "updated_at": now})
                # else: keep current status — reviewer should verify the cancellation
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


def update_history_description(record_id: str, description: str, error_type: str = None):
    """Update a single record's error_description and optionally error_type."""
    records = load_history()
    now = datetime.now().isoformat(timespec="seconds")
    for rec in records:
        if rec["id"] == record_id:
            upd = {"error_description": description, "updated_at": now}
            if error_type:
                upd["error_type"] = error_type
            rec.update(upd)
            break
    save_history(records)


def delete_history_record(record_id: str):
    """Remove a single record by id."""
    records = [r for r in load_history() if r["id"] != record_id]
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
        def _get_ci(row, *candidates):
            """Case-insensitive row.get supporting ё/е equivalence."""
            row_lower = {k.lower().replace("ё", "е"): v for k, v in row.items()}
            for c in candidates:
                val = row_lower.get(c.lower().replace("ё", "е"))
                if val is not None:
                    return val
            return ""

        result = {}
        for row in rows:
            name = str(_get_ci(row, "ФИО")).strip()
            if not name or name.lower() == "nan":
                continue
            _raw_status = str(_get_ci(row, "Статус набора")).strip()
            _role_val = str(_get_ci(row, "Статус")).strip().lower().replace("ё", "е")
            result[name] = {
                "is_intern": _role_val in ("стажер", "стажёр"),
                "subject":   str(_get_ci(row, "Предмет")).strip(),
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
#  HEADER — title + dates + settings popover
# ─────────────────────────────────────────────────────────────────────────────

# Sensitive credentials — read from Secrets only (not editable in UI)
subdomain    = _read_secret("subdomain")    or cfg.get("subdomain", "matrix")
api_key      = _read_secret("api_key")      or cfg.get("api_key", "")
pachca_token = _read_secret("pachca_token") or cfg.get("pachca_token", "")

_title_col, _gear_col = st.columns([5, 0.7])

with _title_col:
    _sb_check, _sb_check_err = _get_supabase_client()
    if _sb_check is None:
        _status_badge = (
            '<span title="Supabase недоступен — данные могут не сохраняться" '
            'style="display:inline-flex;align-items:center;gap:5px;'
            'background:rgba(239,68,68,.15);color:#f87171;'
            'border:1px solid rgba(239,68,68,.3);border-radius:8px;'
            'font-size:.72rem;font-weight:600;padding:2px 9px;'
            'vertical-align:middle;margin-left:10px;cursor:default;">'
            '⚠️ БД недоступна</span>'
        )
    else:
        _status_badge = ""
    st.markdown(
        f'## 📋 Проверка отчётов HolliHop{_status_badge}',
        unsafe_allow_html=True,
    )

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

        st.markdown("**📑 Отчёты**")
        short_comment_threshold = st.number_input(
            "Порог «короткого» отчёта (символов)",
            min_value=50, max_value=2000,
            value=int(cfg.get("short_comment_threshold", 400)),
            step=50, key="cfg_short_threshold",
            help="Отчёты короче этого значения помечаются 🟡",
        )

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
            if _last_save_error:
                st.error(f"⚠️ Последняя ошибка записи: `{_last_save_error}`")
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

            # Restore from JSON upload
            st.markdown("**📂 Восстановить из JSON-бэкапа**")
            _restore_file = st.file_uploader(
                "Загрузи файл history_backup_*.json", type=["json"],
                key="restore_json_upload",
            )
            if _restore_file is not None:
                if st.button("📥 Загрузить в Supabase", key="restore_json_btn",
                             use_container_width=True):
                    with st.spinner("Восстанавливаю…"):
                        try:
                            _restore_data = json.loads(_restore_file.read().decode("utf-8"))
                            save_history(_restore_data)
                            st.success(f"✅ Восстановлено {len(_restore_data)} записей")
                        except Exception as _re:
                            st.error(f"Ошибка: {_re}")

            # Restore missing records from Google Sheets mirror
            st.markdown("**📗 Восстановить из Google Sheets**")
            st.caption("Найти записи, которые есть в Sheets-зеркале, но пропали из Supabase, и восстановить их.")
            if st.button("🔄 Восстановить пропавшие записи из Sheets",
                         key="restore_from_sheets_btn", use_container_width=True):
                _ws_restore, _ = _get_history_sheet()
                if _ws_restore is None:
                    st.error("❌ Google Sheets не подключён")
                else:
                    with st.status("Восстанавливаю из Sheets…", expanded=True) as _rs_status:
                        try:
                            st.write("📖 Читаю данные из Sheets…")
                            _rows_sh = _ws_restore.get_all_values()
                            _recs_sh = [_hist_row_to_dict(r) for r in _rows_sh[1:] if any(r)]
                            st.write(f"   Найдено {len(_recs_sh)} записей в Sheets")

                            st.write("📖 Читаю данные из Supabase…")
                            _recs_sb = load_history()
                            _sb_ids = {r["id"] for r in _recs_sb}
                            st.write(f"   Найдено {len(_recs_sb)} записей в Supabase")

                            _missing = [r for r in _recs_sh if r["id"] not in _sb_ids]
                            st.write(f"🔍 Пропавших записей: **{len(_missing)}**")

                            if _missing:
                                # Merge: existing Supabase + missing from Sheets
                                _merged = _recs_sb + _missing
                                save_history(_merged)
                                st.write(f"✅ Восстановлено {len(_missing)} записей")
                                _rs_status.update(
                                    label=f"✅ Восстановлено {len(_missing)} записей",
                                    state="complete",
                                )
                            else:
                                st.write("✅ Все записи на месте, восстанавливать нечего")
                                _rs_status.update(
                                    label="✅ Данные в Supabase актуальны",
                                    state="complete",
                                )
                        except Exception as _rse:
                            st.error(f"Ошибка: {_rse}")
                            _rs_status.update(label="❌ Ошибка", state="error")
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
3. **Settings → API** → скопируй `URL` и `service_role key`
4. Добавь в Streamlit Secrets:
```toml
supabase_url = "https://xxxx.supabase.co"
supabase_key = "eyJ..."
```
""")
            # Migration from Sheets → Supabase (when Supabase not yet connected)
            if st.button("📥 Migrate из Sheets в Supabase",
                         key="migrate_to_sb", use_container_width=True):
                with st.spinner("Читаю данные из Sheets…"):
                    _ws_migrate, _ = _get_history_sheet()
                    if _ws_migrate is None:
                        st.error("Sheets не подключён, нечего мигрировать")
                    else:
                        try:
                            _rows_m = _ws_migrate.get_all_values()
                            _recs_m = [_hist_row_to_dict(r) for r in _rows_m[1:] if any(r)]
                            save_history(_recs_m)
                            st.success(f"✅ Перенесено {len(_recs_m)} записей в Supabase")
                        except Exception as _me:
                            st.error(f"Ошибка миграции: {_me}")

        st.divider()
        st.markdown("**📗 Зеркало (Google Sheets)**")
        _hist_ws, _hist_sid = _get_history_sheet()
        if _hist_sid:
            st.caption(f"✅ Подключено: `{_hist_sid[:20]}…`")
            _hist_url = f"https://docs.google.com/spreadsheets/d/{_hist_sid}"
            st.caption(f"[Открыть таблицу]({_hist_url})")
        else:
            st.caption("⚠️ Sheets не настроен — резервная копия недоступна (основные данные в Supabase)")

        st.caption("Укажи ID существующей таблицы, к которой у сервисного аккаунта есть доступ:")
        _new_hist_sid = st.text_input(
            "ID таблицы-зеркала",
            value=_hist_sid or "",
            placeholder="1xd69v1vGx…",
            key="cfg_hist_sheet_id",
            label_visibility="collapsed",
        )
        if _new_hist_sid and _new_hist_sid != _hist_sid:
            if st.button("🔗 Подключить таблицу", key="connect_hist_sheet",
                         use_container_width=True):
                with st.spinner("Проверяю доступ…"):
                    try:
                        gc_test = _get_gspread_client()
                        if gc_test is None:
                            st.error("❌ Нет подключения к Google — проверь gcp_service_account в Secrets")
                        else:
                            ws_test = gc_test.open_by_key(_new_hist_sid).sheet1
                            # Ensure header row exists
                            existing_vals = ws_test.row_values(1)
                            if existing_vals != _HIST_HEADERS:
                                ws_test.update("A1", [_HIST_HEADERS])
                                ws_test.freeze(rows=1)
                            st.success("✅ Доступ подтверждён!")
                            st.info("Добавь в **Streamlit Secrets** и нажми Reboot app:")
                            st.code(f'history_sheet_id = "{_new_hist_sid}"', language="toml")
                    except Exception as _ce:
                        st.error(f"❌ Не удалось открыть таблицу: {_ce}")

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
                    "short_comment_threshold": short_comment_threshold,
                })
                st.success("✅ Сохранено!")

BASE_URL  = f"https://{subdomain}.t8s.ru/Api/V2"

tab4, tab5, tab1, tab2, tab3, tab6, tab7, tab8 = st.tabs(["📚 История", "📊 Статистика", "📥 Загрузить данные", "✉️ Сообщения", "📤 Отправить", "📢 Рассылка", "📑 Отчёты", "📅 Посещаемость"])


# Load history once per render — reused across Tab 4 and Tab 5
_all_history: list = load_history()

# ─────────────────────────────────────────────────────────────────────────────
#  REPORT HELPERS (used by tab4 history and tab7 reports)
# ─────────────────────────────────────────────────────────────────────────────

def _rp_grade(r):
    """Extract grade from Skills array."""
    skills = r.get("Skills") or []
    if not skills:
        return ""
    parts = []
    for sk in skills:
        name      = sk.get("SkillName") or sk.get("Name") or ""
        valid     = sk.get("ValidScore")
        score     = sk.get("Score")
        max_score = sk.get("MaxScore")
        mark = score if score is not None else valid
        if mark is None:
            continue
        def _fmt(v):
            return str(int(v)) if isinstance(v, float) and v == int(v) else str(v)
        val = f"{_fmt(mark)}/{_fmt(max_score)}" if max_score is not None else _fmt(mark)
        if name and not (len(skills) == 1 and name == "Общий"):
            val = f"{name}: {val}"
        parts.append(val)
    return " · ".join(parts) if parts else ""

def _strip_html(html: str) -> str:
    text = re.sub(r'<[^>]+>', ' ', html)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def _rp_comment(r):
    """Extract comment/report text. Prefers CommentHtml (stripped) when it contains more content."""
    plain = str(r.get("CommentText") or "").strip()
    html  = str(r.get("CommentHtml") or "").strip()
    if html:
        from_html = _strip_html(html)
        if len(from_html) > len(plain):
            return from_html
    if plain:
        return plain
    for _f in ("Comment", "Description", "Note", "Notes",
               "LessonDescription", "TopicDescription", "Text", "Content"):
        _v = r.get(_f)
        if _v and str(_v).strip():
            return str(_v).strip()
    return ""

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 1 — LOAD & ANALYSE
# ─────────────────────────────────────────────────────────────────────────────

with tab1:
    _t1c1, _t1c2, _t1c3 = st.columns([2, 2, 1])
    date_from_val = _t1c1.date_input("Период с", value=date.today() - timedelta(days=7), key="t1_date_from")
    date_to_val   = _t1c2.date_input("по",       value=date.today(),                     key="t1_date_to")
    DATE_FROM = date_from_val.strftime("%Y-%m-%d")
    DATE_TO   = date_to_val.strftime("%Y-%m-%d")
    with _t1c3:
        st.markdown('<div style="height:1.75rem"></div>', unsafe_allow_html=True)
        load_btn = st.button("🔄 Загрузить", type="primary", use_container_width=True)

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

                # Build resolved_keys: (teacher, date, "no_report") where report was actually written
                _resolved_keys = set()
                for (_eu_id_r, _d_r, _) in filled_reports:
                    for _t_r in eu_map.get(_eu_id_r, {}).get("teachers", []):
                        _resolved_keys.add((_t_r, _d_r, "no_report"))

                # Save to history
                if all_errors:
                    upsert_history(all_errors, DATE_FROM, DATE_TO, reviewer_name,
                                   resolved_keys=_resolved_keys)

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
        fc1, fc2, fc3, fc4, fc5 = st.columns([2, 2, 2, 2, 1])
        with fc1:
            all_hist_teachers = sorted({r["teacher"] for r in history_records})
            f_teacher = st.selectbox("Преподаватель", ["Все"] + all_hist_teachers, key="hf_teacher")
        with fc2:
            f_status = st.selectbox("Статус", ["Все"] + list(_STATUS_FILTER_MAP.keys()), key="hf_status")
        with fc3:
            f_error = st.selectbox(
                "Тип ошибки",
                ["Все", "Нет отчёта", "Нет комментария к пропуску", "Некорректный отчёт"],
                key="hf_error",
            )
        with fc4:
            f_role = st.selectbox(
                "Категория",
                ["Все", "Основной состав", "Стажёры"],
                key="hf_role",
            )
        with fc5:
            st.markdown('<div style="height:1.75rem"></div>', unsafe_allow_html=True)
            if st.button("🔄 Обновить", use_container_width=True, key="hist_refresh"):
                # Save dates from session_state (set by previous render of date inputs)
                _ss_from = st.session_state.get("hf_date_from")
                _ss_to   = st.session_state.get("hf_date_to")
                if _ss_from:
                    st.session_state["_hist_recheck_from"] = (
                        _ss_from.strftime("%Y-%m-%d") if hasattr(_ss_from, "strftime") else str(_ss_from)
                    )
                if _ss_to:
                    st.session_state["_hist_recheck_to"] = (
                        _ss_to.strftime("%Y-%m-%d") if hasattr(_ss_to, "strftime") else str(_ss_to)
                    )
                st.session_state["_hist_recheck"] = True
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
            "Некорректный отчёт":         "bad_report",
        }
        _all_dates = sorted({r["date"] for r in history_records if r.get("date")})
        if _all_dates:
            _hd_min = datetime.strptime(_all_dates[0],  "%Y-%m-%d").date()
            _hd_max = datetime.strptime(_all_dates[-1], "%Y-%m-%d").date()
        else:
            _hd_min = _hd_max = date.today()

        
        # Ensure session_state values for history date inputs are proper date objects
        for _k in ["hf_date_from", "hf_date_to"]:
            if _k in st.session_state and not isinstance(st.session_state[_k], date):
                del st.session_state[_k]
        fd1, fd2 = st.columns(2)
        with fd1:
            _h_def_from = max(_hd_min, min(_hd_max, date.today() - timedelta(days=date.today().weekday() + 7)))
            _h_date_from = st.date_input("Дата занятия с", value=_h_def_from,
                                         min_value=_hd_min, max_value=_hd_max, key="hf_date_from")
        with fd2:
            _h_def_to = max(_hd_min, min(_hd_max, date.today()))
            _h_date_to   = st.date_input("по",             value=_h_def_to,
                                         min_value=_hd_min, max_value=_hd_max, key="hf_date_to")
        _hdf = _h_date_from.strftime("%Y-%m-%d")
        _hdt = _h_date_to.strftime("%Y-%m-%d")

        # ── Перепроверка данных из HolliHop ───────────────────────────────────
        if st.session_state.pop("_hist_recheck", False):
            if not api_key:
                st.error("Укажи API ключ HolliHop в настройках!")
            else:
                _rc_from = st.session_state.pop("_hist_recheck_from", _hdf)
                _rc_to   = st.session_state.pop("_hist_recheck_to",   _hdt)
                _recheck_base = f"https://{subdomain}.t8s.ru/Api/V2"
                with st.status(f"Проверяю HolliHop за {_rc_from} — {_rc_to}…", expanded=True) as _rc_status:
                    st.write("👨‍🏫 Преподаватели…")
                    _rc_teachers = api_paginated(_recheck_base, api_key, "GetTeachers", "Teachers")
                    st.write("📚 Учебные единицы…")
                    _rc_eu = api_paginated(_recheck_base, api_key, "GetEdUnits", "EdUnits", params={
                        "types": "Group,MiniGroup,Individual",
                        "dateFrom": _rc_from, "dateTo": _rc_to, "queryDays": "true",
                    })
                    _rc_eu_map = {}
                    for _eu in _rc_eu:
                        _eu_id, _t_names = _eu["Id"], []
                        for _si in _eu.get("ScheduleItems", []):
                            for _tn in _si.get("Teachers", []):
                                if _tn not in _t_names:
                                    _t_names.append(_tn)
                        _rc_eu_map[_eu_id] = _t_names
                    st.write("📋 Посещаемость…")
                    _rc_eus = api_paginated(_recheck_base, api_key, "GetEdUnitStudents", "EdUnitStudents", params={
                        "dateFrom": _rc_from, "dateTo": _rc_to, "queryDays": "true",
                    })
                    st.write("📝 Отчёты…")
                    _rc_filled = {
                        (r["EdUnitId"], r["Date"], r["StudentClientId"])
                        for r in load_test_results(_recheck_base, api_key, _rc_from, _rc_to)
                    }
                    st.write("🔍 Анализирую…")
                    from collections import defaultdict as _rcdd
                    _rc_no_rep = _rcdd(lambda: _rcdd(int))
                    _rc_no_cmt = _rcdd(lambda: _rcdd(int))
                    _rc_nr_st  = _rcdd(lambda: _rcdd(list))
                    _rc_nc_st  = _rcdd(lambda: _rcdd(list))
                    for _eus in _rc_eus:
                        _eu_id2  = _eus.get("EdUnitId")
                        _sid     = _eus.get("StudentClientId")
                        _sname   = (_eus.get("Name") or _eus.get("StudentName")
                                    or _eus.get("ClientName") or str(_sid))
                        _tnames  = _rc_eu_map.get(_eu_id2, ["Неизвестный"])
                        for _day in _eus.get("Days", []):
                            _d = _day.get("Date", "")
                            if _d < _rc_from or _d > _rc_to:
                                continue
                            _is_pass = _day.get("Pass", False)
                            _desc    = (_day.get("Description") or "").strip()
                            if not _is_pass and (_eu_id2, _d, _sid) not in _rc_filled:
                                for _t in _tnames:
                                    _rc_no_rep[_t][_d] += 1
                                    _rc_nr_st[_t][_d].append(_sname)
                            if _is_pass and not _desc:
                                for _t in _tnames:
                                    _rc_no_cmt[_t][_d] += 1
                                    _rc_nc_st[_t][_d].append(_sname)
                    _rc_errors = []
                    for _t, _dc in _rc_no_rep.items():
                        for _d in sorted(_dc):
                            _rc_errors.append({"date": _d, "teacher": _t,
                                "error_type": "no_report", "error_description": "Нет отчёта",
                                "count": _dc[_d], "students": sorted(set(_rc_nr_st[_t][_d]))})
                    for _t, _dc in _rc_no_cmt.items():
                        for _d in sorted(_dc):
                            _rc_errors.append({"date": _d, "teacher": _t,
                                "error_type": "no_abs_comment",
                                "error_description": "Нет комментария к пропуску",
                                "count": _dc[_d], "students": sorted(set(_rc_nc_st[_t][_d]))})
                    # resolved_keys: teacher+date pairs where report was actually written
                    _rc_resolved = set()
                    for (_eu_r, _d_r, _) in _rc_filled:
                        for _t_r in _rc_eu_map.get(_eu_r, []):
                            _rc_resolved.add((_t_r, _d_r, "no_report"))
                    upsert_history(_rc_errors, _rc_from, _rc_to, reviewer_name,
                                   resolved_keys=_rc_resolved)
                    _rc_status.update(label=f"✅ Готово — найдено {len(_rc_errors)} ошибок за период", state="complete")
                st.rerun()

        # ── Применяем все фильтры одним проходом ──────────────────────────────
        _f_status_val = _STATUS_FILTER_MAP.get(f_status)
        _f_error_val  = _ERR_FILTER_MAP.get(f_error)
        all_filtered = [
            r for r in history_records
            if (f_teacher == "Все" or r["teacher"] == f_teacher)
            and (f_status  == "Все" or r["status"]     == _f_status_val)
            and (f_error   == "Все" or r["error_type"] == _f_error_val)
            and _hdf <= r.get("date", "") <= _hdt
            and (f_role == "Все"
                 or (f_role == "Стажёры"        and _teacher_info.get(r["teacher"], {}).get("is_intern", False))
                 or (f_role == "Основной состав" and not _teacher_info.get(r["teacher"], {}).get("is_intern", False)))
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
                                    # Для bad_report — динамически загружаем отчёты из API
                                    if rec.get("error_type") == "bad_report":
                                        _rp_cache_key = f"_hist_rp_{rec['id']}"
                                        if st.button("📋 Загрузить отчёты", key=f"load_rp_{rec['id']}"):
                                            with st.spinner("Загрузка из HolliHop…"):
                                                try:
                                                    _raw_rp = load_test_results(
                                                        BASE_URL, api_key,
                                                        rec["date"], rec["date"],
                                                    )
                                                    st.session_state[_rp_cache_key] = {
                                                        r.get("StudentName"): _rp_comment(r)
                                                        for r in _raw_rp
                                                        if r.get("StudentName") in _students
                                                    }
                                                except Exception as _e:
                                                    st.error(f"Ошибка: {_e}")
                                        _loaded_rp = st.session_state.get(_rp_cache_key, {})
                                        for _sn in _students:
                                            st.markdown(f"**{_sn}**")
                                            _rc = _loaded_rp.get(_sn)
                                            if _rc:
                                                st.caption(_rc)
                                            elif _loaded_rp:
                                                st.caption("_— отчёт не найден —_")
                                    else:
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
                            # Редактирование причины ошибки (всегда доступно)
                            _TYPE_OPTIONS = {
                                "Нет отчёта":                 "no_report",
                                "Нет комментария к пропуску": "no_abs_comment",
                                "Некорректный отчёт":         "bad_report",
                            }
                            _cur_type_label = next(
                                (k for k, v in _TYPE_OPTIONS.items() if v == rec.get("error_type")),
                                "Некорректный отчёт",
                            )
                            new_type_label = st.selectbox(
                                "Тип",
                                list(_TYPE_OPTIONS.keys()),
                                index=list(_TYPE_OPTIONS.keys()).index(_cur_type_label),
                                key=f"htype_{rid}",
                                label_visibility="collapsed",
                            )
                            if _TYPE_OPTIONS[new_type_label] != rec.get("error_type") and st.button(
                                "💾 Сохранить", key=f"hdsave_{rid}", use_container_width=True
                            ):
                                update_history_description(rid, rec.get("error_description", ""), _TYPE_OPTIONS[new_type_label])
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
    import plotly.express as px

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

        _fp1_col, _fp2_col, _ftype_col, _frole_col = st.columns([1.5, 1.5, 1.5, 1.5])

        # ── Date range filter (arbitrary only) ────────────────────────────────
        _hist_dates = sorted(r["date"] for r in stat_records)
        _min_d = datetime.strptime(_hist_dates[0], "%Y-%m-%d").date() if _hist_dates else date.today() - timedelta(days=30)
        _max_d = datetime.strptime(_hist_dates[-1], "%Y-%m-%d").date() if _hist_dates else date.today()
        with _fp1_col:
            _stat_def_from = max(_min_d, min(_max_d, date.today() - timedelta(days=date.today().weekday() + 7)))
            custom_from = st.date_input("С", value=_stat_def_from, key="stat_custom_from")
        with _fp2_col:
            _stat_def_to = max(_min_d, min(_max_d, date.today()))
            custom_to   = st.date_input("По", value=_stat_def_to, key="stat_custom_to")
        _cf = custom_from.strftime("%Y-%m-%d")
        _ct = custom_to.strftime("%Y-%m-%d")
        sr = [r for r in stat_records if _cf <= r["date"] <= _ct]

        # Пропущенные не учитываются в статистике совсем
        sr = [r for r in sr if r.get("status") != "skipped"]

        # ── Error type filter ─────────────────────────────────────────────────
        with _ftype_col:
            _err_type_opt = st.selectbox(
                "Тип ошибки",
                ["Все типы", "📋 Нет отчёта", "💬 Нет комментария"],
                key="stat_err_type",
            )
        if _err_type_opt == "📋 Нет отчёта":
            sr = [r for r in sr if r.get("error_type") == "no_report"]
        elif _err_type_opt == "💬 Нет комментария":
            sr = [r for r in sr if r.get("error_type") == "no_abs_comment"]

        # ── Intern / teacher role filter ──────────────────────────────────────
        with _frole_col:
            _role_opt = st.selectbox(
                "Категория",
                ["Все", "Только преподаватели", "Только стажёры"],
                key="stat_role_filter",
            )
        if _role_opt == "Только преподаватели":
            sr = [r for r in sr if not _teacher_info.get(r["teacher"], {}).get("is_intern", False)]
        elif _role_opt == "Только стажёры":
            sr = [r for r in sr if _teacher_info.get(r["teacher"], {}).get("is_intern", False)]

        # ── Single pass: group by teacher + count statuses ───────────────────
        from collections import Counter as _Counter, defaultdict as _dd
        _sr_by_teacher: dict = _dd(list)
        _sr_status_cnt = _Counter()
        _sr_teachers   = set()
        for _r in sr:
            _sr_by_teacher[_r["teacher"]].append(_r)
            _sr_status_cnt[_r["status"]] += 1
            _sr_teachers.add(_r["teacher"])

        # no_report считаем отдельно — независимо от фильтра по типу ошибки
        _nr_by_teacher: dict = _dd(list)
        _sr_no_filter = [
            r for r in stat_records
            if _cf <= r.get("date", "") <= _ct
            and r.get("status") != "skipped"
            and r.get("error_type") == "no_report"
            and (_role_opt == "Все"
                 or (_role_opt == "Только стажёры" and _teacher_info.get(r["teacher"], {}).get("is_intern", False))
                 or (_role_opt == "Только преподаватели" and not _teacher_info.get(r["teacher"], {}).get("is_intern", False)))
        ]
        for _r in _sr_no_filter:
            _nr_by_teacher[_r["teacher"]].append(_r)

        _processed_statuses = {"handled", "resolved"}
        _in_progress_statuses = {"message_sent", "pass_set"}
        total_rec      = len(sr)
        processed_n    = sum(_sr_status_cnt.get(s, 0) for s in _processed_statuses)
        in_progress_n  = sum(_sr_status_cnt.get(s, 0) for s in _in_progress_statuses)
        open_n         = _sr_status_cnt.get("open", 0)
        teachers_n     = len(_sr_teachers)
        not_written_total = sum(
            len(r.get("students") or []) or int(r.get("count") or 1)
            for r in _sr_no_filter
        )

        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("📋 Всего ошибок",      total_rec)
        m2.metric("✅ Исправлено",         processed_n,
                  help="Ошибки, где преподаватель реально написал отчёт (подтверждено перепроверкой) или закрыты вручную")
        m3.metric("⏳ В работе",           in_progress_n,
                  help="Отправлено напоминание или выставлен пропуск — ждём, пока преподаватель исправит")
        m4.metric("🔴 Не обработано",      open_n,
                  help="Ошибки, по которым ещё не предпринято никаких действий")
        m5.metric("✏️ Не написано",        not_written_total,
                  help="Всего учеников без отчёта за период (включая исправленные, кроме пропущенных)")
        m6.metric("👩‍🏫 Преподавателей",    teachers_n)

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
                not_written = sum(
                    len(r.get("students") or []) or int(r.get("count") or 1)
                    for r in _nr_by_teacher.get(teacher, [])
                )
                teacher_rows.append({
                    "Преподаватель":    teacher,
                    "📋 Не написано":   not_written,
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
                    "📋 Не написано":   st.column_config.NumberColumn(width="small", help="Кол-во учеников без отчёта (активные записи)"),
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

        # ── Charts ────────────────────────────────────────────────────────────
        _STATUS_COLORS = {
            "open":         "#ef4444",
            "handled":      "#22c55e",
            "resolved":     "#22c55e",
            "in_progress":  "#f59e0b",
            "message_sent": "#f59e0b",
            "pass_set":     "#f59e0b",
            "skipped":      "#6b7280",
        }
        _CHART_TEMPLATE = "plotly_dark"

        if sr:
            st.divider()

            # ── Row 1: bar chart by teacher + pie chart by status ─────────────
            _ch1, _ch2 = st.columns([2, 1])

            with _ch1:
                # Build data: top-N teachers by total errors, stacked by status
                _top_n = 20
                _teacher_totals = {t: len(recs) for t, recs in _sr_by_teacher.items()}
                _top_teachers = sorted(_teacher_totals, key=lambda t: _teacher_totals[t], reverse=True)[:_top_n]

                _bar_rows = []
                for _t in _top_teachers:
                    _t_cnt = _Counter(r["status"] for r in _sr_by_teacher[_t])
                    for _st, _n in _t_cnt.items():
                        if _n > 0:
                            _bar_rows.append({
                                "Преподаватель": _t,
                                "Статус": _st,
                                "Кол-во": _n,
                            })

                _STATUS_LABEL = {
                    "open": "Открыто",
                    "handled": "Исправлено",
                    "resolved": "Исправлено",
                    "message_sent": "В работе",
                    "pass_set": "В работе",
                    "skipped": "Пропущено",
                }
                if _bar_rows:
                    _df_bar = pd.DataFrame(_bar_rows)
                    _df_bar["Категория"] = _df_bar["Статус"].map(
                        lambda s: _STATUS_LABEL.get(s, s)
                    )
                    _df_bar_agg = (
                        _df_bar.groupby(["Преподаватель", "Категория"])["Кол-во"]
                        .sum()
                        .reset_index()
                    )
                    _cat_order = ["Открыто", "В работе", "Исправлено", "Пропущено"]
                    _cat_colors = {
                        "Открыто":    "#ef4444",
                        "В работе":   "#f59e0b",
                        "Исправлено": "#22c55e",
                        "Пропущено":  "#6b7280",
                    }
                    _fig_bar = px.bar(
                        _df_bar_agg,
                        x="Кол-во",
                        y="Преподаватель",
                        color="Категория",
                        orientation="h",
                        title="Ошибки по преподавателям",
                        template=_CHART_TEMPLATE,
                        color_discrete_map=_cat_colors,
                        category_orders={
                            "Категория": _cat_order,
                            "Преподаватель": _top_teachers[::-1],
                        },
                    )
                    _fig_bar.update_layout(
                        margin=dict(l=0, r=0, t=40, b=0),
                        legend_title_text="",
                        xaxis_showgrid=False,
                        yaxis_showgrid=False,
                        plot_bgcolor="rgba(0,0,0,0)",
                        paper_bgcolor="rgba(0,0,0,0)",
                        height=max(300, len(_top_teachers) * 28),
                    )
                    st.plotly_chart(_fig_bar, use_container_width=True)

            with _ch2:
                # Pie/donut: distribution by status category
                _pie_data = {
                    "Исправлено": processed_n,
                    "В работе":   in_progress_n,
                    "Открыто":    open_n,
                    "Пропущено":  _sr_status_cnt.get("skipped", 0),
                }
                _pie_data = {k: v for k, v in _pie_data.items() if v > 0}
                if _pie_data:
                    _pie_colors = {
                        "Исправлено": "#22c55e",
                        "В работе":   "#f59e0b",
                        "Открыто":    "#ef4444",
                        "Пропущено":  "#6b7280",
                    }
                    _fig_pie = px.pie(
                        names=list(_pie_data.keys()),
                        values=list(_pie_data.values()),
                        title="Распределение по статусам",
                        template=_CHART_TEMPLATE,
                        color=list(_pie_data.keys()),
                        color_discrete_map=_pie_colors,
                        hole=0.45,
                    )
                    _fig_pie.update_traces(textposition="inside", textinfo="percent+label")
                    _fig_pie.update_layout(
                        margin=dict(l=0, r=0, t=40, b=0),
                        showlegend=False,
                        plot_bgcolor="rgba(0,0,0,0)",
                        paper_bgcolor="rgba(0,0,0,0)",
                        height=300,
                    )
                    st.plotly_chart(_fig_pie, use_container_width=True)

            # ── Row 2: errors by lesson date ──────────────────────────────────
            _date_cnt: dict = _dd(lambda: _Counter())
            for _r in sr:
                _date_cnt[_r["date"]][_r["status"]] += 1

            if _date_cnt:
                _date_rows = []
                for _d in sorted(_date_cnt.keys()):
                    _dc = _date_cnt[_d]
                    _date_rows.append({
                        "Дата":       _d,
                        "Открыто":    _dc.get("open", 0),
                        "В работе":   _dc.get("message_sent", 0) + _dc.get("pass_set", 0),
                        "Исправлено": _dc.get("handled", 0) + _dc.get("resolved", 0),
                        "Пропущено":  _dc.get("skipped", 0),
                    })
                _df_dates = pd.DataFrame(_date_rows)
                _df_dates_melted = _df_dates.melt(
                    id_vars="Дата",
                    value_vars=["Открыто", "В работе", "Исправлено", "Пропущено"],
                    var_name="Категория",
                    value_name="Кол-во",
                )
                _df_dates_melted = _df_dates_melted[_df_dates_melted["Кол-во"] > 0]
                if not _df_dates_melted.empty:
                    _fig_line = px.bar(
                        _df_dates_melted,
                        x="Дата",
                        y="Кол-во",
                        color="Категория",
                        title="Ошибки по датам занятий",
                        template=_CHART_TEMPLATE,
                        color_discrete_map={
                            "Открыто":    "#ef4444",
                            "В работе":   "#f59e0b",
                            "Исправлено": "#22c55e",
                            "Пропущено":  "#6b7280",
                        },
                        category_orders={"Категория": ["Открыто", "В работе", "Исправлено", "Пропущено"]},
                    )
                    _fig_line.update_layout(
                        margin=dict(l=0, r=0, t=40, b=0),
                        legend_title_text="",
                        xaxis_showgrid=False,
                        yaxis_showgrid=False,
                        bargap=0.15,
                        plot_bgcolor="rgba(0,0,0,0)",
                        paper_bgcolor="rgba(0,0,0,0)",
                        height=320,
                    )
                    st.plotly_chart(_fig_line, use_container_width=True)

        # ── Per-period dynamics (always shown when multiple periods exist) ────
        if len(all_stat_periods) > 1:
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


# ─────────────────────────────────────────────────────────────────────────────
#  TAB 7 — ОТЧЁТЫ
# ─────────────────────────────────────────────────────────────────────────────


def _rp_quality(comment: str, grade: str, threshold: int = 400):
    """Return (icon, label, css_color) quality indicator."""
    if not comment:
        return "🔴", "Нет комментария", "#f87171"
    if len(comment) < threshold:
        return "🟡", "Короткий комментарий", "#fbbf24"
    return "🟢", "OK", "#34d399"

with tab7:
    st.subheader("📑 Отчёты преподавателей")
    _rp_threshold = int(st.session_state.get("cfg_short_threshold", cfg.get("short_comment_threshold", 400)))

    _rp_c1, _rp_c2, _rp_c3 = st.columns([2, 2, 1])
    _rp_date_from = _rp_c1.date_input("Период с", value=date.today() - timedelta(days=7), key="rp_date_from")
    _rp_date_to   = _rp_c2.date_input("по",       value=date.today(),                     key="rp_date_to")
    with _rp_c3:
        st.markdown('<div style="height:1.75rem"></div>', unsafe_allow_html=True)
        _rp_load_btn = st.button("🔍 Загрузить", type="primary", use_container_width=True, key="rp_load")

    _rp_from_str = _rp_date_from.strftime("%Y-%m-%d")
    _rp_to_str   = _rp_date_to.strftime("%Y-%m-%d")

    _rp_cache_key = f"rp_data_{_rp_from_str}_{_rp_to_str}"

    if _rp_load_btn:
        if not api_key:
            st.error("Укажи API ключ HolliHop в настройках ⚙️")
        elif _rp_from_str > _rp_to_str:
            st.error("Дата «С» должна быть раньше даты «По»!")
        else:
            with st.status(f"Загружаю отчёты за {_rp_from_str} — {_rp_to_str}…", expanded=True) as _rp_status:
                st.write("📚 Учебные единицы…")
                _rp_eu_list = api_paginated(BASE_URL, api_key, "GetEdUnits", "EdUnits", params={
                    "types": "Group,MiniGroup,Individual",
                    "dateFrom": _rp_from_str, "dateTo": _rp_to_str, "queryDays": "true",
                })
                _rp_eu_map = {}
                for _eu in _rp_eu_list:
                    _eu_id, _t_names = _eu["Id"], []
                    for _si in _eu.get("ScheduleItems", []):
                        for _tn in _si.get("Teachers", []):
                            if _tn not in _t_names:
                                _t_names.append(_tn)
                    _rp_eu_map[_eu_id] = {"name": _eu.get("Name", "—"), "teachers": _t_names}
                st.write(f"   → {len(_rp_eu_list)} учебных единиц")

                st.write("📝 Отчёты…")
                _rp_raw = load_test_results(BASE_URL, api_key, _rp_from_str, _rp_to_str)
                st.write(f"   → {len(_rp_raw)} записей")

                _rp_rows = []
                for _r in _rp_raw:
                    _eu_info = _rp_eu_map.get(_r.get("EdUnitId"), {})
                    _grade   = _rp_grade(_r)
                    _comment = _rp_comment(_r)
                    _rp_rows.append({
                        "date":        _r.get("Date", ""),
                        "teacher":     ", ".join(_eu_info.get("teachers", [])) or "—",
                        "subject":     (_r.get("EdUnitName") or _eu_info.get("name", "—")),
                        "student":     (_r.get("StudentName") or _r.get("ClientName")
                                        or str(_r.get("StudentClientId", "—"))),
                        "test_type":   (_r.get("TestTypeName") or ""),
                        "grade":       _grade,
                        "comment":     _comment,
                        "_skills":     _r.get("Skills") or [],
                    })

                st.session_state[_rp_cache_key] = _rp_rows
                _rp_status.update(label=f"✅ Загружено {len(_rp_rows)} записей", state="complete")

    _rp_data = st.session_state.get(_rp_cache_key, [])

    if not _rp_data and not _rp_load_btn:
        st.info("Выбери период и нажми «Загрузить».")
    elif _rp_data:
        # ── Summary bar ───────────────────────────────────────────────────────
        _rp_no_comment = sum(1 for r in _rp_data if not r["comment"])
        _rp_short      = sum(1 for r in _rp_data if r["comment"] and len(r["comment"]) < _rp_threshold)
        st.markdown(
            f'<div class="info-bar">'
            f'Всего: <b>{len(_rp_data)}</b> &nbsp;·&nbsp; '
            f'🔴 Без комментария: <b>{_rp_no_comment}</b> &nbsp;·&nbsp; '
            f'🟡 Короткий: <b>{_rp_short}</b>'
            f'</div>', unsafe_allow_html=True)

        # ── Фильтры ───────────────────────────────────────────────────────────
        _rp_f1, _rp_f2, _rp_f3 = st.columns([2, 2, 2])
        _rp_teachers_all = sorted({r["teacher"] for r in _rp_data})
        _rp_tf = _rp_f1.selectbox("Преподаватель", ["Все"] + _rp_teachers_all, key="rp_f_teacher")
        _rp_qf = _rp_f2.selectbox("Качество отчёта",
                                   ["Все", "🔴 Нет комментария", "🟡 Короткий", "🟢 OK"],
                                   key="rp_f_quality")
        _rp_search = _rp_f3.text_input("Поиск по ученику", key="rp_search", placeholder="Введи ФИО…")

        # Apply filters
        _rp_filtered = _rp_data[:]
        if _rp_tf != "Все":
            _rp_filtered = [r for r in _rp_filtered if r["teacher"] == _rp_tf]
        if _rp_qf == "🔴 Нет комментария":
            _rp_filtered = [r for r in _rp_filtered if not r["comment"]]
        elif _rp_qf == "🟡 Короткий":
            _rp_filtered = [r for r in _rp_filtered if r["comment"] and len(r["comment"]) < _rp_threshold]
        elif _rp_qf == "🟢 OK":
            _rp_filtered = [r for r in _rp_filtered if r["comment"] and len(r["comment"]) >= 25]
        if _rp_search:
            _rp_filtered = [r for r in _rp_filtered
                            if _rp_search.lower() in r["student"].lower()]

        st.caption(f"Показано: {len(_rp_filtered)} записей")

        # ── Группировка по преподавателю ──────────────────────────────────────
        from collections import defaultdict as _rpdd
        _rp_by_teacher = _rpdd(list)
        for _r in sorted(_rp_filtered, key=lambda x: (x["teacher"], x["date"])):
            _rp_by_teacher[_r["teacher"]].append(_r)

        for _teacher, _t_records in _rp_by_teacher.items():
            _t_bad  = sum(1 for r in _t_records if not r["comment"])
            _t_warn = sum(1 for r in _t_records if r["comment"] and len(r["comment"]) < _rp_threshold)
            _t_badge = (f"  🔴 {_t_bad} без комментария" if _t_bad else "") + \
                       (f"  🟡 {_t_warn} коротких" if _t_warn else "")
            _t_label = f"**{_teacher}** — {len(_t_records)} отчётов{_t_badge}"
            with st.expander(_t_label, expanded=True):
                # Кнопка жалобы на всего преподавателя
                _tc_btn, _tc_spacer = st.columns([3, 9])
                if _tc_btn.button(
                    "📋 Жалоба на все отчёты",
                    key=f"flag_teacher_{_teacher}",
                    use_container_width=True,
                ):
                    _flag_teacher_dialog(_teacher, _t_records)

                # Group by date+subject within teacher
                _rp_by_lesson = _rpdd(list)
                for _r in _t_records:
                    _rp_by_lesson[(_r["date"], _r["subject"])].append(_r)

                for (_d, _subj), _lesson_recs in sorted(_rp_by_lesson.items()):
                    try:
                        _d_label = datetime.strptime(_d, "%Y-%m-%d").strftime("%d.%m.%Y")
                    except Exception:
                        _d_label = _d
                    _l_bad  = sum(1 for r in _lesson_recs if not r["comment"])
                    _l_warn = sum(1 for r in _lesson_recs if r["comment"] and len(r["comment"]) < _rp_threshold)
                    _l_color = "#f87171" if _l_bad else "#fbbf24" if _l_warn else "#34d399"
                    _test_types = sorted({r["test_type"] for r in _lesson_recs if r["test_type"]})
                    _tt_str = f' &nbsp;·&nbsp; <i style="color:#a78bfa;">{", ".join(_test_types)}</i>' if _test_types else ""
                    st.markdown(
                        f'<div style="margin:6px 0 4px 0; padding:6px 12px; '
                        f'border-left:3px solid {_l_color}; '
                        f'background:rgba(255,255,255,0.03); border-radius:0 8px 8px 0;">'
                        f'<b>{_d_label}</b> &nbsp;·&nbsp; {_subj}{_tt_str} '
                        f'<span style="color:rgba(255,255,255,0.45); font-size:0.82rem;">'
                        f'({len(_lesson_recs)} уч.)</span></div>',
                        unsafe_allow_html=True,
                    )
                    for _ri, _rec in enumerate(_lesson_recs):
                        _icon, _, _color = _rp_quality(_rec["comment"], _rec["grade"], _rp_threshold)
                        _grade_str = _rec["grade"] if _rec["grade"] else ""
                        _comment_str = _rec["comment"] if _rec["comment"] else "*— комментарий отсутствует —*"
                        _row_col, _btn_col = st.columns([20, 1])
                        with _row_col:
                            st.markdown(
                                f'<div style="display:flex; gap:10px; align-items:flex-start; '
                                f'padding:5px 8px 5px 16px; border-bottom:1px solid rgba(255,255,255,0.05);">'
                                f'<span style="min-width:18px;">{_icon}</span>'
                                f'<span style="min-width:220px; color:rgba(255,255,255,0.85);">{_rec["student"]}</span>'
                                f'<span style="min-width:60px; color:#c4b5fd;">{_grade_str}</span>'
                                f'<span style="color:rgba(255,255,255,0.6); font-size:0.88rem;">{_comment_str}</span>'
                                f'</div>',
                                unsafe_allow_html=True,
                            )
                        with _btn_col:
                            _flag_key = f"flag_{_rec['teacher']}_{_rec['date']}_{_rec['student']}_{_ri}"
                            if st.button("⚠️", key=_flag_key, help="Отметить как некорректный"):
                                _flag_report_dialog(_rec)

        # ── Скачать отчёты для проверки нейронкой (Markdown) ─────────────────
        if _rp_data:
            def _md_cell(s: str) -> str:
                return str(s or "").replace("|", "｜").replace("\n", " ").replace("\r", "")
            _dl_lines = [
                "| Преподаватель | Дата | Предмет | Тип | Ученик | Оценка | Комментарий |",
                "|---|---|---|---|---|---|---|",
            ]
            for _rec in sorted(_rp_data, key=lambda x: (x["teacher"], x["date"], x["subject"], x["student"])):
                try:
                    _dl_d = datetime.strptime(_rec["date"], "%Y-%m-%d").strftime("%d.%m.%Y")
                except Exception:
                    _dl_d = _rec["date"]
                _dl_lines.append(
                    f"| {_md_cell(_rec['teacher'])} "
                    f"| {_dl_d} "
                    f"| {_md_cell(_rec['subject'])} "
                    f"| {_md_cell(_rec['test_type'])} "
                    f"| {_md_cell(_rec['student'])} "
                    f"| {_md_cell(_rec['grade'])} "
                    f"| {_md_cell(_rec['comment'])} |"
                )
            _dl_text = "\n".join(_dl_lines)
            st.download_button(
                label="⬇ Скачать отчёты (.md)",
                data=_dl_text,
                file_name=f"reports_{_rp_from_str}_{_rp_to_str}.md",
                mime="text/markdown",
            )

# ── Tab 8: Посещаемость ────────────────────────────────────────────────────
with tab8:
    st.subheader("📅 Посещаемость")

    _at_c1, _at_c2 = st.columns([2, 2])
    _at_today = date.today()
    _at_week_start = _at_today - timedelta(days=_at_today.weekday())
    _at_date_from = _at_c1.date_input(
        "Период с",
        value=_at_week_start,
        key="at_date_from",
    )
    _at_date_to = _at_c2.date_input(
        "по",
        value=_at_today,
        key="at_date_to",
    )

    # Список преподавателей из кэша (загружается при первом открытии)
    if "at_teachers_list" not in st.session_state:
        st.session_state["at_teachers_list"] = []
    _at_tc1, _at_tc2 = st.columns([3, 1])
    with _at_tc1:
        _at_teacher_filter = st.selectbox(
            "Преподаватель",
            options=["Все преподаватели"] + sorted(st.session_state["at_teachers_list"]),
            key="at_teacher_filter",
        )
    with _at_tc2:
        st.markdown('<div style="height:1.75rem"></div>', unsafe_allow_html=True)
        _at_load = st.button("🔍 Загрузить", type="primary", use_container_width=True, key="at_load")

    if _at_load:
        _at_from_str = _at_date_from.strftime("%Y-%m-%d")
        _at_to_str   = _at_date_to.strftime("%Y-%m-%d")

        with st.spinner("Загрузка данных из HolliHop…"):
            _at_ed_units = api_paginated(BASE_URL, api_key, "GetEdUnits", "EdUnits", params={
                "types": "Group,MiniGroup,Individual",
                "dateFrom": _at_from_str, "dateTo": _at_to_str, "queryDays": "true",
            })
            # eu_map: преподаватели + отменённые даты по каждому EdUnit
            _at_eu_map = {}
            _at_cancelled_dates = defaultdict(set)  # eu_id → {dates}
            for _at_eu in _at_ed_units:
                _at_eu_id, _at_t_names = _at_eu["Id"], []
                for _at_si in _at_eu.get("ScheduleItems", []):
                    for _at_tn in _at_si.get("Teachers", []):
                        if _at_tn not in _at_t_names:
                            _at_t_names.append(_at_tn)
                _at_eu_map[_at_eu_id] = {
                    "name":     _at_eu.get("Name", ""),
                    "teachers": _at_t_names,
                }
                # Ищем отменённые дни в расписании EdUnit
                for _at_eday in _at_eu.get("Days", []):
                    _at_eday_d = _at_eday.get("Date", "")
                    if not _at_eday_d:
                        continue
                    if (
                        _at_eday.get("IsCancelled")
                        or _at_eday.get("Cancelled")
                        or str(_at_eday.get("Status", "")).lower() in ("cancelled", "cancel", "отмена")
                        or "отмен" in str(_at_eday.get("Description", "")).lower()
                        or "отмен" in str(_at_eday.get("Comment", "")).lower()
                    ):
                        _at_cancelled_dates[_at_eu_id].add(_at_eday_d)

            _at_eus = api_paginated(BASE_URL, api_key, "GetEdUnitStudents", "EdUnitStudents", params={
                "dateFrom": _at_from_str, "dateTo": _at_to_str, "queryDays": "true",
            })
            _at_test_results = load_test_results(BASE_URL, api_key, _at_from_str, _at_to_str)

        # Обновляем список преподавателей для селектора
        _at_all_teachers = sorted({
            t
            for v in _at_eu_map.values()
            for t in v["teachers"]
        })
        st.session_state["at_teachers_list"] = _at_all_teachers

        # Фильтр по преподавателю
        _at_chosen_teacher = st.session_state.get("at_teacher_filter", "Все преподаватели")

        # Урок состоялся если есть хоть один отчёт по EdUnit+Date (любой студент)
        # Отчёт (GetEdUnitTestResults) ≠ Комментарий (Day.Description)
        _at_lesson_happened = {
            (r.get("EdUnitId"), r.get("Date"))
            for r in _at_test_results
        }

        # ── Классифицируем Accepted=False записи ───────────────────────────
        # Обрабатываем только два случая:
        #   1. Pass=True + Accepted=False → пропуск стоит, подтверждаем
        #   2. Pass=False + Accepted=False + урок был → присутствовал
        # Всё остальное (нет отчёта → нет посещаемости) = нормально, не трогаем
        _at_auto_present   = []  # pass=False → присутствовал
        _at_auto_cancelled = []  # pass=True  → подтверждаем пропуск
        _at_unset          = []  # нет отчёта, нет пропуска → показываем, просим проставить

        for _at_rec in _at_eus:
            _at_eu_id     = _at_rec.get("EdUnitId") or _at_rec.get("Id")
            _at_client_id = _at_rec.get("StudentClientId")
            _at_sname     = _at_rec.get("StudentName") or _at_rec.get("ClientName") or "—"
            _at_group     = _at_rec.get("EdUnitName") or _at_eu_map.get(_at_eu_id, {}).get("name", "")
            _at_teachers  = _at_eu_map.get(_at_eu_id, {}).get("teachers", [])

            # Фильтр по выбранному преподавателю
            if _at_chosen_teacher != "Все преподаватели" and _at_chosen_teacher not in _at_teachers:
                continue

            for _at_day in _at_rec.get("Days", []):
                _at_d = _at_day.get("Date", "")
                if not _at_d or _at_d < _at_from_str or _at_d > _at_to_str:
                    continue
                if _at_day.get("Accepted"):
                    continue  # посещаемость уже подтверждена — пропускаем

                _at_entry = {
                    "edUnitId":        _at_eu_id,
                    "studentClientId": _at_client_id,
                    "date":            _at_d,
                    "student":         _at_sname,
                    "group":           _at_group,
                    "teachers":        _at_teachers,
                    "existing_desc":   (_at_day.get("Description") or "").strip(),
                    "minutes":         _at_day.get("Minutes"),
                }

                if _at_day.get("Pass") is True:
                    # Пропуск уже выставлен (Pass=True) но Accepted=False → подтверждаем
                    _at_auto_cancelled.append(_at_entry)
                elif (_at_eu_id, _at_d) in _at_lesson_happened:
                    # Урок состоялся (есть отчёт по группе) → студент присутствовал
                    _at_auto_present.append(_at_entry)
                else:
                    # Нет отчёта, нет пропуска → посещаемость не проставлена, показываем
                    _at_unset.append(_at_entry)

        # Сохраняем в session_state чтобы не перезагружать при нажатии кнопки
        st.session_state["at_auto_present"]   = _at_auto_present
        st.session_state["at_auto_cancelled"] = _at_auto_cancelled
        st.session_state["at_unset"]          = _at_unset
        st.session_state["at_period"]         = (_at_from_str, _at_to_str)

    # ── Показываем результат (из session_state) ──────────────────────────────
    _at_auto_present   = st.session_state.get("at_auto_present")
    _at_auto_cancelled = st.session_state.get("at_auto_cancelled")
    _at_unset          = st.session_state.get("at_unset")

    if _at_auto_present is not None:
        st.markdown("---")
        _at_p_n  = len(_at_auto_present)
        _at_c_n  = len(_at_auto_cancelled)
        _at_un_n = len(_at_unset) if _at_unset else 0

        # Итоговые метрики
        _at_m1, _at_m2, _at_m3 = st.columns(3)
        _at_m1.metric("✅ Урок был → присутствовал", f"{_at_p_n} уч.",
                       help="Pass=false — урок состоялся (есть отчёт по группе)")
        _at_m2.metric("🚫 Пропуск → подтвердить",   f"{_at_c_n} уч.",
                       help="Pass=true — пропуск уже стоит, убираем знак вопроса")
        _at_m3.metric("❓ Не проставлено",           f"{_at_un_n} уч.",
                       help="Нет ни пропуска, ни отчёта — нужно уточнить у преподавателя")

        _at_all_to_mark = _at_auto_present + _at_auto_cancelled
        if _at_all_to_mark:
            st.markdown("")
            if st.button(
                f"🚀 Проставить посещаемость для {len(_at_all_to_mark)} учеников",
                type="primary",
                key="at_apply",
            ):
                _at_ok, _at_err = 0, 0
                _at_errs = []
                _at_batch = []

                def _at_make_item(e, pass_val, desc=None):
                    item = {
                        "edUnitId":        e["edUnitId"],
                        "studentClientId": e["studentClientId"],
                        "date":            e["date"],
                        "pass":            pass_val,
                        "description":     desc if desc is not None else e["existing_desc"],
                    }
                    if e.get("minutes") is not None:
                        item["minutes"] = e["minutes"]
                    return item

                for _at_e in _at_auto_present:
                    _at_batch.append(_at_make_item(_at_e, False))   # присутствовал
                for _at_e in _at_auto_cancelled:
                    _at_batch.append(_at_make_item(_at_e, True))    # пропуск (уже стоял)

                with st.expander("🛠 Отладка: пример запроса (первая запись)", expanded=False):
                    if _at_batch:
                        st.json(_at_batch[0])

                _at_api_responses = []
                with st.spinner(f"Отправляем {len(_at_batch)} записей в HolliHop…"):
                    _at_url = f"{BASE_URL}/SetStudentPasses"
                    for _i in range(0, len(_at_batch), 100):
                        _at_chunk = _at_batch[_i:_i + 100]
                        try:
                            _at_resp = requests.post(
                                _at_url,
                                params={"authkey": api_key},
                                json=_at_chunk,
                                timeout=60,
                            )
                            _at_api_responses.append(
                                f"Пакет {_i//100+1}: HTTP {_at_resp.status_code} | {_at_resp.text[:500]}"
                            )
                            if _at_resp.status_code in (200, 204):
                                _at_ok += len(_at_chunk)
                            else:
                                _at_err += len(_at_chunk)
                                _at_errs.append(f"HTTP {_at_resp.status_code}: {_at_resp.text[:300]}")
                        except Exception as _at_ex:
                            _at_err += len(_at_chunk)
                            _at_errs.append(str(_at_ex))
                            _at_api_responses.append(f"Пакет {_i//100+1}: Exception — {_at_ex}")

                with st.expander("🛠 Ответ сервера HolliHop", expanded=True):
                    for _at_ar in _at_api_responses:
                        st.code(_at_ar)

                if _at_err == 0:
                    st.success(f"✅ Запрос принят для **{_at_ok}** учеников — проверь HolliHop через минуту")
                    st.session_state["at_auto_present"]   = []
                    st.session_state["at_auto_cancelled"] = []
                else:
                    st.warning(f"Готово: {_at_ok} ✅  /  {_at_err} ❌")
                    for _at_em in _at_errs:
                        st.caption(f"⚠️ {_at_em}")
        else:
            st.success("✅ Нет записей для проставления — всё уже отмечено!")

        # ── Не проставленная посещаемость + рассылка в Пачку ────────────────
        if _at_unset:
            st.markdown("---")
            st.markdown(f"### ❓ Не проставлена посещаемость — {_at_un_n} уч.")
            st.caption("Нет ни пропуска, ни отчёта. Нужно уточнить у преподавателя.")

            # Группируем по преподавателю → дата → ученики
            _at_un_by_t = defaultdict(lambda: defaultdict(list))
            for _at_ur in _at_unset:
                for _at_ut in (_at_ur["teachers"] or ["—"]):
                    _at_un_by_t[_at_ut][_at_ur["date"]].append(_at_ur["student"])

            # Показываем список
            for _at_ut in sorted(_at_un_by_t):
                _at_ut_cnt = sum(len(v) for v in _at_un_by_t[_at_ut].values())
                with st.expander(f"👤 {_at_ut}  —  {_at_ut_cnt} уч.", expanded=True):
                    for _at_ud in sorted(_at_un_by_t[_at_ut]):
                        try:
                            _at_ud_fmt = datetime.strptime(_at_ud, "%Y-%m-%d").strftime("%d.%m.%Y")
                        except Exception:
                            _at_ud_fmt = _at_ud
                        _at_ust = ", ".join(_at_un_by_t[_at_ut][_at_ud])
                        st.markdown(f"**{_at_ud_fmt}:** {_at_ust}")

            # ── Рассылка в Пачку ────────────────────────────────────────────
            st.markdown("---")
            st.markdown("#### 📨 Рассылка в Пачку")

            # Шаблон сообщения
            _at_default_msg = (
                "Добрый день! Просьба проставить посещаемость в системе HolliHop за следующие даты:\n\n"
                "{details}\n\n"
                "Спасибо!"
            )
            _at_msg_tpl = st.text_area(
                "Шаблон сообщения (`{details}` → даты и ученики, `{teacher}` → имя преподавателя)",
                value=_at_default_msg,
                height=160,
                key="at_msg_tpl",
            )

            # Предпросмотр для каждого преподавателя
            with st.expander("👁 Предпросмотр сообщений", expanded=False):
                for _at_ut in sorted(_at_un_by_t):
                    _at_details_lines = []
                    for _at_ud in sorted(_at_un_by_t[_at_ut]):
                        try:
                            _at_ud_fmt = datetime.strptime(_at_ud, "%Y-%m-%d").strftime("%d.%m.%Y")
                        except Exception:
                            _at_ud_fmt = _at_ud
                        _at_ust = ", ".join(_at_un_by_t[_at_ut][_at_ud])
                        _at_details_lines.append(f"• {_at_ud_fmt}: {_at_ust}")
                    _at_preview = _at_msg_tpl.replace("{details}", "\n".join(_at_details_lines)).replace("{teacher}", _at_ut.split()[0])
                    st.markdown(f"**{_at_ut}:**")
                    st.code(_at_preview, language=None)

            if st.button("📨 Отправить всем в Пачку", type="primary", key="at_send_pachca"):
                if not pachca_token:
                    st.error("Токен Пачки не настроен — откройте ⚙️")
                else:
                    _at_pachca = PachcaAPI(pachca_token)
                    with st.spinner("Загружаю пользователей Пачки…"):
                        _at_pachca.load_users()
                    _at_sent, _at_not_found = [], []
                    for _at_ut in sorted(_at_un_by_t):
                        _at_details_lines = []
                        for _at_ud in sorted(_at_un_by_t[_at_ut]):
                            try:
                                _at_ud_fmt = datetime.strptime(_at_ud, "%Y-%m-%d").strftime("%d.%m.%Y")
                            except Exception:
                                _at_ud_fmt = _at_ud
                            _at_ust = ", ".join(_at_un_by_t[_at_ut][_at_ud])
                            _at_details_lines.append(f"• {_at_ud_fmt}: {_at_ust}")
                        _at_msg = _at_msg_tpl.replace("{details}", "\n".join(_at_details_lines)).replace("{teacher}", _at_ut.split()[0])
                        _at_user = _at_pachca.find_user(_at_ut)
                        if not _at_user:
                            _at_not_found.append(_at_ut)
                            continue
                        try:
                            _at_pachca.send_dm(_at_user["id"], _at_msg)
                            _at_sent.append(_at_ut)
                        except Exception as _at_ex:
                            _at_not_found.append(f"{_at_ut} (ошибка: {_at_ex})")
                    if _at_sent:
                        st.success(f"✅ Отправлено: {', '.join(_at_sent)}")
                    if _at_not_found:
                        st.warning(f"⚠️ Не найдены в Пачке: {', '.join(_at_not_found)}")

        # ── Исправить ошибочно выставленные пропуски ────────────────────────
        st.markdown("---")
        st.markdown("### 🔧 Исправить ошибочные пропуски")
        st.caption("Ищет тех, кому стоит pass=true (пропуск), но при этом есть отчёт → исправляет на pass=false (присутствовал)")

        if st.button("🔍 Найти ошибочные пропуски", key="at_fix_find"):
            _at_ff_str = _at_date_from.strftime("%Y-%m-%d")
            _at_ft_str = _at_date_to.strftime("%Y-%m-%d")
            with st.spinner("Загружаю данные…"):
                _at_fix_eus = api_paginated(BASE_URL, api_key, "GetEdUnitStudents", "EdUnitStudents", params={
                    "dateFrom": _at_ff_str, "dateTo": _at_ft_str, "queryDays": "true",
                })
                _at_fix_ed_units = api_paginated(BASE_URL, api_key, "GetEdUnits", "EdUnits", params={
                    "types": "Group,MiniGroup,Individual",
                    "dateFrom": _at_ff_str, "dateTo": _at_ft_str, "queryDays": "true",
                })
                _at_fix_eu_map = {}
                for _at_feu in _at_fix_ed_units:
                    _at_feu_id, _at_ft_names = _at_feu["Id"], []
                    for _at_fsi in _at_feu.get("ScheduleItems", []):
                        for _at_ftn in _at_fsi.get("Teachers", []):
                            if _at_ftn not in _at_ft_names:
                                _at_ft_names.append(_at_ftn)
                    _at_fix_eu_map[_at_feu_id] = {
                        "name": _at_feu.get("Name", ""),
                        "teachers": _at_ft_names,
                    }
                _at_fix_results = load_test_results(BASE_URL, api_key, _at_ff_str, _at_ft_str)
            _at_fix_has_report = {
                (r.get("EdUnitId"), r.get("StudentClientId"), r.get("Date"))
                for r in _at_fix_results
            }
            # Ищем: Pass=True (пропуск) И есть отчёт → ошибка
            _at_fix_list = []
            for _at_frec in _at_fix_eus:
                _at_feu_id    = _at_frec.get("EdUnitId") or _at_frec.get("Id")
                _at_fclient   = _at_frec.get("StudentClientId")
                _at_fsname    = _at_frec.get("StudentName") or "—"
                _at_fgroup    = _at_frec.get("EdUnitName") or _at_fix_eu_map.get(_at_feu_id, {}).get("name", "")
                _at_fteachers = _at_fix_eu_map.get(_at_feu_id, {}).get("teachers", [])
                for _at_fday in _at_frec.get("Days", []):
                    _at_fd = _at_fday.get("Date", "")
                    if not _at_fd or _at_fd < _at_ff_str or _at_fd > _at_ft_str:
                        continue
                    if _at_fday.get("Pass") is True and (_at_feu_id, _at_fclient, _at_fd) in _at_fix_has_report:
                        _at_fix_list.append({
                            "edUnitId":        _at_feu_id,
                            "studentClientId": _at_fclient,
                            "date":            _at_fd,
                            "student":         _at_fsname,
                            "group":           _at_fgroup,
                            "teachers":        _at_fteachers,
                            "existing_desc":   (_at_fday.get("Description") or "").strip(),
                            "minutes":         _at_fday.get("Minutes"),
                        })
            st.session_state["at_fix_list"] = _at_fix_list

        _at_fix_list = st.session_state.get("at_fix_list")
        if _at_fix_list is not None:
            if not _at_fix_list:
                st.success("✅ Ошибочных пропусков не найдено!")
            else:
                st.error(f"⚠️ Найдено **{len(_at_fix_list)}** ошибочных пропусков — у людей с отчётом стоит пропуск")
                # Показываем список
                _at_fix_by_t = defaultdict(list)
                for _at_fr in _at_fix_list:
                    for _at_ft in (_at_fr["teachers"] or ["—"]):
                        _at_fix_by_t[_at_ft].append(_at_fr)
                for _at_ft in sorted(_at_fix_by_t):
                    with st.expander(f"👤 {_at_ft}  —  {len(_at_fix_by_t[_at_ft])} чел.", expanded=True):
                        for _at_fr in _at_fix_by_t[_at_ft]:
                            try:
                                _at_fd_fmt = datetime.strptime(_at_fr["date"], "%Y-%m-%d").strftime("%d.%m.%Y")
                            except Exception:
                                _at_fd_fmt = _at_fr["date"]
                            _at_fg = f" · _{_at_fr['group']}_" if _at_fr["group"] else ""
                            st.markdown(f"- **{_at_fd_fmt}** {_at_fr['student']}{_at_fg}")

                if st.button(
                    f"✅ Исправить — пометить всех {len(_at_fix_list)} как присутствовал (pass=false)",
                    type="primary",
                    key="at_fix_apply",
                ):
                    _at_fix_batch = []
                    for _at_fr in _at_fix_list:
                        _at_fix_item = {
                            "edUnitId":        _at_fr["edUnitId"],
                            "studentClientId": _at_fr["studentClientId"],
                            "date":            _at_fr["date"],
                            "pass":            False,
                            "description":     _at_fr["existing_desc"],
                        }
                        if _at_fr.get("minutes") is not None:
                            _at_fix_item["minutes"] = _at_fr["minutes"]
                        _at_fix_batch.append(_at_fix_item)

                    _at_fix_ok, _at_fix_err = 0, 0
                    with st.spinner("Исправляю…"):
                        _at_fix_url = f"{BASE_URL}/SetStudentPasses"
                        for _i in range(0, len(_at_fix_batch), 100):
                            _at_fc = _at_fix_batch[_i:_i + 100]
                            try:
                                _at_fr2 = requests.post(
                                    _at_fix_url,
                                    params={"authkey": api_key},
                                    json=_at_fc,
                                    timeout=60,
                                )
                                if _at_fr2.status_code in (200, 204):
                                    _at_fix_ok += len(_at_fc)
                                else:
                                    _at_fix_err += len(_at_fc)
                            except Exception:
                                _at_fix_err += len(_at_fc)
                    if _at_fix_err == 0:
                        st.success(f"✅ Исправлено {_at_fix_ok} записей!")
                        st.session_state["at_fix_list"] = []
                    else:
                        st.warning(f"Исправлено: {_at_fix_ok} ✅ / {_at_fix_err} ❌")

