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
    page_title="Проверка отчётов",
    page_icon="📋",
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
#  HISTORY — read / write / update
# ─────────────────────────────────────────────────────────────────────────────

def load_history() -> list:
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []


def save_history(records: list):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2, default=str)


def upsert_history(all_errors: list, period_from: str, period_to: str, reviewer: str):
    """Add/update history records for a check run. Preserves existing status/comments."""
    import uuid as _uuid
    records = load_history()
    now = datetime.now().isoformat(timespec="seconds")

    # Index existing records for this period
    existing = {
        (r["teacher"], r["date"], r["error_type"]): r
        for r in records
        if r["period_from"] == period_from and r["period_to"] == period_to
    }

    other = [r for r in records if not (r["period_from"] == period_from and r["period_to"] == period_to)]

    updated = []
    current_keys = set()

    for e in all_errors:
        key = (e["teacher"], e["date"], e["error_type"])
        current_keys.add(key)
        if key in existing:
            rec = existing[key].copy()
            rec.update({"count": e["count"], "error_description": e["error_description"], "updated_at": now})
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
                "status":            "open",
                "reviewer":          reviewer,
                "reviewer_comment":  "",
                "updated_at":        now,
            }
        updated.append(rec)

    # Records that existed before but are no longer errors = teacher fixed it themselves
    for key, rec in existing.items():
        if key not in current_keys:
            rec = rec.copy()
            if rec["status"] == "open":
                rec.update({"status": "resolved", "updated_at": now})
            updated.append(rec)

    save_history(other + updated)


def update_history_sent(period_from: str, period_to: str, teachers_sent: set, reviewer: str):
    """Mark no_report records as message_sent for the given teachers."""
    records = load_history()
    now = datetime.now().isoformat(timespec="seconds")
    for rec in records:
        if (rec["period_from"] == period_from and rec["period_to"] == period_to
                and rec["error_type"] == "no_report"
                and rec["teacher"] in teachers_sent
                and rec["status"] in ("open",)):
            rec.update({"status": "message_sent", "reviewer": reviewer, "updated_at": now})
    save_history(records)


def update_history_passes(period_from: str, period_to: str, reviewer: str):
    """Mark open/sent no_report records as pass_set for the given period."""
    records = load_history()
    now = datetime.now().isoformat(timespec="seconds")
    for rec in records:
        if (rec["period_from"] == period_from and rec["period_to"] == period_to
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
def load_teacher_info(creds_json_str: str = "") -> dict:
    """
    Returns dict: { "ФИО": {"is_intern": bool, "subject": str, "fired": bool} }
    Uses gspread (service account) if credentials provided, else tries public CSV URL.
    Cached for 1 hour.
    """
    import io

    def _parse_rows(rows: list) -> dict:
        result = {}
        for row in rows:
            name = str(row.get("ФИО", "")).strip()
            if not name or name.lower() == "nan":
                continue
            result[name] = {
                "is_intern": str(row.get("стажер", "")).strip().lower() == "стажер",
                "subject":   str(row.get("Предмет", "")).strip(),
                "fired":     str(row.get("Статус набора", "")).strip().lower() == "уволен",
            }
        return result

    # — Via gspread (service account) ————————————————————————————————————————
    if creds_json_str:
        try:
            gc = _gs_client(creds_json_str)
            sh = gc.open_by_key(_TEACHER_SHEET_ID)
            ws = sh.get_worksheet_by_id(int(_TEACHER_SHEET_GID))
            return _parse_rows(ws.get_all_records())
        except Exception:
            pass  # fall through to public URL

    # — Fallback: public CSV (works if sheet is shared "anyone with the link") —
    url = (
        f"https://docs.google.com/spreadsheets/d/{_TEACHER_SHEET_ID}"
        f"/export?format=csv&gid={_TEACHER_SHEET_GID}"
    )
    try:
        r = requests.get(url, timeout=15, allow_redirects=True)
        r.raise_for_status()
        df = pd.read_csv(io.StringIO(r.content.decode("utf-8", errors="replace")))
        return _parse_rows(df.to_dict("records"))
    except Exception:
        return {}


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
    "Если занятие было проведено — просьба заполнить отчёт "
    "(тема, д/з, комментарий). Если занятие не состоялось — "
    "необходимо проставить отмену занятия в системе."
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
        batch = [
            {
                "edUnitId":        p["edUnitId"],
                "studentClientId": p["studentClientId"],
                "date":            p["date"],
                "pass":            True,
                "description":     absence_comment,
            }
            for p in passes[i : i + batch_size]
        ]
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


# ─────────────────────────────────────────────────────────────────────────────
#  HEADER — title + dates + settings popover
# ─────────────────────────────────────────────────────────────────────────────

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
        subdomain = st.text_input("Субдомен", value=cfg.get("subdomain", "matrix"), key="cfg_subdomain")
        api_key   = st.text_input("API ключ",  value=cfg.get("api_key", ""), type="password", key="cfg_api_key")

        st.markdown("**📨 Пачка**")
        pachca_token    = st.text_input("Токен Пачки", value=cfg.get("pachca_token", ""), type="password", key="cfg_pachca")
        create_reminder = st.checkbox("Создавать напоминания", value=cfg.get("create_reminder", False), key="cfg_reminder")
        reminder_days   = st.number_input("Дней до дедлайна", min_value=1, max_value=14,
                                          value=int(cfg.get("reminder_days", 1)), key="cfg_rdays")

        st.markdown("**👤 Проверяющий**")
        reviewer_name = st.text_input("Имя", value=cfg.get("reviewer_name", "Артём"), key="cfg_reviewer")

        with st.expander("📊 Google Таблица (необязательно)"):
            sheet_id   = st.text_input("ID таблицы",  value=cfg.get("sheet_id", ""),    key="cfg_sheet_id")
            sheet_name = st.text_input("Имя листа",   value=cfg.get("sheet_name", "Лист1"), key="cfg_sheet_name")
            creds_file = st.file_uploader("Service Account JSON", type=["json"], key="cfg_creds")
            creds_json = creds_file.read().decode() if creds_file else cfg.get("creds_json", None)

        st.divider()
        st.markdown("**📋 Данные о преподавателях**")
        _ti_count = len(_teacher_info)
        if _ti_count:
            st.caption(f"✅ Загружено: {_ti_count} преподавателей")
        else:
            st.caption("⚠️ Не загружено — нужен Service Account JSON выше")
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
                    "subdomain": subdomain, "api_key": api_key,
                    "pachca_token": pachca_token, "create_reminder": create_reminder,
                    "reminder_days": reminder_days, "sheet_id": sheet_id,
                    "sheet_name": sheet_name, "reviewer_name": reviewer_name,
                    "creds_json": creds_json,
                })
                st.success("✅ Сохранено!")

DATE_FROM = date_from_val.strftime("%Y-%m-%d")
DATE_TO   = date_to_val.strftime("%Y-%m-%d")
BASE_URL  = f"https://{subdomain}.t8s.ru/Api/V2"

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📥 Загрузить данные", "✉️ Сообщения", "📤 Отправить", "📚 История", "📊 Статистика"])

# Load teacher info from Google Sheets (cached 1 h)
_teacher_info = load_teacher_info(creds_json_str=creds_json or "")


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
                passes_to_set = []   # конкретные записи для SetStudentPasses

                for eus in eu_students:
                    eu_id      = eus.get("EdUnitId")
                    student_id = eus.get("StudentClientId")
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
                            passes_to_set.append({
                                "edUnitId":        eu_id,
                                "studentClientId": student_id,
                                "date":            d,
                            })
                        if is_pass and not desc:
                            for t in teachers:
                                no_comment[t][d] += 1

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
    "resolved":     ("🟢 Исправлено самостоятельно",   "#22c55e"),
}
_STATUS_PILL = {
    "open":         "pill pill-open",
    "message_sent": "pill pill-sent",
    "pass_set":     "pill pill-pass",
    "handled":      "pill pill-handled",
    "skipped":      "pill pill-skipped",
    "resolved":     "pill pill-handled",
}
_STATUS_FILTER_MAP = {
    "🔴 Открыто":                   "open",
    "💬 Сообщение отправлено":       "message_sent",
    "🚫 Пропуск выставлен":         "pass_set",
    "✅ Обработано":                 "handled",
    "⚪ Пропущено":                 "skipped",
    "🟢 Исправлено самостоятельно":  "resolved",
}

with tab4:
    st.subheader("📚 История проверок")

    history_records = load_history()

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

        # ── Группируем по периоду ─────────────────────────────────────────────
        from collections import OrderedDict as _OD

        def _period_sort_key(pf_pt):
            try:
                return datetime.strptime(pf_pt[0], "%Y-%m-%d")
            except Exception:
                return datetime.min

        periods_dict = _OD()
        for r in history_records:
            pk = (r["period_from"], r["period_to"])
            periods_dict.setdefault(pk, []).append(r)

        # Sort periods newest first
        sorted_periods = sorted(periods_dict.items(), key=lambda x: _period_sort_key(x[0]), reverse=True)

        _ERR_FILTER_MAP = {
            "Нет отчёта":                    "no_report",
            "Нет комментария к пропуску":    "no_abs_comment",
        }

        for (pf, pt), recs in sorted_periods:
            # Apply filters
            filtered = recs
            if f_teacher != "Все":
                filtered = [r for r in filtered if r["teacher"] == f_teacher]
            if f_status != "Все":
                filtered = [r for r in filtered if r["status"] == _STATUS_FILTER_MAP.get(f_status)]
            if f_error != "Все":
                filtered = [r for r in filtered if r["error_type"] == _ERR_FILTER_MAP.get(f_error)]
            if not filtered:
                continue

            # Period header
            try:
                pf_str = datetime.strptime(pf, "%Y-%m-%d").strftime("%d.%m")
                pt_str = datetime.strptime(pt, "%Y-%m-%d").strftime("%d.%m.%Y")
            except Exception:
                pf_str, pt_str = pf, pt

            open_n    = sum(1 for r in filtered if r["status"] == "open")
            handled_n = sum(1 for r in filtered if r["status"] == "handled")
            header    = f"📅 {pf_str} — {pt_str}  ·  {len(filtered)} записей"
            if open_n:
                header += f"  ·  🔴 {open_n} открытых"
            if handled_n:
                header += f"  ·  ✅ {handled_n} обработано"

            with st.expander(header, expanded=(open_n > 0)):

                # Per-period mini bulk bar
                _period_ids        = {r["id"] for r in filtered}
                _period_action_ids = {r["id"] for r in filtered if r["status"] in _actionable_statuses}
                _period_close_ids  = {r["id"] for r in filtered if r["status"] in _closeable_statuses}
                _pb1, _pb2, _pb3, _pb4, _pb5 = st.columns([1.2, 1.2, 1.4, 1.4, 4])
                if _pb1.button("☑️ Все", key=f"psel_{pf}_{pt}", use_container_width=True):
                    for rid in _period_ids:
                        st.session_state[f"hsel_{rid}"] = True
                    st.rerun()
                if _pb2.button("⬜ Снять", key=f"pdes_{pf}_{pt}", use_container_width=True):
                    for rid in _period_ids:
                        st.session_state[f"hsel_{rid}"] = False
                    st.rerun()
                _p_sel_open   = sum(1 for rid in _period_action_ids if st.session_state.get(f"hsel_{rid}", False))
                _p_sel_closed = sum(1 for rid in _period_close_ids  if st.session_state.get(f"hsel_{rid}", False))
                if _p_sel_open > 0:
                    if _pb3.button(f"✅ Обработано ({_p_sel_open})", key=f"ph_{pf}_{pt}",
                                   use_container_width=True, type="primary"):
                        for r in filtered:
                            if st.session_state.get(f"hsel_{r['id']}", False) and r["status"] in _actionable_statuses:
                                update_history_record(r["id"], "handled", r.get("reviewer_comment", ""), reviewer_name)
                                st.session_state[f"hsel_{r['id']}"] = False
                        st.rerun()
                if _p_sel_closed > 0:
                    if _pb4.button(f"↺ Переоткрыть ({_p_sel_closed})", key=f"pr_{pf}_{pt}",
                                   use_container_width=True):
                        for r in filtered:
                            if st.session_state.get(f"hsel_{r['id']}", False) and r["status"] in _closeable_statuses:
                                update_history_record(r["id"], "open", r.get("reviewer_comment", ""), reviewer_name)
                                st.session_state[f"hsel_{r['id']}"] = False
                        st.rerun()

                st.divider()

                for rec in sorted(filtered, key=lambda r: (r["teacher"], r["date"])):
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
        exp_records = history_records
        if f_teacher != "Все":
            exp_records = [r for r in exp_records if r["teacher"] == f_teacher]
        if f_status != "Все":
            exp_records = [r for r in exp_records if r["status"] == _STATUS_FILTER_MAP.get(f_status)]
        if f_error != "Все":
            exp_records = [r for r in exp_records if r["error_type"] == _ERR_FILTER_MAP.get(f_error)]

        if exp_records:
            def _fmt_date(d):
                try: return datetime.strptime(d, "%Y-%m-%d").strftime("%d.%m.%Y")
                except: return d

            export_df = pd.DataFrame([{
                "Период":         f"{_fmt_date(r['period_from'])} — {_fmt_date(r['period_to'])}",
                "Дата":           _fmt_date(r["date"]),
                "Преподаватель":  r["teacher"],
                "Тег ученика":    r.get("student_tag", "Все"),
                "Ошибка":         r["error_description"],
                "Кол-во":         r.get("count", ""),
                "Статус":         _STATUS_META.get(r["status"], ("?", ""))[0],
                "Проверяющий":    r.get("reviewer", ""),
                "Комментарий":    r.get("reviewer_comment", ""),
                "Обновлено":      r.get("updated_at", ""),
            } for r in exp_records])
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
    return r["status"] in ("handled", "pass_set", "message_sent", "resolved")

def _is_open(r):
    return r["status"] == "open"

with tab5:
    st.subheader("📊 Статистика по проверкам")

    stat_records = load_history()

    if not stat_records:
        st.info("История пуста — запустите проверку для получения статистики.")
    else:
        # ── Period filter ─────────────────────────────────────────────────────
        all_stat_periods = sorted(
            set((r["period_from"], r["period_to"]) for r in stat_records),
            key=lambda x: x[0], reverse=True,
        )

        _fmode_col, _fp1_col, _fp2_col = st.columns([1.5, 1.5, 1.5])
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
            # Custom date range
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
            sel_period = "custom"  # used later to control dynamics block

        # ── Top metrics ───────────────────────────────────────────────────────
        total_rec   = len(sr)
        processed_n = sum(1 for r in sr if _is_processed(r))
        open_n      = sum(1 for r in sr if _is_open(r))
        skipped_n   = sum(1 for r in sr if r["status"] == "skipped")
        teachers_n  = len(set(r["teacher"] for r in sr))

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("📋 Всего ошибок",      total_rec)
        m2.metric("✅ Обработано",         processed_n)
        m3.metric("🔴 Не обработано",      open_n)
        m4.metric("⚪ Пропущено",          skipped_n)
        m5.metric("👩‍🏫 Преподавателей",    teachers_n)

        st.divider()

        # ── Two-column layout: teacher table + summary ────────────────────────
        col_tbl, col_sum = st.columns([3, 1])

        with col_tbl:
            st.markdown("**По преподавателям**")

            teacher_rows = []
            for teacher in sorted(set(r["teacher"] for r in sr)):
                t = [r for r in sr if r["teacher"] == teacher]
                done  = sum(1 for r in t if _is_processed(r))
                total = len(t)
                teacher_rows.append({
                    "Преподаватель":   teacher,
                    "✅ Обработано":    done,
                    "🔴 Не обработано": sum(1 for r in t if _is_open(r)),
                    "Всего":           total,
                    "% выполнено":     round(done / total * 100) if total else 0,
                })

            df_teachers = pd.DataFrame(teacher_rows)

            st.dataframe(
                df_teachers,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Преподаватель":    st.column_config.TextColumn(width="large"),
                    "✅ Обработано":    st.column_config.NumberColumn(width="small"),
                    "🔴 Не обработано": st.column_config.NumberColumn(width="small"),
                    "Всего":            st.column_config.NumberColumn(width="small"),
                    "% выполнено":      st.column_config.ProgressColumn(
                        "% выполнено",
                        help="Доля обработанных ошибок",
                        format="%d%%",
                        min_value=0,
                        max_value=100,
                    ),
                },
            )

            # Export
            csv_stat = df_teachers.to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ Скачать CSV", csv_stat, "stats.csv", "text/csv")

        with col_sum:
            st.markdown("**Сводка по статусам**")
            summary_rows = [
                {"Статус": "✅ Обработано",                 "Кол-во": sum(1 for r in sr if r["status"] == "handled")},
                {"Статус": "🟢 Исправлено самостоятельно",  "Кол-во": sum(1 for r in sr if r["status"] == "resolved")},
                {"Статус": "💬 Сообщение отправлено",        "Кол-во": sum(1 for r in sr if r["status"] == "message_sent")},
                {"Статус": "🚫 Пропуск выставлен",          "Кол-во": sum(1 for r in sr if r["status"] == "pass_set")},
                {"Статус": "🔴 Открыто",                    "Кол-во": sum(1 for r in sr if r["status"] == "open")},
                {"Статус": "⚪ Пропущено",                  "Кол-во": sum(1 for r in sr if r["status"] == "skipped")},
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

            period_rows = []
            for pf, pt in all_stat_periods:
                p = [r for r in stat_records if r["period_from"] == pf and r["period_to"] == pt]
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
