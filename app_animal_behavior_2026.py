# app_animal_behavior_2026_oauth_A_full_v2_6_pdf_preload_status.py
# ------------------------------------------------------------
# 版本變更說明（覆蓋版｜v2.7｜PDF預載＋跳頁狀態列＋iPhone iframe 重載）
# 1) ✅ 移除「摘要索引」匯入/解析/比對功能（你說不需要）
# 2) ✅ 預載摘要集 PDF（本機掛載檔優先）：2026 動物行為研討會摘要集.pdf
#    - 若檔不存在：可改用上傳 PDF 或填入 PDF URL
# 3) ✅ PDF 跳頁狀態列：顯示「目前顯示第 N 頁」＋「來源/命中理由」
# 4) ✅ iPhone Safari 常見 iframe 不重載 → 加入 cache-buster 參數，跳頁更可靠
# 5) ✅ Mobile mode 預設開啟
# 6) ✅ 搜尋列（query）與日期選擇（days）「永遠在頁面上方」：不放在折疊 expander 裡
#
# ------------------------------------------------------------
# 重要部署提醒（可選）
# - 若你要做「PDF 內文搜尋」(例如輸入作者/編號在 PDF 裡找頁碼)，需要 PyPDF2：
#   requirements.txt 加入：PyPDF2
# - 若沒有 PyPDF2，本 app 仍可用「手動頁碼跳頁」與「固定預載 PDF」。
#
# ------------------------------------------------------------
from __future__ import annotations

import os
import re
import io
import json
import time
import base64
import hashlib
import sqlite3
import datetime as dt
from dataclasses import dataclass
from typing import Dict, Tuple, Optional, List, Set, Any

import pandas as pd
import streamlit as st

# ----------------------------
# Streamlit compatibility helpers (avoid blank page on older Streamlit)
# ----------------------------
def _st_link_button(label: str, url: str, use_container_width: bool = False):
    if hasattr(st, "link_button"):
        return _st_link_button(label, url, use_container_width=use_container_width)
    # fallback
    return st.markdown(f"- [{label}]({url})")

def _container(border: bool = False):
    # streamlit < 1.25 may not support border arg
    try:
        return st.container(border=border)  # type: ignore
    except TypeError:
        return st.container()


APP_TITLE = "2026 動物行為暨生態研討會｜議程搜尋＋個人化行事曆"
DEFAULT_EXCEL_PATH = "2026 動行議程.xlsx"

# ✅ 預載 PDF（掛載在 Streamlit Cloud/本地專案時，請把檔案放到專案可讀路徑）
DEFAULT_PDF_PATH = "2026 動物行為研討會摘要集.pdf"
# 預載 PDF：會依序嘗試下列路徑（先找到先用）
# - Streamlit Cloud：請把 PDF 放進 repo（與 app 同層或 static/ 資料夾）
# - 本機：可用絕對路徑
PRELOAD_PDF_CANDIDATES = [
    DEFAULT_PDF_PATH,
    '2026 動物行為研討會摘要集.pdf',
    'static/2026 動物行為研討會摘要集.pdf',
    'data/2026 動物行為研討會摘要集.pdf',
]

def _pick_first_existing_path(paths):
    for p in paths:
        try:
            if p and os.path.exists(p):
                return p
        except Exception:
            pass
    return None

@st.cache_data(show_spinner=False)
def load_pdf_bytes_from_path(path: str) -> Optional[bytes]:
    if not path:
        return None
    try:
        with open(path, 'rb') as f:
            return f.read()
    except Exception:
        return None


DATE_MAP = {
    "D1": dt.date(2026, 1, 26),
    "D2": dt.date(2026, 1, 27),
}
TITLE_SPAN_RIGHT = 6

# ----------------------------
# CSS（Mobile：隱藏 sidebar + 放大點擊目標）
# ----------------------------
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.markdown(
    """
<style>
@media (max-width: 768px) {
  section[data-testid="stSidebar"] { display: none !important; }
  div.block-container { padding-top: 0.8rem; padding-left: 0.8rem; padding-right: 0.8rem; }
  .stButton button, .stDownloadButton button { padding: 0.65rem 0.9rem; font-size: 1rem; }
  .stToggle { transform: scale(1.05); transform-origin: left center; }
}
.small-muted { color: rgba(49, 51, 63, 0.65); font-size: 0.9rem; }
.hr-soft { margin: 0.35rem 0 0.65rem 0; border-top: 1px solid rgba(49,51,63,0.15); }
</style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# 方案A：Google OAuth + Persisted User State (SQLite)
# ============================================================
APP_DB_PATH = "user_state.db"
APP_STATE_TABLE = "user_state_v1"

# Optional Google OAuth dependencies
try:
    from google_auth_oauthlib.flow import Flow
    from google.oauth2 import id_token as google_id_token
    from google.auth.transport import requests as google_requests
    _GOOGLE_LIBS_OK = True
except Exception:
    _GOOGLE_LIBS_OK = False


def _get_secret(path: str, default: Optional[str] = None) -> Optional[str]:
    """Read from st.secrets with dotted path, e.g. 'google_oauth.client_id'."""
    try:
        cur = st.secrets
        for part in path.split("."):
            cur = cur[part]
        return str(cur)
    except Exception:
        return default


def _sha256(s: str) -> str:
    return hashlib.sha256(s.encode("utf-8")).hexdigest()


def _b64url_encode(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).decode("utf-8").rstrip("=")


def _b64url_decode(s: str) -> bytes:
    pad = "=" * (-len(s) % 4)
    return base64.urlsafe_b64decode(s + pad)


def hmac_compare(a: str, b: str) -> bool:
    if len(a) != len(b):
        return False
    out = 0
    for x, y in zip(a.encode("utf-8"), b.encode("utf-8")):
        out |= x ^ y
    return out == 0


def _sign_payload(payload: Dict[str, Any], secret: str) -> str:
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")
    sig = _sha256(_b64url_encode(raw) + secret)
    token = _b64url_encode(raw) + "." + sig
    return token


def _verify_payload(token: str, secret: str) -> Optional[Dict[str, Any]]:
    try:
        raw_b64, sig = token.split(".", 1)
        expected = _sha256(raw_b64 + secret)
        if not hmac_compare(sig, expected):
            return None
        payload = json.loads(_b64url_decode(raw_b64).decode("utf-8"))
        return payload
    except Exception:
        return None


def db_init(db_path: str = APP_DB_PATH) -> None:
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            f"""
            CREATE TABLE IF NOT EXISTS {APP_STATE_TABLE} (
                user_id TEXT PRIMARY KEY,
                state_json TEXT NOT NULL,
                updated_at INTEGER NOT NULL
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def db_load_state(user_id: str, db_path: str = APP_DB_PATH) -> Dict[str, Any]:
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(f"SELECT state_json FROM {APP_STATE_TABLE} WHERE user_id = ?", (user_id,))
        row = cur.fetchone()
        if not row:
            return {}
        return json.loads(row[0])
    except Exception:
        return {}
    finally:
        conn.close()


def db_save_state(user_id: str, state: Dict[str, Any], db_path: str = APP_DB_PATH) -> None:
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        now = int(time.time())
        cur.execute(
            f"""
            INSERT INTO {APP_STATE_TABLE} (user_id, state_json, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET
                state_json=excluded.state_json,
                updated_at=excluded.updated_at
            """,
            (user_id, json.dumps(state, ensure_ascii=False), now),
        )
        conn.commit()
    finally:
        conn.close()


@dataclass
class AuthUser:
    user_id: str
    email: Optional[str]
    name: Optional[str]
    picture: Optional[str]


def get_oauth_config() -> Optional[Dict[str, str]]:
    client_id = _get_secret("google_oauth.client_id")
    client_secret = _get_secret("google_oauth.client_secret")
    redirect_uri = _get_secret("google_oauth.redirect_uri")
    cookie_secret = _get_secret("google_oauth.cookie_secret")
    if not all([client_id, client_secret, redirect_uri, cookie_secret]):
        return None
    return {
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": redirect_uri,
        "cookie_secret": cookie_secret,
    }


def build_flow(config: Dict[str, str]) -> "Flow":
    scopes = [
        "openid",
        "https://www.googleapis.com/auth/userinfo.email",
        "https://www.googleapis.com/auth/userinfo.profile",
    ]
    client_config = {
        "web": {
            "client_id": config["client_id"],
            "client_secret": config["client_secret"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
        }
    }
    return Flow.from_client_config(client_config, scopes=scopes, redirect_uri=config["redirect_uri"])


def auth_ui_sidebar() -> Optional[AuthUser]:
    st.session_state.setdefault("auth_user", None)
    st.session_state.setdefault("auth_error", None)

    config = get_oauth_config()
    if (config is None) or (not _GOOGLE_LIBS_OK):
        return None

    if st.session_state.get("auth_user") is not None:
        return st.session_state["auth_user"]

    qp = st.query_params
    code = qp.get("code", None)
    state_token = qp.get("state", None)

    cookie_secret = config["cookie_secret"]

    if not code:
        flow = build_flow(config)
        state_payload = {"ts": int(time.time()), "nonce": _sha256(str(time.time()) + os.urandom(8).hex())}
        signed_state = _sign_payload(state_payload, cookie_secret)
        auth_url, _ = flow.authorization_url(
            access_type="offline",
            include_granted_scopes="true",
            state=signed_state,
            prompt="select_account",
        )
        _st_link_button("用 Google 登入（記住我的選擇）", auth_url, use_container_width=True)
        return None

    if not state_token:
        st.session_state["auth_error"] = "OAuth callback missing state."
        return None

    verified = _verify_payload(state_token, cookie_secret)
    if verified is None:
        st.session_state["auth_error"] = "OAuth state verification failed."
        return None

    try:
        flow = build_flow(config)
        flow.fetch_token(code=code)
        creds = flow.credentials
        req = google_requests.Request()
        idinfo = google_id_token.verify_oauth2_token(creds.id_token, req, config["client_id"])

        user = AuthUser(
            user_id=str(idinfo.get("sub")),
            email=idinfo.get("email"),
            name=idinfo.get("name"),
            picture=idinfo.get("picture"),
        )
        st.session_state["auth_user"] = user
        st.query_params.clear()
        return user
    except Exception as e:
        st.session_state["auth_error"] = f"OAuth failed: {e}"
        return None


def logout_ui():
    if st.button("登出", use_container_width=True):
        for k in ["auth_user", "auth_error"]:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()


class UserStateManager:
    """Persistent state if logged in; session-only if not."""
    def __init__(self, user: Optional[AuthUser]):
        self.user = user
        self._state: Dict[str, Any] = {}
        self._loaded = False

    def load(self):
        if self._loaded:
            return
        if self.user is not None:
            self._state = db_load_state(self.user.user_id)
        else:
            self._state = st.session_state.get("_anon_state", {})
        self._loaded = True

    def get(self, key: str, default: Any = None) -> Any:
        self.load()
        return self._state.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self.load()
        self._state[key] = value

    def save(self) -> None:
        self.load()
        if self.user is not None:
            db_save_state(self.user.user_id, self._state)
        else:
            st.session_state["_anon_state"] = self._state


# ----------------------------
# Parsing helpers
# ----------------------------
_TIME_RANGE_RE = re.compile(r"^(\d{1,2}:\d{2})\s*[-–~]\s*(\d{1,2}:\d{2})$")
_TIME_RANGE_IN_TEXT_RE = re.compile(r"(\d{1,2}:\d{2})\s*[-–~]\s*(\d{1,2}:\d{2})")


def _parse_time_str(s: str) -> Optional[dt.time]:
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None
    s = s.replace("：", ":").replace("．", ".")
    s = re.split(r"\s|\(|（", s)[0].strip()

    m = re.fullmatch(r"(\d{1,2})(?::(\d{1,2}))?", s)
    if m:
        h = int(m.group(1))
        mi = int(m.group(2) or 0)
        if 0 <= h <= 23 and 0 <= mi <= 59:
            return dt.time(hour=h, minute=mi)

    m = re.fullmatch(r"(\d{1,2})\.(\d{1,2})", s)
    if m:
        h = int(m.group(1))
        mi = int(m.group(2))
        if 0 <= h <= 23 and 0 <= mi <= 59:
            return dt.time(hour=h, minute=mi)

    m = re.fullmatch(r"(\d{2})(\d{2})", s)
    if m:
        h = int(m.group(1))
        mi = int(m.group(2))
        if 0 <= h <= 23 and 0 <= mi <= 59:
            return dt.time(hour=h, minute=mi)

    m = re.fullmatch(r"(\d{1,2}):(\d{2})", s)
    if m:
        hh = int(m.group(1))
        mm = int(m.group(2))
        if 0 <= hh <= 23 and 0 <= mm <= 59:
            return dt.time(hour=hh, minute=mm)

    return None


def _parse_time_range(x: object) -> Optional[Tuple[str, str]]:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip()
    m = _TIME_RANGE_RE.match(s)
    if not m:
        return None
    return m.group(1), m.group(2)


def _extract_time_range_from_text(s: object) -> Optional[Tuple[str, str]]:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    txt = str(s)
    m = _TIME_RANGE_IN_TEXT_RE.search(txt)
    if not m:
        return None
    return m.group(1), m.group(2)


def _safe_str(x: object) -> Optional[str]:
    if x is None:
        return None
    if isinstance(x, float) and pd.isna(x):
        return None
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return None
    return s


def _find_col(cols: List[str], candidates: List[str]) -> Optional[str]:
    for c in cols:
        if not isinstance(c, str):
            continue
        for cand in candidates:
            if cand in c:
                return c
    return None


def _find_col_prefer_candidates(cols: List[str], candidates: List[str]) -> Optional[str]:
    for cand in candidates:
        for c in cols:
            if not isinstance(c, str):
                continue
            if cand in c:
                return c
    return None


def _join_nonempty(parts: List[Optional[str]], sep: str = " ") -> Optional[str]:
    xs = [p.strip() for p in parts if p and str(p).strip()]
    if not xs:
        return None
    s = sep.join(xs)
    s = re.sub(r"\s+", " ", s).strip()
    return s or None


def _extract_title_with_span(row: pd.Series, cols: List[str], base_col: Optional[str], span_right: int) -> Optional[str]:
    if not base_col or base_col not in row.index:
        return None
    try:
        i0 = cols.index(base_col)
    except ValueError:
        return _safe_str(row.get(base_col))

    parts: List[Optional[str]] = []
    for j in range(i0, min(len(cols), i0 + 1 + span_right)):
        v = _safe_str(row.get(cols[j]))
        cname = str(cols[j])
        if j != i0 and re.search(r"(單位|主持|講者|作者|編號|時間|報告時間)", cname):
            continue
        parts.append(v)

    title = _join_nonempty(parts, sep=" ")
    if title in ("投稿題目", "演講主題", "主題領域", "題目", "講題"):
        return None
    return title


def _fallback_title_from_row(row: pd.Series) -> Optional[str]:
    best: Optional[str] = None
    best_score = -1
    for _, v in row.items():
        s = _safe_str(v)
        if not s:
            continue
        if _parse_time_range(s):
            continue
        if re.fullmatch(r"[A-Za-z]?\d{2,6}", s):
            continue
        if re.search(r"^D[12]$", s.strip()):
            continue
        if len(s) < 8:
            continue

        score = len(s)
        if re.search(r"[\u4e00-\u9fff]", s):
            score += 10
        if " " in s:
            score += 3
        if score > best_score:
            best = s
            best_score = score
    return best


@st.cache_data(show_spinner=False)
def load_excel_all_sheets(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    return {name: pd.read_excel(io.BytesIO(file_bytes), sheet_name=name) for name in xl.sheet_names}


@st.cache_data(show_spinner=False)
def build_master_df(sheets: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    master: List[Dict[str, object]] = []

    if "大會議程" in sheets:
        df = sheets["大會議程"].copy()
        cur_day: Optional[str] = None

        for _, row in df.iterrows():
            first = row.iloc[0]
            if isinstance(first, str) and first.strip() in ("D1", "D2"):
                cur_day = first.strip()
                continue

            tr = _parse_time_range(first) if isinstance(first, str) else None
            if not (tr and cur_day):
                continue

            start, end = tr
            t_start = _parse_time_str(start)
            t_end = _parse_time_str(end)
            if t_start is None or t_end is None:
                continue

            for col in df.columns[1:]:
                val = row[col]
                title = _safe_str(val)
                if not title:
                    continue
                if "請點我" in title:
                    continue

                room = str(col).strip()
                master.append(
                    dict(
                        source_sheet="大會議程",
                        day=cur_day,
                        date=DATE_MAP[cur_day].isoformat(),
                        room=room,
                        location=room,
                        code=None,
                        session=None,
                        title=title,
                        speaker=None,
                        affiliation=None,
                        start=start,
                        end=end,
                        start_dt=dt.datetime.combine(DATE_MAP[cur_day], t_start),
                        end_dt=dt.datetime.combine(DATE_MAP[cur_day], t_end),
                        kind="main_schedule",
                    )
                )

    for sheet_name, df0 in sheets.items():
        if sheet_name == "大會議程":
            continue

        if str(sheet_name).strip() == "海報":
            dfp = df0.copy()
            cols_p = [str(c) for c in dfp.columns]

            col_code_p = cols_p[0] if len(cols_p) >= 1 else None
            col_author_p = cols_p[1] if len(cols_p) >= 2 else None
            col_title_p = cols_p[2] if len(cols_p) >= 3 else None
            col_area_p = cols_p[3] if len(cols_p) >= 4 else None

            cur_day: Optional[str] = None
            poster_session_tr: Optional[Tuple[str, str]] = None

            for _, rowp in dfp.iterrows():
                v0 = _safe_str(rowp.get(col_code_p)) if col_code_p else None

                if v0 and "Day 1" in v0:
                    cur_day = "D1"
                    poster_session_tr = None
                    continue
                if v0 and "Day 2" in v0:
                    cur_day = "D2"
                    poster_session_tr = None
                    continue

                if v0 and ("海報競賽時間" in v0 or "海報解說時間" in v0):
                    poster_session_tr = _extract_time_range_from_text(v0)
                    continue

                if not (v0 and re.fullmatch(r"P[A-Z]\d{2}", v0.strip())):
                    continue
                if not cur_day or cur_day not in DATE_MAP:
                    continue
                if not poster_session_tr:
                    continue

                start, end = poster_session_tr
                t_start = _parse_time_str(start)
                t_end = _parse_time_str(end)
                if t_start is None or t_end is None:
                    continue

                author = _safe_str(rowp.get(col_author_p)) if col_author_p else None
                title = _safe_str(rowp.get(col_title_p)) if col_title_p else None
                area = _safe_str(rowp.get(col_area_p)) if col_area_p else None

                master.append(
                    dict(
                        source_sheet=sheet_name,
                        day=cur_day,
                        date=DATE_MAP[cur_day].isoformat(),
                        room="海報",
                        location="海報區",
                        code=v0.strip(),
                        session=area,
                        title=title,
                        speaker=author,
                        affiliation=None,
                        start=start,
                        end=end,
                        start_dt=dt.datetime.combine(DATE_MAP[cur_day], t_start),
                        end_dt=dt.datetime.combine(DATE_MAP[cur_day], t_end),
                        kind="poster",
                    )
                )
            continue

        df = df0.copy()

        def _infer_default_day_from_sheet(sheet: str, df_: pd.DataFrame) -> Optional[str]:
            if "D1" in sheet:
                return "D1"
            if "D2" in sheet:
                return "D2"
            try:
                c0 = str(df_.columns[0])
                if "D1" in c0:
                    return "D1"
                if "D2" in c0:
                    return "D2"
            except Exception:
                pass
            return None

        def _promote_header_row_if_needed(df_: pd.DataFrame) -> pd.DataFrame:
            cols_ = [str(c) for c in df_.columns]
            if _find_col(cols_, ["時間"]):
                return df_
            header_idx: Optional[int] = None
            for i in range(min(len(df_), 30)):
                row_vals = [str(x).strip() for x in df_.iloc[i].tolist()]
                if any(v == "時間" or ("時間" in v and len(v) <= 6) for v in row_vals):
                    header_idx = i
                    break
            if header_idx is None:
                return df_
            new_cols = [str(x).strip() for x in df_.iloc[header_idx].tolist()]
            df2 = df_.iloc[header_idx + 1 :].copy()
            df2.columns = new_cols
            return df2

        default_day = _infer_default_day_from_sheet(sheet_name, df)
        df = _promote_header_row_if_needed(df)

        cols = [str(c) for c in df.columns]
        col_time = _find_col(cols, ["時間"])
        col_code = _find_col(cols, ["編號"])
        col_report = _find_col(cols, ["報告時間"])
        if str(sheet_name).strip() in ("S101國家公園", "E102林保署"):
            speaker_candidates = ["講者", "作者姓名", "主持人"]
        else:
            speaker_candidates = ["作者姓名", "講者", "主持人"]
        col_speaker = _find_col_prefer_candidates(cols, speaker_candidates)

        col_aff = _find_col(cols, ["講者單位", "單位"])

        title_candidates = [
            "投稿題目", "演講主題", "主題領域", "題目", "講題", "報告題目", "題名",
            "Title", "TITLE", "Topic", "TOPIC", "Presentation Title",
        ]
        col_title = _find_col(cols, title_candidates)

        cur_day: Optional[str] = default_day
        current_session_time: Optional[str] = None

        for _, row in df.iterrows():
            first = row.iloc[0]

            if isinstance(first, str) and re.search(r"/D[12]\s*$", first.strip()):
                cur_day = first.strip().split("/")[-1]
                current_session_time = None
                continue

            if col_time and isinstance(row.get(col_time), str):
                tr_block = _parse_time_range(row.get(col_time))
                if tr_block:
                    current_session_time = str(row.get(col_time)).strip()

            if not cur_day:
                continue

            tr: Optional[Tuple[str, str]] = None
            if col_report and isinstance(row.get(col_report), str):
                tr = _parse_time_range(row.get(col_report))
            if tr is None and current_session_time:
                tr = _parse_time_range(current_session_time)
            if tr is None and col_time and isinstance(row.get(col_time), str):
                tr = _parse_time_range(row.get(col_time))
            if tr is None:
                continue

            start, end = tr
            t_start = _parse_time_str(start)
            t_end = _parse_time_str(end)
            if t_start is None or t_end is None:
                continue

            code = _safe_str(row.get(col_code)) if col_code else None
            speaker = _safe_str(row.get(col_speaker)) if col_speaker else None
            aff = _safe_str(row.get(col_aff)) if col_aff else None

            title = _extract_title_with_span(row, cols, col_title, TITLE_SPAN_RIGHT)
            if not title:
                title = _fallback_title_from_row(row)

            if title in ("投稿題目", "演講主題", "主題領域", "題目", "講題") and (speaker is None) and (code is None):
                continue
            if (not title) and (not speaker) and (not code):
                continue
            if cur_day not in DATE_MAP:
                continue

            master.append(
                dict(
                    source_sheet=sheet_name,
                    day=cur_day,
                    date=DATE_MAP[cur_day].isoformat(),
                    room=sheet_name,
                    location=sheet_name,
                    code=code,
                    session=None,
                    title=title,
                    speaker=speaker,
                    affiliation=aff,
                    start=start,
                    end=end,
                    start_dt=dt.datetime.combine(DATE_MAP[cur_day], t_start),
                    end_dt=dt.datetime.combine(DATE_MAP[cur_day], t_end),
                    kind="room_detail",
                )
            )

    mdf = pd.DataFrame(master)
    if len(mdf) == 0:
        mdf = pd.DataFrame(columns=[
            "source_sheet","day","date","room","location","code","session","title",
            "speaker","affiliation","start","end","start_dt","end_dt","kind",
        ])

    mdf = mdf.drop_duplicates(subset=["date", "room", "start", "end", "code", "title", "speaker"], keep="first")
    mdf = mdf.sort_values(["start_dt", "room", "code"], na_position="last").reset_index(drop=True)

    mdf["display_date"] = mdf["date"].map(
        lambda s: "D1 (2026-01-26)" if s == "2026-01-26" else ("D2 (2026-01-27)" if s == "2026-01-27" else str(s))
    )
    mdf["time"] = mdf["start"].astype(str) + "–" + mdf["end"].astype(str)
    mdf["where"] = mdf["location"].fillna(mdf["room"])
    mdf["key"] = (
        mdf["date"].astype(str)
        + "|" + mdf["room"].astype(str)
        + "|" + mdf["start"].astype(str)
        + "|" + mdf["end"].astype(str)
        + "|" + mdf["code"].fillna("").astype(str)
        + "|" + mdf["title"].fillna("").astype(str)
    )
    return mdf


def _match_query(text: str, q: str) -> bool:
    tokens = [t.strip() for t in re.split(r"\s+", q) if t.strip()]
    text_low = text.lower()
    return all(t.lower() in text_low for t in tokens)


def filter_events(df: pd.DataFrame, query: str, days: List[str], rooms: List[str], include_main: bool) -> pd.DataFrame:
    out = df.copy()
    if not include_main:
        out = out[out["kind"] != "main_schedule"]
    if days:
        out = out[out["day"].isin(days)]
    if rooms:
        out = out[out["room"].isin(rooms)]

    if query.strip():
        q = query.strip()
        blob = (
            out["code"].fillna("") + " "
            + out["title"].fillna("") + " "
            + out["speaker"].fillna("") + " "
            + out["affiliation"].fillna("") + " "
            + out["room"].fillna("") + " "
            + out["source_sheet"].fillna("") + " "
            + out["session"].fillna("")
        )
        out = out[blob.map(lambda s: _match_query(s, q))]

    return out.sort_values(["start_dt", "room", "code"], na_position="last").reset_index(drop=True)


def events_from_selected(df_all: pd.DataFrame, selected_keys: Set[str]) -> pd.DataFrame:
    out = df_all[df_all["key"].isin(list(selected_keys))].copy()
    return out.sort_values(["start_dt", "room", "code"], na_position="last").reset_index(drop=True)


def add_conflict_flags(selected_df: pd.DataFrame) -> pd.DataFrame:
    if selected_df is None or len(selected_df) == 0:
        return selected_df

    df = selected_df.copy()
    df["conflict"] = False

    non_poster = df[df["kind"] != "poster"].copy()
    if len(non_poster) == 0:
        return df

    for day in non_poster["day"].dropna().unique().tolist():
        sub = non_poster[non_poster["day"] == day].sort_values(["start_dt", "end_dt"]).copy()
        if len(sub) <= 1:
            continue

        active_end = None
        active_idx = None

        for idx, r in sub.iterrows():
            s = r["start_dt"]
            e = r["end_dt"]

            if active_end is None:
                active_end = e
                active_idx = idx
                continue

            if s < active_end:
                df.loc[idx, "conflict"] = True
                if active_idx is not None:
                    df.loc[active_idx, "conflict"] = True
                if e > active_end:
                    active_end = e
                    active_idx = idx
            else:
                active_end = e
                active_idx = idx

    return df


def mark_conflict_with_selected(candidates: pd.DataFrame, selected: pd.DataFrame) -> pd.DataFrame:
    out = candidates.copy()
    out["conflict_with_selected"] = False

    if out is None or len(out) == 0 or selected is None or len(selected) == 0:
        return out

    sel_basis = selected[selected["kind"] != "poster"].copy()
    if len(sel_basis) == 0:
        return out

    sel_by_day: Dict[str, List[Tuple[dt.datetime, dt.datetime, str]]] = {}
    for _, r in sel_basis.iterrows():
        day = str(r.get("day", ""))
        sdt = r.get("start_dt")
        edt = r.get("end_dt")
        key = str(r.get("key", ""))
        if not day or pd.isna(sdt) or pd.isna(edt):
            continue
        sel_by_day.setdefault(day, []).append((sdt, edt, key))

    for i, r in out.iterrows():
        if str(r.get("kind", "")) == "poster":
            continue

        day = str(r.get("day", ""))
        sdt = r.get("start_dt")
        edt = r.get("end_dt")
        key = str(r.get("key", ""))
        if not day or pd.isna(sdt) or pd.isna(edt):
            continue

        intervals = sel_by_day.get(day, [])
        conflict = False
        for ss, ee, skey in intervals:
            if skey == key:
                continue
            if sdt < ee and ss < edt:
                conflict = True
                break
        out.loc[i, "conflict_with_selected"] = conflict

    return out


def df_for_picker(df: pd.DataFrame, selected_keys: Set[str], show_conflict_with_selected: bool = True) -> pd.DataFrame:
    cols = ["key", "display_date", "time", "room", "code", "title", "speaker", "session", "affiliation", "where"]
    if "conflict_with_selected" in df.columns and show_conflict_with_selected:
        cols.insert(1, "conflict_with_selected")

    show = df[cols].copy()
    show.insert(0, "選取", show["key"].map(lambda k: k in selected_keys))

    if "conflict_with_selected" in show.columns:
        show["conflict_with_selected"] = show["conflict_with_selected"].map(lambda x: "⚠️" if bool(x) else "")

    show = show.drop(columns=["key"])
    show = show.rename(
        columns={
            "conflict_with_selected": "衝突",
            "display_date": "日期",
            "time": "時間",
            "room": "教室/分會場",
            "code": "編號",
            "title": "投稿題目/演講主題",
            "speaker": "作者/講者/主持",
            "session": "主題領域",
            "affiliation": "單位",
            "where": "地點",
        }
    )
    return show


def build_ics(events: pd.DataFrame, cal_name: str = "Animal Behavior Workshop 2026") -> str:
    def fmt_dt(d: dt.datetime) -> str:
        return d.strftime("%Y%m%dT%H%M%S")

    def ics_escape(s: str) -> str:
        if s is None:
            return ""
        s = str(s)
        s = s.replace("\\", "\\\\")
        s = s.replace("\n", "\\n")
        s = s.replace(",", "\\,")
        s = s.replace(";", "\\;")
        return s

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Yuju//ABW2026//EN",
        f"X-WR-CALNAME:{ics_escape(cal_name)}",
        "CALSCALE:GREGORIAN",
    ]

    for _, r in events.iterrows():
        uid = re.sub(r"[^A-Za-z0-9]", "", str(r.get("key", "")))[:40] + "@abw2026"

        kind = (r.get("kind") or "").strip()
        code = (r.get("code") or "").strip() if r.get("code") else ""
        title = (r.get("title") or "").strip() if r.get("title") else ""
        speaker = (r.get("speaker") or "").strip() if r.get("speaker") else ""
        room = (r.get("where") or r.get("room") or "").strip()
        affiliation = (r.get("affiliation") or "").strip() if r.get("affiliation") else ""
        area = (r.get("session") or "").strip() if r.get("session") else ""

        if code and title:
            summary = f"{code}｜{title}"
        else:
            summary = title or code or ("Poster" if kind == "poster" else "Event")

        desc_parts = []
        if kind == "poster" and area:
            desc_parts.append(f"主題領域: {area}")
        if code:
            desc_parts.append(f"Code: {code}")
        if speaker:
            desc_parts.append(f"Speaker/Author: {speaker}")
        if affiliation:
            desc_parts.append(f"Affiliation: {affiliation}")
        if room:
            desc_parts.append(f"Room: {room}")
        description = "\\n".join(desc_parts) if desc_parts else ""

        lines.extend(
            [
                "BEGIN:VEVENT",
                f"UID:{ics_escape(uid)}",
                f"DTSTART:{fmt_dt(r['start_dt'])}",
                f"DTEND:{fmt_dt(r['end_dt'])}",
                f"SUMMARY:{ics_escape(summary)}",
                f"LOCATION:{ics_escape(room)}",
                f"DESCRIPTION:{ics_escape(description)}",
                "END:VEVENT",
            ]
        )

    lines.append("END:VCALENDAR")
    return "\n".join(lines)


def _as_set(x: Any) -> Set[str]:
    if x is None:
        return set()
    if isinstance(x, set):
        return set(map(str, x))
    if isinstance(x, (list, tuple)):
        return set(map(str, x))
    return set()


# ============================================================
# PDF helpers (preload + iframe + iOS reload)
# ============================================================

@st.cache_data(show_spinner=False)
def load_default_pdf_bytes(path: str) -> Optional[bytes]:
    try:
        if path and os.path.exists(path):
            with open(path, "rb") as f:
                return f.read()
    except Exception:
        return None
    return None


@st.cache_data(show_spinner=False)
def make_pdf_data_uri(pdf_bytes: bytes) -> str:
    b64 = base64.b64encode(pdf_bytes).decode("utf-8")
    return f"data:application/pdf;base64,{b64}"


def pdf_iframe_html(src: str, height: int = 650) -> str:
    return f'<iframe src="{src}" width="100%" height="{int(height)}" style="border: 1px solid rgba(49,51,63,0.15); border-radius: 8px;"></iframe>'


def build_pdf_src(pdf_url: str, pdf_data_uri: Optional[str], page: Optional[int]) -> Optional[str]:
    base = None
    if pdf_data_uri:
        base = pdf_data_uri
    elif pdf_url and pdf_url.strip():
        base = pdf_url.strip()
    else:
        return None

    p = int(page) if page and int(page) > 0 else 1

    # ✅ cache-buster：逼 iOS Safari 重新載入 iframe
    nonce = int(time.time() * 1000)

    if "#" in base:
        return f"{base}&page={p}&_={nonce}"
    else:
        return f"{base}#page={p}&_={nonce}"


# ============================================================
