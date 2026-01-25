# app_animal_behavior_2026_oauth_A_full_v2_2_2_abstract_mount.py
# ------------------------------------------------------------
# 版本變更說明（覆蓋版｜v2.2.2：摘要集 PDF 預先掛載 + 自動回填摘要頁碼）
# 1) ✅ 保留你現有 Excel 解析（大會議程/分會場/海報）、衝突規則、.ics 匯出、原始分頁 tabs。
# 2) ✅ 保留 Google OAuth 登入（openid / email / profile）+ SQLite 保存選取狀態。
# 3) ✅ 新增「摘要集 PDF（預先掛載）」支援：
#    - 預設路徑：data/2026 動物行為研討會摘要集.pdf（放在 GitHub repo 內）
#    - 啟動時一次掃描 PDF 逐頁文字，抽出 token 建 code->page 對照（單次 pass）
#    - 在搜尋結果表格/手機卡片中顯示「摘要頁 p.X」
#    - 提供「下載摘要 PDF」按鈕（直接下載 repo 內檔案）
# 4) ✅ 若缺 PDF 或缺 pymupdf 套件：顯示 warning，不影響議程工具運作。
#
# ⚠️ Streamlit Cloud 注意
# - SQLite（user_state.db）可能在重啟/重新部署後被重置；要跨重啟永久保存請換外部 DB。
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

APP_TITLE = "動物行為研討會 2026｜議程搜尋＋個人化行事曆"
DEFAULT_EXCEL_PATH = "2026 動行議程.xlsx"

# ✅ 摘要集 PDF：預設 repo 內掛載路徑（你要把 PDF 放在這裡）
DEFAULT_ABSTRACT_PDF_PATH = os.path.join("data", "2026 動物行為研討會摘要集.pdf")
# ✅ 摘要索引：只從第 34 頁開始（人類頁碼）
ABSTRACT_INDEX_START_PAGE_1IDX = 34

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

# Optional PDF dependency (PyMuPDF)
try:
    import fitz  # pymupdf
    _PDF_LIBS_OK = True
except Exception:
    _PDF_LIBS_OK = False


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
    # constant-time compare
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
    flow = Flow.from_client_config(client_config, scopes=scopes, redirect_uri=config["redirect_uri"])
    return flow


def auth_ui_sidebar() -> Optional[AuthUser]:
    """Sidebar auth UI. Return AuthUser if logged in, else None."""
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
        st.link_button("用 Google 登入（記住我的選擇）", auth_url, use_container_width=True)
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
    if x is None or (isinstance(x, float) and pd.isna(x)) or (isinstance(x, pd._libs.missing.NAType)):  # type: ignore
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
    """Find the first matching column by *candidate priority* (cand-first), not by sheet column order."""
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

        # Speaker column preference for certain sheets
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
    mdf["who"] = mdf["speaker"].fillna("")
    mdf["where"] = mdf["location"].fillna(mdf["room"])
    mdf["what"] = mdf["title"].fillna("")

    mdf["key"] = (
        mdf["date"].astype(str)
        + "|" + mdf["room"].astype(str)
        + "|" + mdf["start"].astype(str)
        + "|" + mdf["end"].astype(str)
        + "|" + mdf["code"].fillna("").astype(str)
        + "|" + mdf["title"].fillna("").astype(str)
    )

    # ✅ 預留摘要頁欄位（稍後在 main() 中依 PDF 自動回填）
    if "abstract_page" not in mdf.columns:
        mdf["abstract_page"] = pd.NA

    return mdf


# ============================================================
# Abstract PDF indexing (pre-mounted)
# ============================================================

def _normalize_code(x: Any) -> Optional[str]:
    s = _safe_str(x)
    if not s:
        return None
    return s.strip().upper().replace("｜", "|")


@st.cache_resource(show_spinner=True)
def build_abstract_code_to_page_map(pdf_path: str) -> Dict[str, int]:
    """
    Single-pass scan through PDF pages:
    - extract text per page
    - find "token-like" codes (e.g., PA01, PB12, O101, S203, etc.)
    - map token -> earliest page (1-based)
    """
    if (not _PDF_LIBS_OK) or (not os.path.exists(pdf_path)):
        return {}

    doc = fitz.open(pdf_path)

    # 規則：抽「看起來像編號」的 token
    # 1) 海報常見：P[A-Z]\d{2}  (PA01)
    # 2) 口頭/分組常見：A-Z 1-2字 + 2-4數字 (O101 / S101 / E102 / A12...)
    pat = re.compile(r"\b(P[A-Z]\d{2}|[A-Z]{1,2}\d{2,4})\b")

    m: Dict[str, int] = {}
    start_i = max(0, ABSTRACT_INDEX_START_PAGE_1IDX - 1)  # 34頁(1-index) -> 33(0-index)
    for i in range(start_i, doc.page_count):

        text = doc.load_page(i).get_text("text") or ""
        text = re.sub(r"\s+", " ", text).strip()
        if not text:
            continue
        toks = pat.findall(text)
        if not toks:
            continue
        page_no = i + 1
        for t in toks:
            tt = (t or "").strip().upper()
            if not tt:
                continue
            if tt not in m:
                m[tt] = page_no
    return m


def attach_abstract_page(df_all: pd.DataFrame, code2page: Dict[str, int]) -> pd.DataFrame:
    if df_all is None or len(df_all) == 0:
        return df_all
    if not code2page:
        df_all = df_all.copy()
        df_all["abstract_page"] = pd.NA
        return df_all

    df_all = df_all.copy()
    df_all["code_norm"] = df_all["code"].map(_normalize_code)
    df_all["abstract_page"] = df_all["code_norm"].map(lambda c: code2page.get(c) if c else None)
    df_all = df_all.drop(columns=["code_norm"], errors="ignore")
    return df_all


# ============================================================
# Search / filtering
# ============================================================

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
    # ✅ 新增 abstract_page 顯示
    cols = ["key", "display_date", "time", "room", "code", "abstract_page", "title", "speaker", "session", "affiliation", "where"]
    if "conflict_with_selected" in df.columns and show_conflict_with_selected:
        cols.insert(1, "conflict_with_selected")

    show = df[cols].copy()
    show.insert(0, "選取", show["key"].map(lambda k: k in selected_keys))

    if "conflict_with_selected" in show.columns:
        show["conflict_with_selected"] = show["conflict_with_selected"].map(lambda x: "⚠️" if bool(x) else "")

    # 摘要頁格式化
    show["abstract_page"] = show["abstract_page"].map(lambda x: (f"p.{int(x)}" if (x is not None and not pd.isna(x)) else "—"))

    show = show.drop(columns=["key"])
    show = show.rename(
        columns={
            "conflict_with_selected": "衝突",
            "display_date": "日期",
            "time": "時間",
            "room": "教室/分會場",
            "code": "編號",
            "abstract_page": "摘要頁",
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
        ap = r.get("abstract_page", None)

        if code and title:
            summary = f"{code}｜{title}"
        else:
            summary = title or code or ("Poster" if kind == "poster" else "Event")

        desc_parts = []
        if kind == "poster" and area:
            desc_parts.append(f"主題領域: {area}")
        if code:
            desc_parts.append(f"Code: {code}")
        if ap is not None and (not pd.isna(ap)):
            try:
                desc_parts.append(f"Abstract page: p.{int(ap)}")
            except Exception:
                pass
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


def main():
    db_init()
    st.title(APP_TITLE)

    # --- Auth panel (visible on mobile too) ---
    with st.expander("狀態保存（Google 登入）", expanded=False):
        user = auth_ui_sidebar()

        err = st.session_state.get("auth_error")
        if err:
            st.error(err)

        if user is None:
            if get_oauth_config() is None:
                st.warning("尚未設定 Google OAuth secrets；目前只能匿名模式（重整/跳掉可能會遺失勾選）。")
            if not _GOOGLE_LIBS_OK:
                st.warning("缺少 google-auth / google-auth-oauthlib，無法啟用登入。")
        else:
            c1, c2 = st.columns([1, 3])
            with c1:
                if user.picture:
                    st.image(user.picture, width=48)
            with c2:
                st.write(f"**{user.name or '已登入'}**")
                st.caption(user.email or "（email 未提供）")
            logout_ui()

        st.markdown("---")
        st.caption("🔒 登入僅用於記住你勾選的議程，不讀 Gmail、不改 Google Calendar。")

    query = ""
    include_main = True
    days = ["D1", "D2"]
    rooms: List[str] = []

    # ✅ 搜尋欄：獨立區塊（放在登入之後）
    st.markdown("### 🔎 搜尋")
    qcol1, qcol2, qcol3 = st.columns([0.62, 0.22, 0.16])
    
    with qcol1:
        query = st.text_input("關鍵字（可輸入多個詞，空格=AND）", value="", placeholder="例：社會學習  或  PA01  或  講者名字")
    
    with qcol2:
        days = st.multiselect("日期", options=["D1", "D2"], default=["D1", "D2"])
    
    with qcol3:
        include_main = st.checkbox("含大會", value=True, help="包含『大會議程』主表事件（報到/開幕等）")
    
    st.markdown("---")

    
    
    
    # --- Persistent state manager ---
    mgr = UserStateManager(st.session_state.get("auth_user"))
    st.session_state.setdefault("force_mobile_mode", bool(mgr.get("force_mobile_mode",True)))
    st.session_state.setdefault("selected_keys", _as_set(mgr.get("selected_keys", [])))
    st.session_state.setdefault("marked_delete_keys", _as_set(mgr.get("marked_delete_keys", [])))
    st.session_state.setdefault("confirm_delete_marked", bool(mgr.get("confirm_delete_marked", False)))

    # --- Mobile toggle ---
    tcol1, tcol2 = st.columns([0.75, 0.25])
    with tcol2:
        st.session_state.force_mobile_mode = st.toggle("Mobile mode", value=bool(st.session_state.force_mobile_mode))
    is_mobile = bool(st.session_state.force_mobile_mode)

    # --- Abstract PDF panel (pre-mounted) ---
    with st.expander("摘要集 PDF（預先掛載）", expanded=not is_mobile):
        st.caption(f"預設路徑：`{DEFAULT_ABSTRACT_PDF_PATH}`（請把 PDF 放在 repo 的 data/ 目錄）")
        if not _PDF_LIBS_OK:
            st.warning("尚未安裝 pymupdf（import fitz 失敗）。請在 requirements.txt 加上 `pymupdf`。")
        elif not os.path.exists(DEFAULT_ABSTRACT_PDF_PATH):
            st.warning("找不到摘要集 PDF。請確認 repo 內存在 data/2026 動物行為研討會摘要集.pdf")
        else:
            st.success("已偵測到摘要集 PDF（可自動回填摘要頁碼）。")
            try:
                st.download_button(
                    "下載摘要集 PDF",
                    data=open(DEFAULT_ABSTRACT_PDF_PATH, "rb").read(),
                    file_name=os.path.basename(DEFAULT_ABSTRACT_PDF_PATH),
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as e:
                st.warning(f"PDF 讀取失敗（仍可運作議程功能）：{e}")

    uploaded = None
    use_default = True
    query = ""
    include_main = True
    days = ["D1", "D2"]
    rooms: List[str] = []

    if is_mobile:
        with st.expander("控制面板（檔案/搜尋/篩選）", expanded=False):
            st.markdown("### 輸入議程檔案")
            uploaded = st.file_uploader("上傳 Excel（.xlsx）", type=["xlsx"])
            use_default = st.checkbox("使用預設檔案路徑（已掛載）", value=(uploaded is None))
            st.caption("預設檔案：" + DEFAULT_EXCEL_PATH)

    else:
        with st.sidebar:
            st.markdown("### 輸入議程檔案")
            uploaded = st.file_uploader("上傳 Excel（.xlsx）", type=["xlsx"])
            use_default = st.checkbox("使用預設檔案路徑（已掛載）", value=(uploaded is None))
            st.caption("預設檔案：" + DEFAULT_EXCEL_PATH)


    file_bytes: Optional[bytes] = None
    if uploaded is not None:
        file_bytes = uploaded.getvalue()
    elif use_default:
        try:
            with open(DEFAULT_EXCEL_PATH, "rb") as f:
                file_bytes = f.read()
        except Exception as e:
            st.error(f"讀取預設檔案失敗：{e}")

    if not file_bytes:
        st.info("請上傳 Excel 檔，或勾選使用預設檔案。")
        st.stop()

    # ⬇⬇⬇ 這裡「完全不要多縮排」
    sheets = load_excel_all_sheets(file_bytes)
    df_all = build_master_df(sheets)

    # ✅ 場地 / 分會場篩選（往上移）
    all_rooms = sorted(df_all["room"].dropna().unique().tolist())
    if is_mobile:
        with st.expander("教室/分會場篩選（可選）", expanded=False):
            rooms = st.multiselect("教室/分會場", options=all_rooms, default=[])
    else:
        with st.sidebar:
            rooms = st.multiselect("教室/分會場", options=all_rooms, default=[])


    # ✅ 在這裡把摘要頁碼回填進 df_all（依 code）
    code2page: Dict[str, int] = {}
    if _PDF_LIBS_OK and os.path.exists(DEFAULT_ABSTRACT_PDF_PATH):
        try:
            code2page = build_abstract_code_to_page_map(DEFAULT_ABSTRACT_PDF_PATH)
        except Exception as e:
            st.warning(f"摘要集索引建立失敗（不影響議程功能）：{e}")
            code2page = {}
    df_all = attach_abstract_page(df_all, code2page)

    

    selected_keys: Set[str] = set(st.session_state["selected_keys"])
    marked_delete: Set[str] = set(st.session_state["marked_delete_keys"])

    selected_df = add_conflict_flags(events_from_selected(df_all, selected_keys))

    df_hit = filter_events(df_all, query=query, days=days, rooms=rooms, include_main=include_main)
    df_hit2 = mark_conflict_with_selected(df_hit, selected_df)

    # ----------------------------
    # 1) 搜尋結果
    # ----------------------------
    st.subheader("1) 搜尋結果（加入／移除個人行事曆）")
    st.caption(f"符合筆數：{len(df_hit2)}（⚠️ 表示會與你已選的『非海報』行程時間重疊；海報不標衝突）")

    if not is_mobile:
        picker_df = df_for_picker(df_hit2, selected_keys, show_conflict_with_selected=True)

        edited = st.data_editor(
            picker_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "選取": st.column_config.CheckboxColumn("選取", help="勾選加入個人化行事曆"),
                "衝突": st.column_config.TextColumn("衝突", width="small", help="⚠️ 表示會與已選（非海報）行程撞期；海報不標"),
                "摘要頁": st.column_config.TextColumn("摘要頁", width="small", help="若摘要集 PDF 命中編號，顯示 p.X"),
                "投稿題目/演講主題": st.column_config.TextColumn(width="large"),
                "作者/講者/主持": st.column_config.TextColumn(width="medium"),
                "主題領域": st.column_config.TextColumn(width="medium"),
                "單位": st.column_config.TextColumn(width="medium"),
            },
            disabled=[
                "衝突", "日期", "時間", "教室/分會場", "編號",
                "摘要頁",
                "投稿題目/演講主題", "作者/講者/主持", "主題領域", "單位", "地點",
            ],
            key="editor_results",
        )

        hit_keys = df_hit2["key"].tolist()
        new_selected = set(selected_keys)
        for i, row in edited.iterrows():
            k = hit_keys[i]
            if bool(row["選取"]):
                new_selected.add(k)
            else:
                new_selected.discard(k)

        selected_keys = new_selected
        st.session_state["selected_keys"] = selected_keys

        c1, c2, c3 = st.columns([0.22, 0.22, 0.56])
        with c1:
            if st.button("全選（本頁）"):
                st.session_state["selected_keys"] = set(st.session_state["selected_keys"]).union(set(hit_keys))
                st.rerun()
        with c2:
            if st.button("全取消"):
                st.session_state["selected_keys"] = set()
                st.session_state["marked_delete_keys"] = set()
                st.session_state["confirm_delete_marked"] = False
                st.rerun()
        with c3:
            st.caption("提示：你可以先用關鍵字或教室篩選縮小範圍，再全選。")

    else:
        n_total = int(len(df_hit2))
        if n_total == 0:
            st.warning("沒有符合的結果：請放寬關鍵字/日期/教室篩選。")
            df_show = df_hit2
        elif n_total <= 10:
            st.caption(f"目前結果 {n_total} 筆（少於 10 筆，不顯示筆數滑桿）")
            df_show = df_hit2
        else:
            max_n = min(200, n_total)
            default_n = min(30, max_n)
            show_n = st.slider("顯示筆數", min_value=10, max_value=max_n, value=default_n, step=10)
            df_show = df_hit2.head(show_n).copy()

        for _, r in df_show.iterrows():
            k = str(r["key"])
            picked = (k in selected_keys)
            conflict_flag = "⚠️" if bool(r.get("conflict_with_selected")) else ""
            kind = str(r.get("kind") or "")

            ap = r.get("abstract_page", None)
            ap_txt = ""
            if ap is not None and (not pd.isna(ap)):
                try:
                    ap_txt = f"📄 摘要頁 p.{int(ap)}"
                except Exception:
                    ap_txt = ""

            with st.container(border=True):
                top = st.columns([0.74, 0.26])
                with top[0]:
                    st.markdown(f"**{r['day']} · {r['start']}–{r['end']} · {r['room']}**")
                    code = str(r.get("code") or "").strip()
                    title = str(r.get("title") or "").strip()
                    who = str(r.get("speaker") or "").strip()

                    if code:
                        st.markdown(f"{conflict_flag} **{code}**  {title}")
                    else:
                        st.markdown(f"{conflict_flag} {title}")
                    if who:
                        st.caption(who)
                    if ap_txt:
                        st.caption(ap_txt)
                    if kind == "poster":
                        st.caption("（Poster：不顯示衝突⚠️，也不計入衝突統計）")

                with top[1]:
                    if picked:
                        if st.button("移除", key=f"rm_{k}"):
                            selected_keys.discard(k)
                            marked_delete.discard(k)
                            st.session_state["selected_keys"] = selected_keys
                            st.session_state["marked_delete_keys"] = marked_delete
                            st.session_state["confirm_delete_marked"] = False
                            st.rerun()
                    else:
                        if st.button("加入", key=f"add_{k}"):
                            selected_keys.add(k)
                            st.session_state["selected_keys"] = selected_keys
                            st.rerun()

    selected_df = add_conflict_flags(events_from_selected(df_all, set(st.session_state["selected_keys"])))

    # ----------------------------
    # 2) 個人化行事曆（兩天）
    # ----------------------------
    st.markdown("---")
    st.subheader("2) 個人化行事曆（兩天）")

    d1_n = int((selected_df["day"] == "D1").sum()) if len(selected_df) else 0
    d2_n = int((selected_df["day"] == "D2").sum()) if len(selected_df) else 0
    conf_n = int(selected_df["conflict"].sum()) if len(selected_df) and "conflict" in selected_df.columns else 0

    m1, m2, m3 = st.columns(3)
    m1.metric("D1 已選", d1_n)
    m2.metric("D2 已選", d2_n)
    m3.metric("衝突場次（不含海報）", conf_n)

    if len(selected_df) == 0:
        st.info("尚未選取任何議程。")
    else:
        st.markdown("### 🗑️ 在行事曆清單中勾選刪除（勾選後會進待刪除清單）")
        st.caption("海報不計入衝突；衝突事件（非海報）會在清單中標示 ⚠️。")

        def _event_label(r: pd.Series) -> str:
            where = str(r.get("where") or r.get("room") or "").strip()
            code = str(r.get("code") or "").strip()
            title = str(r.get("title") or "").strip()
            ap = r.get("abstract_page", None)

            s = f"{r['start']}–{r['end']}｜{where}"
            if code:
                s += f"｜{code}"
            if ap is not None and (not pd.isna(ap)):
                try:
                    s += f"｜p.{int(ap)}"
                except Exception:
                    pass
            if title:
                s += f"｜{title[:40]}"
                if len(title) > 40:
                    s += "…"
            kind = str(r.get("kind") or "")
            conflict = bool(r.get("conflict")) if (kind != "poster") else False
            prefix = "⚠️ " if conflict else ""
            return prefix + s

        for day, label in [("D1", "D1｜2026-01-26"), ("D2", "D2｜2026-01-27")]:
            sub = selected_df[selected_df["day"] == day].copy().sort_values(["start_dt", "room", "code"])
            expand_default = bool((sub["conflict"].sum() > 0)) if len(sub) else False

            with st.expander(f"{label}（{len(sub)} 場）", expanded=expand_default):
                if len(sub) == 0:
                    st.caption("（此日尚未選取）")
                    continue

                for _, r in sub.iterrows():
                    k = str(r["key"])
                    checked = (k in st.session_state["marked_delete_keys"])
                    new_checked = st.checkbox(_event_label(r), value=checked, key=f"delchk_{day}_{k}")
                    if new_checked and (k not in st.session_state["marked_delete_keys"]):
                        st.session_state["marked_delete_keys"].add(k)
                        st.session_state["confirm_delete_marked"] = False
                    if (not new_checked) and (k in st.session_state["marked_delete_keys"]):
                        st.session_state["marked_delete_keys"].discard(k)
                        st.session_state["confirm_delete_marked"] = False

        st.divider()
        st.subheader("🗑️ 待刪除清單（已勾選）")

        marked_delete = set(st.session_state["marked_delete_keys"])
        marked_df = selected_df[selected_df["key"].isin(list(marked_delete))].copy().sort_values(["start_dt", "room"])

        if len(marked_df) == 0:
            st.caption("（目前沒有勾選任何待刪除行程）")
        else:
            for _, r in marked_df.iterrows():
                with st.container(border=True):
                    st.markdown(f"**{r['day']} · {r['start']}–{r['end']} · {r['room']}**")
                    code = str(r.get("code") or "").strip()
                    title = str(r.get("title") or "").strip()
                    ap = r.get("abstract_page", None)

                    line = ""
                    if code:
                        line += f"**{code}**  "
                    line += title
                    if ap is not None and (not pd.isna(ap)):
                        try:
                            line += f"  ·  📄 p.{int(ap)}"
                        except Exception:
                            pass
                    st.markdown(line)

                    who = str(r.get("speaker") or "").strip()
                    if who:
                        st.caption(who)

            st.divider()
            if not st.session_state["confirm_delete_marked"]:
                if st.button("刪除以上已勾選（需再次確認）", type="primary"):
                    st.session_state["confirm_delete_marked"] = True
                    st.rerun()
            else:
                st.error("再次確認：確定要把這些行程從『已選清單』移除嗎？（可之後再從搜尋結果重新加入）")
                b1, b2 = st.columns(2)
                if b1.button("確定刪除", type="primary"):
                    sel = set(st.session_state["selected_keys"])
                    md = set(st.session_state["marked_delete_keys"])
                    sel -= md
                    st.session_state["selected_keys"] = sel
                    st.session_state["marked_delete_keys"] = set()
                    st.session_state["confirm_delete_marked"] = False
                    st.rerun()
                if b2.button("取消"):
                    st.session_state["confirm_delete_marked"] = False
                    st.rerun()

        ics_text = build_ics(selected_df)
        st.download_button(
            "下載 .ics 行事曆檔（可匯入 Google/Apple Calendar）",
            data=ics_text.encode("utf-8"),
            file_name="animal_behavior_workshop_2026_selected.ics",
            mime="text/calendar",
        )

    # ----------------------------
    # 3) Raw sheets
    # ----------------------------
    st.markdown("---")
    st.subheader("3) 大會議程（Excel 原始分頁）")
    st.caption("下方直接呈現 Excel 每個分頁內容，便於核對。")

    tab_names = list(sheets.keys())
    tabs = st.tabs(tab_names)
    for name, tab in zip(tab_names, tabs):
        with tab:
            st.dataframe(sheets[name], use_container_width=True, hide_index=True)

    # ---- Persist core state (end of run) ----
    mgr.set("force_mobile_mode", bool(st.session_state.force_mobile_mode))
    mgr.set("selected_keys", sorted(list(set(map(str, st.session_state["selected_keys"])))))
    mgr.set("marked_delete_keys", sorted(list(set(map(str, st.session_state["marked_delete_keys"])))))
    mgr.set("confirm_delete_marked", bool(st.session_state["confirm_delete_marked"]))
    mgr.save()


if __name__ == "__main__":
    main()
