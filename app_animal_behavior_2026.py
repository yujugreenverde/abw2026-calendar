# app_animal_behavior_2026.py
# ------------------------------------------------------------
# 版本變更說明（B1：格狀行事曆點選=勾選刪除；Mobile-first 版面）
# 1) ✅ 保留你現有的 Excel 解析邏輯（大會議程/分會場/海報）、衝突規則、.ics 匯出、原始分頁 tabs。
# 2) 📱 Mobile-first：手機改用「上方控制面板 expander」，桌機維持 sidebar。
# 3) 📱 Mobile 搜尋結果改為「卡片式清單 + 加入/移除」；桌機維持 data_editor。
# 4) 🗓️（B1）新增「已選行程：格狀行事曆視圖」
#    - 在行事曆上點事件 => 事件加入/移出「待刪除清單」
#    - 事件標題前加上 🗑️ 表示已勾選待刪除
#    - 下方提供「二次確認」後批次刪除（從 selected_keys 移除）
#    - ⚠️ 衝突事件（非海報）在行事曆標題前加 ⚠️（海報不標衝突）
# 5) 若環境沒有 streamlit-calendar：自動 fallback 成「待刪除清單（無格狀）」仍可刪除，不會整個壞。
#
# Usage:
#   streamlit run app_animal_behavior_2026.py

from __future__ import annotations

import re
import io
import datetime as dt
from typing import Dict, Tuple, Optional, List, Set

import pandas as pd
import streamlit as st

APP_TITLE = "動物行為研討會 2026｜議程搜尋＋個人化行事曆"
DEFAULT_EXCEL_PATH = "2026 動行議程.xlsx"

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
/* Calendar event text visibility (best effort; depends on component CSS) */
.fc .fc-event-title, .fc .fc-event-time { line-height: 1.25 !important; }
</style>
    """,
    unsafe_allow_html=True,
)

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

    # ---- 大會議程（主表）----
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

    # ---- 其他 sheets ----
    for sheet_name, df0 in sheets.items():
        if sheet_name == "大會議程":
            continue

        # ---- 海報 ----
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

        # ---- 一般分會場 ----
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
        col_speaker = _find_col(cols, ["作者姓名", "講者", "主持人"])
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
    """
    同日時間重疊 => conflict=True。
    但：kind == 'poster' 的事件不參與衝突偵測（也不會被標紅）。
    """
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


# ----------------------------
# Calendar (streamlit-calendar optional)
# ----------------------------
def _try_get_calendar():
    try:
        from streamlit_calendar import calendar
        return calendar
    except Exception:
        return None


def _event_title_for_calendar(r: pd.Series, marked_for_delete: Set[str]) -> str:
    # 核心顯示：地點｜編號（你之前喜好）
    where = str(r.get("where") or r.get("room") or "").strip()
    code = str(r.get("code") or "").strip()
    base = f"{where}｜{code}" if (where and code) else (code or where or "Event")

    # 衝突標記（海報不標）
    conflict = bool(r.get("conflict")) if "conflict" in r.index else False
    kind = str(r.get("kind") or "")
    prefix = ""
    if kind != "poster" and conflict:
        prefix += "⚠️ "

    # 待刪除標記
    if str(r.get("key")) in marked_for_delete:
        prefix = "🗑️ " + prefix

    return prefix + base


def _selected_to_calendar_events(selected_df: pd.DataFrame, marked_for_delete: Set[str]) -> List[Dict]:
    events: List[Dict] = []
    if selected_df is None or len(selected_df) == 0:
        return events
    for _, r in selected_df.iterrows():
        key = str(r.get("key"))
        events.append(
            dict(
                id=key,
                title=_event_title_for_calendar(r, marked_for_delete),
                start=pd.to_datetime(r["start_dt"]).isoformat(),
                end=pd.to_datetime(r["end_dt"]).isoformat(),
                extendedProps=dict(
                    key=key,
                    kind=str(r.get("kind") or ""),
                    title_full=str(r.get("title") or ""),
                    speaker=str(r.get("speaker") or ""),
                    room=str(r.get("room") or ""),
                    where=str(r.get("where") or ""),
                    code=str(r.get("code") or ""),
                ),
            )
        )
    return events


def _render_calendar_block(
    selected_df: pd.DataFrame,
    day_view: str,
    marked_for_delete: Set[str],
    height: int = 650,
) -> Tuple[Optional[str], Dict]:
    """
    Returns (clicked_event_id, raw_state)
    """
    cal = _try_get_calendar()
    if cal is None:
        return None, {"fallback": True}

    # filter by day_view
    if day_view in ("D1", "D2"):
        sdf = selected_df[selected_df["day"] == day_view].copy()
        init_date = DATE_MAP[day_view]
    else:
        sdf = selected_df.copy()
        init_date = pd.to_datetime(sdf["start_dt"]).min().date() if len(sdf) else DATE_MAP["D1"]

    events = _selected_to_calendar_events(sdf, marked_for_delete)

    options = {
        "initialView": "timeGridDay" if day_view in ("D1", "D2") else "timeGridWeek",
        "initialDate": init_date.isoformat(),
        "headerToolbar": {
            "left": "prev,next today",
            "center": "title",
            "right": "timeGridDay,timeGridWeek,listWeek",
        },
        "height": height,
        "slotMinTime": "07:00:00",
        "slotMaxTime": "21:30:00",
        "allDaySlot": False,
        "nowIndicator": True,
        "eventDisplay": "block",
        "eventTimeFormat": {"hour": "2-digit", "minute": "2-digit", "hour12": False},
        "expandRows": True,
        "slotEventOverlap": False,
        "eventMaxStack": 99,
    }

    state = cal(events=events, options=options, key=f"calendar_selected_{day_view}")
    clicked_id: Optional[str] = None

    # Try parse click payload (streamlit-calendar versions differ)
    if isinstance(state, dict):
        payload = None
        for k in ("eventClick", "event_click", "eventClickInfo"):
            if k in state and state[k]:
                payload = state[k]
                break
        if isinstance(payload, dict):
            if "event" in payload and isinstance(payload["event"], dict):
                clicked_id = payload["event"].get("id")
            else:
                clicked_id = payload.get("id")

    return clicked_id, state


# ----------------------------
# Mobile detection (manual toggle; Streamlit can't reliably auto-detect viewport)
# ----------------------------
if "force_mobile_mode" not in st.session_state:
    st.session_state.force_mobile_mode = False

# ----------------------------
# UI
# ----------------------------
st.title(APP_TITLE)

# top-right: mobile toggle
tcol1, tcol2 = st.columns([0.75, 0.25])
with tcol2:
    st.session_state.force_mobile_mode = st.toggle("Mobile mode", value=st.session_state.force_mobile_mode)

is_mobile = bool(st.session_state.force_mobile_mode)

# sidebar controls (desktop) or top expander (mobile)
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

        st.markdown("---")
        st.markdown("### 搜尋與篩選")
        query = st.text_input("關鍵字（可輸入多個詞，空格=AND）", value="")
        include_main = st.checkbox("包含『大會議程』的主表事件（報到/開幕等）", value=True)
        days = st.multiselect("日期", options=["D1", "D2"], default=["D1", "D2"])
else:
    with st.sidebar:
        st.markdown("### 輸入議程檔案")
        uploaded = st.file_uploader("上傳 Excel（.xlsx）", type=["xlsx"])
        use_default = st.checkbox("使用預設檔案路徑（已掛載）", value=(uploaded is None))
        st.caption("預設檔案：" + DEFAULT_EXCEL_PATH)

        st.markdown("---")
        st.markdown("### 搜尋與篩選")
        query = st.text_input("關鍵字（可輸入多個詞，空格=AND）", value="")
        include_main = st.checkbox("包含『大會議程』的主表事件（報到/開幕等）", value=True)
        days = st.multiselect("日期", options=["D1", "D2"], default=["D1", "D2"])

# load file bytes
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

sheets = load_excel_all_sheets(file_bytes)
df_all = build_master_df(sheets)

# rooms selector depends on df_all
all_rooms = sorted(df_all["room"].dropna().unique().tolist())

if is_mobile:
    with st.expander("教室/分會場篩選（可選）", expanded=False):
        rooms = st.multiselect("教室/分會場", options=all_rooms, default=[])
else:
    with st.sidebar:
        rooms = st.multiselect("教室/分會場", options=all_rooms, default=[])

# session state
if "selected_keys" not in st.session_state:
    st.session_state["selected_keys"] = set()
if "marked_delete_keys" not in st.session_state:
    st.session_state["marked_delete_keys"] = set()
if "confirm_delete_marked" not in st.session_state:
    st.session_state["confirm_delete_marked"] = False

selected_keys: Set[str] = set(st.session_state["selected_keys"])
marked_delete: Set[str] = set(st.session_state["marked_delete_keys"])

selected_df = events_from_selected(df_all, selected_keys)
selected_df = add_conflict_flags(selected_df)

df_hit = filter_events(df_all, query=query, days=days, rooms=rooms, include_main=include_main)
df_hit2 = mark_conflict_with_selected(df_hit, selected_df)

# ----------------------------
# 1) 搜尋結果
# ----------------------------
st.subheader("1) 搜尋結果（加入／移除個人行事曆）")
st.caption(f"符合筆數：{len(df_hit2)}（⚠️ 表示會與你已選的『非海報』行程時間重疊；海報不標衝突）")

if not is_mobile:
    # desktop: data_editor (your original)
    picker_df = df_for_picker(df_hit2, selected_keys, show_conflict_with_selected=True)

    edited = st.data_editor(
        picker_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "選取": st.column_config.CheckboxColumn("選取", help="勾選加入個人化行事曆"),
            "衝突": st.column_config.TextColumn("衝突", width="small", help="⚠️ 表示會與已選（非海報）行程撞期；海報不標"),
            "投稿題目/演講主題": st.column_config.TextColumn(width="large"),
            "作者/講者/主持": st.column_config.TextColumn(width="medium"),
            "主題領域": st.column_config.TextColumn(width="medium"),
            "單位": st.column_config.TextColumn(width="medium"),
        },
        disabled=[
            "衝突", "日期", "時間", "教室/分會場", "編號",
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
    # mobile: card list + add/remove
    # show first N with slider
    n_total = int(len(df_hit2))

    if n_total <= 10:
        show_n = n_total
        st.caption(f"目前結果 {n_total} 筆（少於 10 筆，不顯示筆數滑桿）")
    else:
        max_n = min(200, n_total)
        default_n = min(30, max_n)
        show_n = st.slider(
            "顯示筆數",
            min_value=10,
            max_value=max_n,
            value=default_n,
            step=10,
        )
    
    df_show = df_hit2.head(show_n).copy()


    for _, r in df_show.iterrows():
        k = str(r["key"])
        picked = (k in selected_keys)
        conflict_flag = "⚠️" if bool(r.get("conflict_with_selected")) else ""
        kind = str(r.get("kind") or "")

        with st.container(border=True):
            top = st.columns([0.74, 0.26])
            with top[0]:
                line1 = f"**{r['day']} · {r['start']}–{r['end']} · {r['room']}**"
                st.markdown(line1)
                code = str(r.get("code") or "").strip()
                title = str(r.get("title") or "").strip()
                who = str(r.get("speaker") or "").strip()
                if code:
                    st.markdown(f"{conflict_flag} **{code}**  {title}")
                else:
                    st.markdown(f"{conflict_flag} {title}")
                if who:
                    st.caption(who)
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

    st.caption("（手機模式下建議先用關鍵字/日期/教室縮小後再加入）")

# recompute selected_df after updates
selected_df = events_from_selected(df_all, set(st.session_state["selected_keys"]))
selected_df = add_conflict_flags(selected_df)

# ----------------------------
# 2) 個人化行事曆（兩天） + B1 行事曆點選刪除
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
    st.markdown("### 🗓️ 行事曆視圖（點事件 = 勾選待刪除）")
    st.caption("已勾選待刪除的事件會顯示 🗑️；按下方按鈕可批次刪除（會從你的已選清單移除）。")

    cal_day_view = st.radio("行事曆顯示", options=["All", "D1", "D2"], horizontal=True, index=0)

    clicked_id, cal_state = _render_calendar_block(
        selected_df=selected_df,
        day_view=cal_day_view,
        marked_for_delete=set(st.session_state["marked_delete_keys"]),
        height=560 if is_mobile else 720,
    )

    if isinstance(cal_state, dict) and cal_state.get("fallback"):
        st.warning("⚠️ 目前環境沒有 streamlit-calendar，因此無法顯示格狀行事曆；你仍可用下方『待刪除清單』勾選刪除。")

    # toggle mark delete by calendar click
    if clicked_id:
        md = set(st.session_state["marked_delete_keys"])
        if clicked_id in md:
            md.discard(clicked_id)
        else:
            md.add(clicked_id)
        st.session_state["marked_delete_keys"] = md
        st.session_state["confirm_delete_marked"] = False
        st.rerun()

    # Marked-for-delete panel
    st.divider()
    st.subheader("🗑️ 待刪除清單（由行事曆點選加入）")

    marked_delete = set(st.session_state["marked_delete_keys"])
    marked_df = selected_df[selected_df["key"].isin(list(marked_delete))].copy().sort_values(["start_dt", "room"])

    if len(marked_df) == 0:
        st.caption("（目前沒有勾選任何待刪除行程）")
    else:
        # compact list (mobile-friendly)
        for _, r in marked_df.iterrows():
            k = str(r["key"])
            with st.container(border=True):
                cL, cR = st.columns([0.78, 0.22])
                with cL:
                    st.markdown(f"**{r['day']} · {r['start']}–{r['end']} · {r['room']}**")
                    code = str(r.get("code") or "").strip()
                    title = str(r.get("title") or "").strip()
                    if code:
                        st.markdown(f"**{code}**  {title}")
                    else:
                        st.markdown(title)
                    who = str(r.get("speaker") or "").strip()
                    if who:
                        st.caption(who)
                with cR:
                    if st.button("取消", key=f"unmark_{k}"):
                        md = set(st.session_state["marked_delete_keys"])
                        md.discard(k)
                        st.session_state["marked_delete_keys"] = md
                        st.session_state["confirm_delete_marked"] = False
                        st.rerun()

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

    # ---- 原本 D1/D2 表格（保留：手機用 expander）----
    def _style_conflicts(df_view: pd.DataFrame) -> pd.io.formats.style.Styler:
        def row_style(r):
            if str(r.get("衝突", "")).strip() == "⚠️":
                return ["background-color: #ffe5e5; font-weight: 600;" for _ in r]
            return ["" for _ in r]
        return df_view.style.apply(row_style, axis=1)

    st.markdown("### 📋 已選清單（D1 / D2）")

    for day, label in [("D1", "D1｜2026-01-26"), ("D2", "D2｜2026-01-27")]:
        sub = selected_df[selected_df["day"] == day].copy()
        expand_default = bool((sub["conflict"].sum() > 0)) if len(sub) else False

        with st.expander(f"{label}（{len(sub)} 場）", expanded=expand_default):
            if len(sub) == 0:
                st.caption("（此日尚未選取）")
                continue

            view = sub[["start", "end", "room", "code", "title", "speaker", "session", "conflict", "kind"]].copy()
            view = view.rename(
                columns={
                    "start": "開始",
                    "end": "結束",
                    "room": "教室/類別",
                    "code": "編號",
                    "title": "主題",
                    "speaker": "講者/作者",
                    "session": "主題領域",
                    "conflict": "衝突",
                    "kind": "類型",
                }
            )
            view["衝突"] = view.apply(lambda r: ("⚠️" if bool(r["衝突"]) else ""), axis=1)

            st.dataframe(_style_conflicts(view.drop(columns=["類型"])), use_container_width=True, hide_index=True)

    # ---- ics 匯出 ----
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
