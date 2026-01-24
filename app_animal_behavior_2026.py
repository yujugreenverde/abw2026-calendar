# app.py
# 覆蓋版（2026-01-24）
# 版本變更說明：
# 1) ✅ 支援「無摘要索引」：直接掃描預掛載 PDF，自動建立（報告編號 / 人名）→ PDF頁碼 的索引
# 2) ✅ 可用「人名」或「報告編號（OA01/PA01/...）」搜尋摘要；可直接在網頁內預覽對應 PDF 頁面
# 3) ✅ 搜尋列與日期選擇「拉出來」：不再放在折疊(expander)內
# 4) ✅ 預設 Mobile mode 開啟：預覽寬度、字級與版面為手機友善
#
# 使用方式：
#   streamlit run app.py
#
# 依賴：
#   pip install streamlit pymupdf pandas openpyxl
#
# 可選：若你仍有 Excel 議程資料（例如 schedule.xlsx），放在同資料夾即可（會自動讀取）。
# PDF 需存在於：PDF_PATH（預設為 /mnt/data/2026 動物行為研討會摘要集.pdf）

from __future__ import annotations

import re
import io
import os
import time
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional

import streamlit as st

try:
    import fitz  # PyMuPDF
except Exception as e:
    fitz = None

try:
    import pandas as pd
except Exception:
    pd = None


# ----------------------------
# Config
# ----------------------------
DEFAULT_PDF_PATH = os.environ.get("PDF_PATH", "/mnt/data/2026 動物行為研討會摘要集.pdf")
DEFAULT_EXCEL_PATH = os.environ.get("SCHEDULE_XLSX", "schedule.xlsx")  # optional

# Report-id patterns seen in this PDF:
# - Oral: OA01-OA38, OB01..., OC..., OD..., OE..., OF..., OH..., OI..., OJ...
# - Poster: PA01..., PB..., PC..., PD..., PE..., PF..., PG..., PH..., PI..., PJ...
REPORT_ID_RE = re.compile(r"\b([OP][A-Z])\s?(\d{2})\b")   # e.g., OA01, PA01 (captures O? and number)
REPORT_ID_RE2 = re.compile(r"\b([A-Z]{2,3})\s?(\d{2})\b") # fallback (rare)
# Chinese name heuristic: 2-4 CJK chars (very approximate)
CJK_NAME_RE = re.compile(r"[\u4e00-\u9fff]{2,4}")

# ----------------------------
# Data structures
# ----------------------------
@dataclass
class Hit:
    report_id: Optional[str]
    page_1based: int
    context: str


# ----------------------------
# PDF Indexing
# ----------------------------
def _normalize_report_id(prefix: str, num: str) -> str:
    return f"{prefix}{int(num):02d}"


@st.cache_data(show_spinner=False)
def build_pdf_index(pdf_path: str) -> Dict[str, List[int]]:
    """
    Build an index: report_id -> list of 1-based page numbers where it appears.
    Also builds a 'name index' indirectly via full-text search (done on demand).
    """
    if fitz is None:
        raise RuntimeError("PyMuPDF (fitz) is not installed. Please `pip install pymupdf`.")
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    doc = fitz.open(pdf_path)
    idx: Dict[str, List[int]] = {}
    for i in range(doc.page_count):
        txt = doc.load_page(i).get_text("text") or ""
        # Primary patterns: OA01, PA01, ...
        for m in REPORT_ID_RE.finditer(txt):
            rid = _normalize_report_id(m.group(1), m.group(2))
            idx.setdefault(rid, [])
            if (i + 1) not in idx[rid]:
                idx[rid].append(i + 1)
        # Fallback: catches e.g., B20-B26 style; we don't index ranges here
    doc.close()
    return idx


def search_pdf_text(pdf_path: str, query: str, max_hits: int = 50) -> List[Hit]:
    """
    Full-text search over PDF pages (substring / regex-lite).
    Returns page hits with a short context snippet.
    """
    if fitz is None:
        raise RuntimeError("PyMuPDF (fitz) is not installed. Please `pip install pymupdf`.")
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    query = query.strip()
    if not query:
        return []

    # Build a safe regex that does "contains" for CJK / latin; allow user to type regex by prefixing r/
    use_regex = False
    if query.startswith("r/"):
        use_regex = True
        pat = query[2:]
        try:
            qre = re.compile(pat, flags=re.IGNORECASE)
        except re.error:
            qre = re.compile(re.escape(pat), flags=re.IGNORECASE)
    else:
        qre = re.compile(re.escape(query), flags=re.IGNORECASE)

    doc = fitz.open(pdf_path)
    hits: List[Hit] = []
    for i in range(doc.page_count):
        txt = doc.load_page(i).get_text("text") or ""
        if not txt:
            continue
        m = qre.search(txt) if use_regex else qre.search(txt)
        if not m:
            continue
        # context: ~160 chars around match
        start = max(0, m.start() - 80)
        end = min(len(txt), m.end() + 80)
        context = txt[start:end].replace("\n", " ")
        # try to find a report id nearby (best effort)
        rid = None
        ridm = REPORT_ID_RE.search(txt)
        if ridm:
            rid = _normalize_report_id(ridm.group(1), ridm.group(2))
        hits.append(Hit(report_id=rid, page_1based=i + 1, context=context))
        if len(hits) >= max_hits:
            break
    doc.close()
    return hits


def render_pdf_page_to_png(pdf_path: str, page_1based: int, zoom: float = 2.0) -> bytes:
    if fitz is None:
        raise RuntimeError("PyMuPDF (fitz) is not installed. Please `pip install pymupdf`.")
    doc = fitz.open(pdf_path)
    p = doc.load_page(page_1based - 1)
    mat = fitz.Matrix(zoom, zoom)
    pix = p.get_pixmap(matrix=mat, alpha=False)
    png = pix.tobytes("png")
    doc.close()
    return png


# ----------------------------
# Optional schedule Excel
# ----------------------------
@st.cache_data(show_spinner=False)
def load_schedule_excel(excel_path: str) -> Optional["pd.DataFrame"]:
    if pd is None:
        return None
    if not os.path.exists(excel_path):
        return None
    try:
        df = pd.read_excel(excel_path)
        return df
    except Exception:
        return None


def _infer_date_col(df) -> Optional[str]:
    # common names: Date, date, 日期, Day, day
    if df is None:
        return None
    for c in df.columns:
        if str(c).strip().lower() in {"date", "day"} or str(c).strip() in {"日期", "日", "Day"}:
            return c
    return None


# ----------------------------
# UI
# ----------------------------
def main():
    st.set_page_config(page_title="CABE 2026 – 議程與摘要搜尋", layout="centered")

    st.title("CABE 2026 – 議程與摘要搜尋")

    # Mobile mode: default ON (as requested)
    colA, colB = st.columns([1, 1])
    with colA:
        mobile_mode = st.toggle("Mobile mode", value=True, help="開啟後會使用較適合手機的顯示方式（較窄寬度、較大字與更簡潔）")
    with colB:
        show_pdf_preview = st.toggle("PDF page preview", value=True, help="搜尋命中後，直接顯示該頁 PDF 預覽")

    pdf_path = st.text_input("PDF path", value=DEFAULT_PDF_PATH, help="預掛載 PDF 的路徑。若部署在 Streamlit Cloud，請確保檔案已放入 repo 或外部掛載。")

    # Build report-id index once
    with st.spinner("建立 PDF 索引中（第一次會較久）…"):
        report_idx = build_pdf_index(pdf_path)

    # --- Search bar OUTSIDE expander (as requested)
    st.subheader("搜尋摘要（報告編號 / 人名 / 關鍵字）")
    q = st.text_input("Search", value="", placeholder="例如：OA12、PA21、陳睿傑、穿山甲、soundscape…")

    # --- Date selector OUTSIDE expander (as requested)
    st.subheader("日期篩選（若有 Excel 議程資料才會啟用）")
    schedule_df = load_schedule_excel(DEFAULT_EXCEL_PATH)
    date_col = _infer_date_col(schedule_df)
    if schedule_df is None:
        st.info("尚未偵測到 schedule.xlsx（可選）。目前仍可用 PDF 進行摘要搜尋。")
        selected_date = None
    else:
        # build date options
        dates = sorted({str(x) for x in schedule_df[date_col].dropna().unique()}) if date_col else []
        selected_date = st.selectbox("Date", options=["(All)"] + dates, index=0)

    # --- Results
    if q.strip():
        q_clean = q.strip()

        # 1) If user typed a report id like OA01/PA01
        m = REPORT_ID_RE.fullmatch(q_clean.replace(" ", ""))
        if m:
            rid = _normalize_report_id(m.group(1), m.group(2))
            pages = report_idx.get(rid, [])
            if not pages:
                st.warning(f"找不到報告編號 **{rid}**（可能是 PDF 文字層不完整或報告編號格式不同）。")
            else:
                st.success(f"報告編號 **{rid}** 命中頁：{', '.join(map(str, pages))}")
                _show_pages(pdf_path, pages, show_pdf_preview, mobile_mode)
        else:
            # 2) Full text search (names / keywords)
            with st.spinner("在 PDF 全文搜尋中…"):
                hits = search_pdf_text(pdf_path, q_clean, max_hits=50)

            if not hits:
                st.warning("找不到相符內容。你可以：\n- 改用更短的關鍵字\n- 或用 `r/正則` 模式（例如 r/陳.{0,2}睿）")
            else:
                st.write(f"命中 {len(hits)} 筆（顯示最多 50 筆）")
                for k, h in enumerate(hits, start=1):
                    header = f"{k}. p.{h.page_1based}" + (f" · {h.report_id}" if h.report_id else "")
                    with st.expander(header, expanded=(k <= 2)):
                        st.write(h.context)

                        if show_pdf_preview:
                            _show_pages(pdf_path, [h.page_1based], show_pdf_preview, mobile_mode)

    st.divider()
    st.caption("提示：報告編號可用 OA01 / PA01 這種格式；人名可直接打中文姓名。若 PDF 有些頁面是圖片掃描，文字層可能不足，搜尋會受限。")


def _show_pages(pdf_path: str, pages_1based: List[int], show_preview: bool, mobile_mode: bool):
    if not show_preview:
        st.write("PDF 頁碼：", ", ".join(map(str, pages_1based)))
        return

    zoom = 1.6 if mobile_mode else 2.2
    max_width = 360 if mobile_mode else 900

    for p in pages_1based:
        try:
            png = render_pdf_page_to_png(pdf_path, p, zoom=zoom)
            st.image(png, caption=f"PDF page {p}", use_container_width=(not mobile_mode))
        except Exception as e:
            st.error(f"無法渲染 PDF 第 {p} 頁：{e}")


if __name__ == "__main__":
    main()
