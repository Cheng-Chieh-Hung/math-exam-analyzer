import io
import re
from dataclasses import dataclass
from typing import List, Tuple, Optional

import pandas as pd
import pdfplumber
import streamlit as st


# -----------------------------
# Data structures
# -----------------------------
@dataclass
class ExamItem:
    order_index: int
    label: str
    section: str  # 單選 / 填充 / 非選 / 未知
    score: Optional[float]
    stem_preview: str


# -----------------------------
# Core functions
# -----------------------------
def extract_text_from_pdf(pdf_bytes: bytes) -> Tuple[str, List[str]]:
    """Extract selectable text from PDF and return full_text + per_page_text."""
    per_page = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            try:
                txt = page.extract_text() or ""
            except Exception:
                txt = ""
            per_page.append(txt)
    full_text = "\n\n".join(per_page)
    return full_text, per_page


def guess_exam_items(full_text: str) -> List[ExamItem]:
    """
    MVP heuristic: detect question anchors by lines starting with "1 ", "1.", "1、" ...
    Note: This is a coarse estimate. We'll improve later for 6(1), 三-2(2), etc.
    """
    lines = [ln.strip() for ln in full_text.splitlines()]
    anchors = []

    for i, ln in enumerate(lines):
        # Match: "1 " or "1." or "1、"
        m = re.match(r"^(\d{1,3})\s*[\.、]?\s+", ln)
        if m:
            anchors.append((i, m.group(1)))

    # De-duplicate consecutive same question number
    filtered = []
    last_qn = None
    for idx, qn in anchors:
        if qn != last_qn:
            filtered.append((idx, qn))
            last_qn = qn

    items: List[ExamItem] = []
    for k, (start_i, qn) in enumerate(filtered):
        end_i = filtered[k + 1][0] if k + 1 < len(filtered) else len(lines)
        block = " ".join([x for x in lines[start_i:end_i] if x])
        preview = block[:80] + ("…" if len(block) > 80 else "")
        items.append(
            ExamItem(
                order_index=k + 1,
                label=str(qn),
                section="未知",
                score=None,
                stem_preview=preview,
            )
        )
    return items


def parse_answer_string(ans: str) -> List[bool]:
    """'-' => correct(True), 'X'/'x' => wrong(False). Ignore other chars."""
    cleaned = [c for c in ans.strip() if c in ["-", "X", "x"]]
    return [c == "-" for c in cleaned]


def build_results_df(items: List[ExamItem], correctness: List[bool]) -> pd.DataFrame:
    n = min(len(items), len(correctness))
    rows = []
    for i in range(n):
        rows.append(
            {
                "order_index": items[i].order_index,
                "label": items[i].label,
                "is_correct": correctness[i],
                "result": "對" if correctness[i] else "錯",
                "stem_preview": items[i].stem_preview,
            }
        )
    return pd.DataFrame(rows)


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "attempt") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


def file_signature(uploaded_file) -> str:
    """
    Fast signature to detect if the uploaded file changed.
    For MVP: name + size is enough.
    """
    return f"{uploaded_file.name}:{uploaded_file.size}"


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="數學考卷分析 MVP", layout="wide")
st.title("數學考卷分析 MVP（可選字 PDF）")

st.markdown(
    """
此 MVP 目前提供：
- 上傳 **可選字 PDF** → 抽取文字（只在換檔時解析一次，避免卡住）
- 粗略偵測作答點（題號 anchor）
- 貼上作答字串（`-` 對、`X` 錯）→ 產生結果與 Excel 下載

下一階段可加：精準切題（含 6(1)、三-2(2)）、全班 CSV 匯入、學習表現/難易度分析與報表。
"""
)

# ---- session state init ----
if "pdf_bytes" not in st.session_state:
    st.session_state.pdf_bytes = None

if "uploaded_sig" not in st.session_state:
    st.session_state.uploaded_sig = None

if "parsed_sig" not in st.session_state:
    st.session_state.parsed_sig = None

if "full_text" not in st.session_state:
    st.session_state.full_text = ""

if "items" not in st.session_state:
    st.session_state.items = []

if "last_result_df" not in st.session_state:
    st.session_state.last_result_df = None

if "last_message" not in st.session_state:
    st.session_state.last_message = ""


# -----------------------------
# Section 1: Upload & Parse
# -----------------------------
st.subheader("1) 上傳考卷 PDF")
pdf_file = st.file_uploader("請上傳可選字 PDF", type=["pdf"], key="pdf_uploader")

col1, col2, col3 = st.columns([1, 1, 1])

with col1:
    auto_parse = st.checkbox("上傳後自動解析（推薦）", value=True)

with col2:
    parse_btn = st.button("開始解析（抽取文字）", type="primary", disabled=(pdf_file is None and st.session_state.pdf_bytes is None))

with col3:
    reset_btn = st.button("清空本次資料（Reset）")

if reset_btn:
    st.session_state.pdf_bytes = None
    st.session_state.uploaded_sig = None
    st.session_state.parsed_sig = None
    st.session_state.full_text = ""
    st.session_state.items = []
    st.session_state.last_result_df = None
    st.session_state.last_message = ""
    st.success("已清空本次資料。請重新上傳 PDF。")

# Detect new upload and store bytes ONCE (only when file changes)
if pdf_file is not None:
    sig = file_signature(pdf_file)
    if st.session_state.uploaded_sig != sig:
        st.session_state.uploaded_sig = sig
        st.session_state.parsed_sig = None
        st.session_state.full_text = ""
        st.session_state.items = []
        st.session_state.last_result_df = None
        st.session_state.last_message = ""
        st.session_state.pdf_bytes = pdf_file.getvalue()
        st.success("已上傳新檔案。")

# Decide whether to parse
should_parse = False
if st.session_state.pdf_bytes is not None:
    # auto-parse only once per file
    if auto_parse and st.session_state.parsed_sig != st.session_state.uploaded_sig:
        should_parse = True
    # or user clicks parse
    if parse_btn:
        should_parse = True

if should_parse:
    with st.spinner("解析中..."):
        full_text, _ = extract_text_from_pdf(st.session_state.pdf_bytes)
        st.session_state.full_text = full_text
        st.session_state.items = guess_exam_items(full_text)
        st.session_state.parsed_sig = st.session_state.uploaded_sig
    st.success("解析完成！")

# Show parse preview (optional)
if st.session_state.full_text:
    st.divider()
    st.subheader("解析預覽")
    with st.expander("文字預覽（前 1200 字）", expanded=False):
        preview_text = st.session_state.full_text[:1200]
        st.text(preview_text + ("…" if len(st.session_state.full_text) > 1200 else ""))

    st.write(f"偵測到作答點數量（粗估）：**{len(st.session_state.items)}**")
    if st.session_state.items:
        df_items = pd.DataFrame([x.__dict__ for x in st.session_state.items])
        st.dataframe(df_items[["order_index", "label", "stem_preview"]], use_container_width=True, height=260)
else:
    st.info("尚未解析文字：請上傳 PDF，並等待自動解析或按「開始解析」。")

# -----------------------------
# Section 2: Answer analysis (ALWAYS visible)
# -----------------------------
st.divider()
st.subheader("2) 作答分析（輸入框永遠顯示）")

with st.form("analyze_form", clear_on_submit=False):
    ans_str = st.text_input("貼上作答字串（例：-------X-X-----XX-XXX--X）", value="")
    submitted = st.form_submit_button("分析作答")

if submitted:
    if not st.session_state.items:
        st.session_state.last_message = "❌ 尚未解析到作答點：請先上傳 PDF 並完成解析。"
        st.session_state.last_result_df = None
    else:
        correctness = parse_answer_string(ans_str)
        if len(correctness) == 0:
            st.session_state.last_message = "❌ 作答字串沒有讀到 '-' 或 'X'，請確認輸入格式（只接受 - 或 X）。"
            st.session_state.last_result_df = None
        else:
            msg = "✅ 已完成作答分析。"
            if len(correctness) != len(st.session_state.items):
                msg += f"（提醒：作答長度 {len(correctness)} ≠ 作答點 {len(st.session_state.items)}，目前只分析前 {min(len(correctness), len(st.session_state.items))} 題）"
            df = build_results_df(st.session_state.items, correctness)
            st.session_state.last_result_df = df
            st.session_state.last_message = msg

# Results area (always shows message if exists)
if st.session_state.last_message:
    if st.session_state.last_result_df is None:
        st.error(st.session_state.last_message)
    else:
        st.success(st.session_state.last_message)

if st.session_state.last_result_df is not None:
    df = st.session_state.last_result_df
    st.dataframe(df, use_container_width=True, height=360)

    wrong = df[df["is_correct"] == False]
    st.markdown(f"**錯題數：{len(wrong)}**")
    if len(wrong) > 0:
        st.write("錯題題號：", ", ".join(wrong["label"].astype(str).tolist()))

    xls = to_excel_bytes(df)
    st.download_button(
        label="下載 Excel 報表",
        data=xls,
        file_name="attempt_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
