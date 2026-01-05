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
    section: str
    score: Optional[float]
    stem_preview: str


# -----------------------------
# Core functions
# -----------------------------
def extract_text_from_pdf(pdf_bytes: bytes) -> Tuple[str, List[str]]:
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
    lines = [ln.strip() for ln in full_text.splitlines()]
    anchors = []
    for i, ln in enumerate(lines):
        # 例： "1 " / "1." / "1、"
        m = re.match(r"^(\d{1,3})\s*[\.、]?\s+", ln)
        if m:
            anchors.append((i, m.group(1)))

    # 去掉連續重複題號
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
    cleaned = [c for c in (ans or "").strip() if c in ["-", "X", "x"]]
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
    return f"{uploaded_file.name}:{uploaded_file.size}"


# -----------------------------
# Session init
# -----------------------------
def init_state():
    defaults = {
        "pdf_bytes": None,
        "uploaded_sig": None,
        "parsed_sig": None,
        "full_text": "",
        "items": [],
        "ans_str": "",
        "last_message": "",
        "last_result_df": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def reset_all():
    st.session_state.pdf_bytes = None
    st.session_state.uploaded_sig = None
    st.session_state.parsed_sig = None
    st.session_state.full_text = ""
    st.session_state.items = []
    st.session_state.ans_str = ""
    st.session_state.last_message = ""
    st.session_state.last_result_df = None


def run_analysis():
    """按下『分析作答』時執行：一定會寫入 last_message / last_result_df。"""
    if not st.session_state.items:
        st.session_state.last_message = "❌ 尚未解析到作答點：請先上傳 PDF 並完成解析。"
        st.session_state.last_result_df = None
        return

    correctness = parse_answer_string(st.session_state.ans_str)
    if len(correctness) == 0:
        st.session_state.last_message = "❌ 作答字串沒有讀到 '-' 或 'X'，請確認輸入格式（只接受 - 或 X）。"
        st.session_state.last_result_df = None
        return

    msg = "✅ 已完成作答分析。"
    if len(correctness) != len(st.session_state.items):
        msg += f"（提醒：作答長度 {len(correctness)} ≠ 作答點 {len(st.session_state.items)}，目前只分析前 {min(len(correctness), len(st.session_state.items))} 題）"

    df = build_results_df(st.session_state.items, correctness)
    st.session_state.last_result_df = df
    st.session_state.last_message = msg


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="數學考卷分析 MVP", layout="wide")
init_state()

st.title("數學考卷分析 MVP（可選字 PDF）")
st.caption("上傳 PDF → 解析一次 → 左側輸入作答字串按『分析作答』→ 主畫面顯示結果（保證不會消失）")

# Sidebar: ALWAYS visible controls
st.sidebar.header("作答分析（永遠顯示）")
st.session_state.ans_str = st.sidebar.text_input(
    "作答字串（- 對 / X 錯）",
    value=st.session_state.ans_str,
    placeholder="例：-------X-X-----XX-XXX--X",
)
st.sidebar.button("分析作答", type="primary", on_click=run_analysis)
st.sidebar.divider()
st.sidebar.button("Reset（清空）", on_click=reset_all)

# Main: upload + parse
st.subheader("1) 上傳考卷 PDF")
pdf_file = st.file_uploader("請上傳可選字 PDF", type=["pdf"], key="pdf_uploader")

col1, col2 = st.columns([1, 1])
with col1:
    auto_parse = st.checkbox("上傳後自動解析一次", value=True)
with col2:
    parse_btn = st.button("手動解析", type="primary", disabled=(pdf_file is None and st.session_state.pdf_bytes is None))

# Detect new file & store bytes once
if pdf_file is not None:
    sig = file_signature(pdf_file)
    if st.session_state.uploaded_sig != sig:
        st.session_state.uploaded_sig = sig
        st.session_state.parsed_sig = None
        st.session_state.full_text = ""
        st.session_state.items = []
        st.session_state.last_message = ""
        st.session_state.last_result_df = None
        st.session_state.pdf_bytes = pdf_file.getvalue()
        st.success("已上傳新檔案。")

# Parse logic (only once per file or manual)
should_parse = False
if st.session_state.pdf_bytes is not None:
    if auto_parse and st.session_state.parsed_sig != st.session_state.uploaded_sig:
        should_parse = True
    if parse_btn:
        should_parse = True

if should_parse:
    with st.spinner("解析中..."):
        full_text, _ = extract_text_from_pdf(st.session_state.pdf_bytes)
        st.session_state.full_text = full_text
        st.session_state.items = guess_exam_items(full_text)
        st.session_state.parsed_sig = st.session_state.uploaded_sig
    st.success("解析完成！請到左側輸入作答字串並按『分析作答』。")

# Preview
st.subheader("2) 解析預覽")
if st.session_state.full_text:
    st.write(f"偵測到作答點數量（粗估）：**{len(st.session_state.items)}**")
    with st.expander("文字預覽（前 1200 字）", expanded=False):
        preview_text = st.session_state.full_text[:1200]
        st.text(preview_text + ("…" if len(st.session_state.full_text) > 1200 else ""))

    if st.session_state.items:
        df_items = pd.DataFrame([x.__dict__ for x in st.session_state.items])
        st.dataframe(df_items[["order_index", "label", "stem_preview"]], use_container_width=True, height=260)
else:
    st.info("尚未解析：請先上傳 PDF 並等待自動解析或按『手動解析』。")

# Result area (ALWAYS visible)
st.divider()
st.subheader("3) 作答分析結果（按『分析作答』後會顯示在這裡）")

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
