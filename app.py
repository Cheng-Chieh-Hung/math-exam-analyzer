import io
import re
from dataclasses import dataclass
from typing import List, Tuple, Optional

import pandas as pd
import pdfplumber
import streamlit as st


# -----------------------------
# Utilities
# -----------------------------
@dataclass
class ExamItem:
    order_index: int
    label: str
    section: str  # 單選 / 填充 / 非選 / 未知
    score: Optional[float]
    stem_preview: str


def extract_text_from_pdf(pdf_bytes: bytes) -> Tuple[str, List[str]]:
    """Return (full_text, per_page_texts). For selectable-text PDFs."""
    per_page = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            per_page.append(txt)
    full_text = "\n\n".join(per_page)
    return full_text, per_page


def guess_exam_items(full_text: str) -> List[ExamItem]:
    """
    MVP heuristic: detect question numbers by lines starting with:
      - '1 ' or '1.' etc (single choice)
    and also detect fill-in numbers like '1.' under filling section.
    For now: we only estimate "answer points count" by finding leading integers.
    """
    lines = [ln.strip() for ln in full_text.splitlines()]
    anchors = []
    for i, ln in enumerate(lines):
        # Match: "1 " or "1." or "1、"
        m = re.match(r"^(\d{1,3})\s*[\.、]?\s+", ln)
        if m:
            qn = m.group(1)
            anchors.append((i, qn, ln))

    items: List[ExamItem] = []
    # De-duplicate consecutive same q number (sometimes wrapped lines)
    filtered = []
    last_qn = None
    for (idx, qn, ln) in anchors:
        if qn != last_qn:
            filtered.append((idx, qn))
            last_qn = qn

    # Build items using anchor to next anchor
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
    """'-' => correct(True), 'X'/'x' => wrong(False). Ignore other chars/spaces."""
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
    # If lengths mismatch, show remainder as warnings in UI
    return pd.DataFrame(rows)


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "results") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="數學考卷分析 MVP", layout="wide")
st.title("數學考卷分析 MVP（可選字 PDF）")

st.markdown(
    """
此 MVP 先做到：
- 上傳 **可選字 PDF** → 抽文字預覽  
- 粗略偵測「作答點（題號）」數量  
- 貼上作答字串（`-`對、`X`錯）→ 生成錯題清單與 Excel 下載  

下一階段再加：更精準切題（含 6(1)、三-2(2)）、學習表現對應、難易度、班級/個人報表。
"""
)

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1) 上傳考卷 PDF")
    pdf_file = st.file_uploader("請上傳可選字 PDF", type=["pdf"])
    extract_btn = st.button("開始解析（抽取文字）", type="primary")

with col2:
    st.subheader("2) 作答字串")
    ans_str = st.text_input('貼上作答字串（例：-------X-X-----XX-XXX--X）', value="")
    analyze_btn = st.button("分析作答", disabled=(ans_str.strip() == ""))

# Session state
if "full_text" not in st.session_state:
    st.session_state.full_text = ""
if "items" not in st.session_state:
    st.session_state.items = []

if extract_btn:
    if not pdf_file:
        st.error("請先上傳 PDF。")
    else:
        pdf_bytes = pdf_file.read()
        with st.spinner("解析中..."):
            full_text, per_page = extract_text_from_pdf(pdf_bytes)
        st.session_state.full_text = full_text
        st.session_state.items = guess_exam_items(full_text)
        st.success("解析完成！")

if st.session_state.full_text:
    st.divider()
    st.subheader("解析預覽")
    with st.expander("文字預覽（前 1500 字）", expanded=True):
        st.text(st.session_state.full_text[:1500] + ("…" if len(st.session_state.full_text) > 1500 else ""))

    st.subheader("作答點偵測（MVP 粗估）")
    items = st.session_state.items
    st.write(f"偵測到作答點數量：**{len(items)}**（此為粗估，之後可做更精準切題）")

    if items:
        df_items = pd.DataFrame([item.__dict__ for item in items])
        st.dataframe(df_items[["order_index", "label", "stem_preview"]], use_container_width=True, height=280)

if analyze_btn:
    if not st.session_state.items:
        st.error("請先完成 PDF 解析與作答點偵測。")
    else:
        correctness = parse_answer_string(ans_str)
        items = st.session_state.items

        if len(correctness) != len(items):
            st.warning(
                f"作答字串長度（{len(correctness)}）與偵測作答點數量（{len(items)}）不一致。"
                f" 目前只分析前 {min(len(correctness), len(items))} 題。"
            )

        df = build_results_df(items, correctness)
        st.divider()
        st.subheader("作答分析結果")
        st.dataframe(df, use_container_width=True, height=360)

        wrong = df[df["is_correct"] == False]
        st.markdown(f"**錯題數：{len(wrong)}**")
        if len(wrong) > 0:
            st.write("錯題題號：", ", ".join(wrong["label"].astype(str).tolist()))

        xls_bytes = to_excel_bytes(df, sheet_name="attempt")
        st.download_button(
            label="下載 Excel 報表",
            data=xls_bytes,
            file_name="attempt_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        