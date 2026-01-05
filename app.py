import io
import re
from dataclasses import dataclass
from typing import List, Tuple, Optional

import pandas as pd
import pdfplumber
import streamlit as st


@dataclass
class ExamItem:
    order_index: int
    label: str
    section: str
    score: Optional[float]
    stem_preview: str


def extract_text_from_pdf(pdf_bytes: bytes) -> Tuple[str, List[str]]:
    per_page = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            per_page.append(txt)
    full_text = "\n\n".join(per_page)
    return full_text, per_page


def guess_exam_items(full_text: str) -> List[ExamItem]:
    lines = [ln.strip() for ln in full_text.splitlines()]
    anchors = []
    for i, ln in enumerate(lines):
        m = re.match(r"^(\d{1,3})\s*[\.、]?\s+", ln)
        if m:
            anchors.append((i, m.group(1)))

    # 去除連續重複題號
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


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="數學考卷分析 MVP", layout="wide")
st.title("數學考卷分析 MVP（可選字 PDF）【測試重部署】")

# 初始化 session_state
if "pdf_bytes" not in st.session_state:
    st.session_state.pdf_bytes = None
if "full_text" not in st.session_state:
    st.session_state.full_text = ""
if "items" not in st.session_state:
    st.session_state.items = []
if "last_result_df" not in st.session_state:
    st.session_state.last_result_df = None
if "last_message" not in st.session_state:
    st.session_state.last_message = ""

# 上傳區
st.subheader("1) 上傳考卷 PDF")
pdf_file = st.file_uploader("請上傳可選字 PDF", type=["pdf"])

colA, colB = st.columns([1, 1])

with colA:
    auto_parse = st.checkbox("上傳後自動解析（推薦）", value=True)
with colB:
    parse_btn = st.button("開始解析（抽取文字）", type="primary")

# 儲存 pdf bytes（避免 rerun 後消失）
if pdf_file is not None:
    st.session_state.pdf_bytes = pdf_file.getvalue()

# 自動解析或手動解析
should_parse = False
if st.session_state.pdf_bytes and auto_parse:
    # 如果之前沒有解析過（或換檔）
    # 用 bytes 長度做簡單判斷；更嚴謹可做 hash
    if not st.session_state.full_text:
        should_parse = True
if parse_btn and st.session_state.pdf_bytes:
    should_parse = True

if should_parse:
    with st.spinner("解析中..."):
        full_text, _ = extract_text_from_pdf(st.session_state.pdf_bytes)
        st.session_state.full_text = full_text
        st.session_state.items = guess_exam_items(full_text)
    st.success("解析完成！")

# 顯示解析預覽
if st.session_state.full_text:
    st.divider()
    st.subheader("解析預覽")
    with st.expander("文字預覽（前 1200 字）", expanded=False):
        st.text(st.session_state.full_text[:1200] + ("…" if len(st.session_state.full_text) > 1200 else ""))

    st.write(f"偵測到作答點數量：**{len(st.session_state.items)}**")
    if st.session_state.items:
        df_items = pd.DataFrame([x.__dict__ for x in st.session_state.items])
        st.dataframe(df_items[["order_index", "label", "stem_preview"]], use_container_width=True, height=260)

# 作答分析區（用 form，避免按鈕被 rerun 吃掉）
st.divider()
st.subheader("2) 作答分析")
with st.form("analyze_form"):
    ans_str = st.text_input("貼上作答字串（例：-------X-X-----XX-XXX--X）", value="")
    submitted = st.form_submit_button("分析作答")

if submitted:
    # 一定會顯示訊息，避免「按了沒反應」
    if not st.session_state.items:
        st.session_state.last_message = "❌ 尚未解析到作答點：請先上傳 PDF 並完成解析。"
        st.session_state.last_result_df = None
    else:
        correctness = parse_answer_string(ans_str)
        items = st.session_state.items

        if len(correctness) == 0:
            st.session_state.last_message = "❌ 作答字串沒有讀到 '-' 或 'X'，請確認輸入格式。"
            st.session_state.last_result_df = None
        else:
            msg = "✅ 已完成分析。"
            if len(correctness) != len(items):
                msg += f"（提醒：作答長度 {len(correctness)} ≠ 作答點 {len(items)}，目前只分析前 {min(len(correctness), len(items))} 題）"
            df = build_results_df(items, correctness)
            st.session_state.last_result_df = df
            st.session_state.last_message = msg

# 結果區（永遠顯示）
if st.session_state.last_message:
    st.info(st.session_state.last_message)

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
