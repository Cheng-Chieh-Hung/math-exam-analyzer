import io
import re
from dataclasses import dataclass
from typing import List, Tuple, Optional, Any

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
    stem: str  # 題幹全文


# -----------------------------
# Helpers (robust session handling)
# -----------------------------
def safe_items() -> List[ExamItem]:
    items = st.session_state.get("items", None)
    if not isinstance(items, list):
        st.session_state["items"] = []
        return []
    return items


def set_items(items: Any) -> None:
    st.session_state["items"] = items if isinstance(items, list) else []


def safe_q_df() -> pd.DataFrame:
    df = st.session_state.get("q_df", None)
    if isinstance(df, pd.DataFrame):
        return df
    return pd.DataFrame(columns=["order_index", "label", "section", "score", "difficulty", "learning_code", "note", "stem"])


def set_q_df(df: Any) -> None:
    st.session_state["q_df"] = df if isinstance(df, pd.DataFrame) else safe_q_df()


def file_signature(uploaded_file) -> str:
    return f"{uploaded_file.name}:{uploaded_file.size}"


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
    """
    只判斷兩種題號格式：
      - 1. 2. 3. ...
      - 1、2、3、...
    題號需在行首。
    """
    lines = [ln.rstrip() for ln in full_text.splitlines()]
    anchors = []
    pattern = re.compile(r"^(?P<q>\d{1,3})(?P<sep>[\.、])\s*(?P<rest>\S.*)$")

    for i, raw in enumerate(lines):
        ln = raw.strip()
        m = pattern.match(ln)
        if not m:
            continue
        anchors.append((i, m.group("q")))

    # 去連續重複
    filtered = []
    last_q = None
    for idx, q in anchors:
        if q != last_q:
            filtered.append((idx, q))
            last_q = q

    items: List[ExamItem] = []
    for k, (start_i, q) in enumerate(filtered):
        end_i = filtered[k + 1][0] if k + 1 < len(filtered) else len(lines)
        block_lines = [x.strip() for x in lines[start_i:end_i] if x.strip()]
        stem = "\n".join(block_lines)

        items.append(
            ExamItem(
                order_index=k + 1,
                label=q,
                section="未知",
                score=None,
                stem=stem,
            )
        )
    return items


def parse_answer_string(ans: str) -> List[bool]:
    cleaned = [c for c in (ans or "").strip() if c in ["-", "X", "x"]]
    return [c == "-" for c in cleaned]


def build_teacher_fill_df(items: List[ExamItem]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "order_index": it.order_index,
                "label": it.label,
                "section": it.section,         # 老師可改
                "score": it.score,             # 老師可填
                "difficulty": "",              # 老師可填：易/中/難
                "learning_code": "",           # 老師可填：代碼
                "note": "",                    # 老師可填：備註
                "stem": it.stem,               # 題幹全文
            }
            for it in items
        ]
    )


def build_results_df(q_df: pd.DataFrame, correctness: List[bool]) -> pd.DataFrame:
    dfq = q_df.copy()
    if "order_index" in dfq.columns:
        dfq = dfq.sort_values("order_index")
    dfq = dfq.reset_index(drop=True)

    n = min(len(dfq), len(correctness))
    rows = []
    for i in range(n):
        rows.append(
            {
                "order_index": int(dfq.loc[i, "order_index"]),
                "label": str(dfq.loc[i, "label"]),
                "difficulty": dfq.loc[i, "difficulty"],
                "learning_code": dfq.loc[i, "learning_code"],
                "is_correct": correctness[i],
                "result": "對" if correctness[i] else "錯",
                "stem": dfq.loc[i, "stem"],
            }
        )
    return pd.DataFrame(rows)


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "sheet1") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


# -----------------------------
# 回填 Excel 套用（核心）
# -----------------------------
REQUIRED_Q_COLS = ["order_index", "label", "stem"]  # 最少要有這些，才有辦法對齊


def read_teacher_excel(uploaded_xlsx, sheet_name: str = "questions") -> pd.DataFrame:
    """讀取老師回填的 Excel（questions sheet）。"""
    df = pd.read_excel(uploaded_xlsx, sheet_name=sheet_name, engine="openpyxl")
    # 標準化欄名（去空白）
    df.columns = [str(c).strip() for c in df.columns]
    return df


def normalize_q_df(df: pd.DataFrame) -> pd.DataFrame:
    """確保題目表包含所有欄位；缺的補空欄。"""
    target_cols = ["order_index", "label", "section", "score", "difficulty", "learning_code", "note", "stem"]
    out = df.copy()
    for c in target_cols:
        if c not in out.columns:
            out[c] = "" if c not in ["score"] else None
    out = out[target_cols]

    # order_index 轉 numeric（容錯）
    out["order_index"] = pd.to_numeric(out["order_index"], errors="coerce")
    return out


def apply_teacher_fill(current_q: pd.DataFrame, filled_q: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    """
    把老師回填資料套回 current_q。
    對齊策略：
    1) 優先用 order_index（最穩）
    2) 若 order_index 缺失，退回用 label
    """
    cur = normalize_q_df(current_q)
    fil = normalize_q_df(filled_q)

    # 檢查必要欄位
    missing = [c for c in REQUIRED_Q_COLS if c not in fil.columns]
    if missing:
        return cur, f"❌ 回填檔缺少必要欄位：{', '.join(missing)}"

    # 決定 join key
    use_order = fil["order_index"].notna().all() and cur["order_index"].notna().all()
    if use_order:
        key = "order_index"
    else:
        key = "label"

    # 我們只想「套用老師可回填欄位」，不要覆蓋題幹/題號（避免錯檔蓋掉）
    fill_cols = ["section", "score", "difficulty", "learning_code", "note"]
    fil_small = fil[[key] + fill_cols].copy()

    merged = cur.merge(fil_small, on=key, how="left", suffixes=("", "_new"))

    # 用 new 覆蓋舊（new 有值才蓋）
    for c in fill_cols:
        newc = f"{c}_new"
        if newc not in merged.columns:
            continue
        merged[c] = merged[newc].combine_first(merged[c])
        merged.drop(columns=[newc], inplace=True)

    merged = merged.sort_values("order_index").reset_index(drop=True)

    applied_count = int(merged[fill_cols].notna().any(axis=1).sum())
    return merged, f"✅ 已套用回填資料（對齊欄位：{key}）。目前題目數 {len(merged)}，至少有回填資料的題數約 {applied_count}。"


# -----------------------------
# Session init + actions
# -----------------------------
def init_state():
    defaults = {
        "pdf_bytes": None,
        "uploaded_sig": None,
        "parsed_sig": None,
        "full_text": "",
        "items": [],
        "q_df": safe_q_df(),
        "ans_str": "",
        "last_message": "",
        "last_result_df": None,
        "apply_msg": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    _ = safe_items()
    if not isinstance(st.session_state.get("q_df", None), pd.DataFrame):
        st.session_state["q_df"] = safe_q_df()


def reset_all():
    st.session_state["pdf_bytes"] = None
    st.session_state["uploaded_sig"] = None
    st.session_state["parsed_sig"] = None
    st.session_state["full_text"] = ""
    st.session_state["items"] = []
    st.session_state["q_df"] = safe_q_df()
    st.session_state["ans_str"] = ""
    st.session_state["last_message"] = ""
    st.session_state["last_result_df"] = None
    st.session_state["apply_msg"] = ""


def run_analysis():
    q_df = safe_q_df()
    if q_df.empty:
        st.session_state["last_message"] = "❌ 尚未建立題目表：請先上傳 PDF 並完成解析。"
        st.session_state["last_result_df"] = None
        return

    correctness = parse_answer_string(st.session_state.get("ans_str", ""))
    if len(correctness) == 0:
        st.session_state["last_message"] = "❌ 作答字串沒有讀到 '-' 或 'X'，請確認輸入格式（只接受 - 或 X）。"
        st.session_state["last_result_df"] = None
        return

    msg = "✅ 已完成作答分析。"
    if len(correctness) != len(q_df):
        msg += f"（提醒：作答長度 {len(correctness)} ≠ 題目數 {len(q_df)}，目前只分析前 {min(len(correctness), len(q_df))} 題）"

    df = build_results_df(q_df, correctness)
    st.session_state["last_result_df"] = df
    st.session_state["last_message"] = msg


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="考卷分析 MVP", layout="wide")
init_state()

st.title("考卷分析 MVP（可選字 PDF）")
st.caption("上傳 PDF → 拆題 → 下載題目 Excel（老師可回填）→ 回填後上傳套用 → 作答分析用最新回填資料")

# Sidebar
st.sidebar.header("作答情形輸入")
st.session_state["ans_str"] = st.sidebar.text_input(
    "作答字串（- 對 / X 錯）",
    value=st.session_state.get("ans_str", ""),
    placeholder="例：-------X-X-----XX-XXX--X",
)
st.sidebar.button("分析作答", type="primary", on_click=run_analysis)
st.sidebar.divider()
st.sidebar.button("Reset（清空）", on_click=reset_all)

# Upload & Parse
st.subheader("1) 上傳考卷 PDF")
pdf_file = st.file_uploader("請上傳可選字 PDF", type=["pdf"], key="pdf_uploader")

col1, col2 = st.columns([1, 1])
with col1:
    auto_parse = st.checkbox("上傳後自動解析一次", value=True)
with col2:
    parse_btn = st.button(
        "手動解析",
        type="primary",
        disabled=(pdf_file is None and st.session_state.get("pdf_bytes") is None),
    )

# Store new file bytes once
if pdf_file is not None:
    sig = file_signature(pdf_file)
    if st.session_state.get("uploaded_sig") != sig:
        st.session_state["uploaded_sig"] = sig
        st.session_state["parsed_sig"] = None
        st.session_state["full_text"] = ""
        st.session_state["last_message"] = ""
        st.session_state["last_result_df"] = None
        st.session_state["apply_msg"] = ""

        st.session_state["pdf_bytes"] = pdf_file.getvalue()
        set_items([])
        set_q_df(safe_q_df())
        st.success("已上傳新檔案。")

# Parse only once per file (or manual)
should_parse = False
if st.session_state.get("pdf_bytes") is not None:
    if auto_parse and st.session_state.get("parsed_sig") != st.session_state.get("uploaded_sig"):
        should_parse = True
    if parse_btn:
        should_parse = True

if should_parse:
    with st.spinner("解析中..."):
        full_text, _ = extract_text_from_pdf(st.session_state["pdf_bytes"])
        st.session_state["full_text"] = full_text

        items = guess_exam_items(full_text)
        set_items(items)

        q_df = build_teacher_fill_df(items)
        set_q_df(q_df)

        st.session_state["parsed_sig"] = st.session_state.get("uploaded_sig")
    st.success("解析完成！已建立『題目表（老師可回填）』。")

# 題目表（老師可回填） + 上傳回填套用
st.subheader("2) 題目表（老師可回填）與回填套用")

q_df = safe_q_df()

if st.session_state.get("full_text") and not q_df.empty:
    st.write(f"題目數：**{len(q_df)}**（題號格式僅支援行首 `1.` 或 `1、`）")

    # (A) 下載原始題目表
    st.download_button(
        label="下載題目 Excel（老師可回填）",
        data=to_excel_bytes(q_df, sheet_name="questions"),
        file_name="questions_teacher_fill.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # (B) 上傳老師回填 Excel 並套用
    st.markdown("### 上傳老師回填 Excel → 套用到系統")
    colA, colB = st.columns([2, 1])
    with colA:
        filled_file = st.file_uploader(
            "上傳老師回填後的 Excel（建議使用剛下載的 questions_teacher_fill.xlsx 回填）",
            type=["xlsx"],
            key="filled_uploader",
        )
    with colB:
        sheet_name = st.text_input("Sheet 名稱", value="questions")

    apply_btn = st.button("套用回填資料", type="primary", disabled=(filled_file is None))

    if apply_btn and filled_file is not None:
        try:
            filled_df = read_teacher_excel(filled_file, sheet_name=sheet_name)
            merged, msg = apply_teacher_fill(q_df, filled_df)
            set_q_df(merged)
            st.session_state["apply_msg"] = msg
        except Exception as e:
            st.session_state["apply_msg"] = f"❌ 套用失敗：請確認檔案格式與 sheet 名稱。({type(e).__name__})"

    if st.session_state.get("apply_msg"):
        if st.session_state["apply_msg"].startswith("✅"):
            st.success(st.session_state["apply_msg"])
        else:
            st.error(st.session_state["apply_msg"])

    # (C) 網頁上直接回填（可選）
    st.markdown("### 直接在網頁回填（可選）")
    edited_df = st.data_editor(
        safe_q_df(),
        use_container_width=True,
        height=420,
        num_rows="fixed",
        column_config={
            "order_index": st.column_config.NumberColumn("order_index", disabled=True),
            "label": st.column_config.TextColumn("label", disabled=True),
            "section": st.column_config.TextColumn("section"),
            "score": st.column_config.NumberColumn("score", min_value=0, step=1),
            "difficulty": st.column_config.SelectboxColumn("difficulty", options=["", "易", "中", "難"]),
            "learning_code": st.column_config.TextColumn("learning_code"),
            "note": st.column_config.TextColumn("note"),
            "stem": st.column_config.TextColumn("stem", disabled=True),
        },
        key="q_df_editor",
    )
    set_q_df(edited_df)

    # (D) 下載最新版（含回填）
    st.download_button(
        label="下載（含目前回填內容）Excel",
        data=to_excel_bytes(safe_q_df(), sheet_name="questions"),
        file_name="questions_teacher_fill_UPDATED.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("尚未解析：請先上傳 PDF 並完成解析，才會產生題目表。")

# 作答分析結果
st.divider()
st.subheader("3) 作答分析結果（使用最新版題目表）")

msg = st.session_state.get("last_message", "")
df = st.session_state.get("last_result_df", None)

if msg:
    if df is None:
        st.error(msg)
    else:
        st.success(msg)

if df is not None:
    st.dataframe(df, use_container_width=True, height=360)

    wrong = df[df["is_correct"] == False]
    st.markdown(f"**錯題數：{len(wrong)}**")
    if len(wrong) > 0:
        st.write("錯題題號：", ", ".join(wrong["label"].astype(str).tolist()))

    st.download_button(
        label="下載作答結果 Excel",
        data=to_excel_bytes(df, sheet_name="attempt"),
        file_name="attempt_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
