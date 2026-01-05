import io
import re
from dataclasses import dataclass
from typing import List, Tuple, Optional, Any

import pandas as pd
import pdfplumber
import streamlit as st


# =========================================================
# 中文欄名對照（下載/上傳都用這套）
# =========================================================
COL_MAP_EN2ZH = {
    "order_index": "題序",
    "label": "題號",
    "section": "題型",
    "score": "配分",
    "difficulty": "難易度",
    "learning_code": "學習表現代碼",
    "note": "備註",
    "stem": "題幹",
}
COL_MAP_ZH2EN = {v: k for k, v in COL_MAP_EN2ZH.items()}

TEACHER_COLS_EN = list(COL_MAP_EN2ZH.keys())
TEACHER_COLS_ZH = [COL_MAP_EN2ZH[c] for c in TEACHER_COLS_EN]

SHEET_NAME_ZH = "題目表"       # 下載模板的 sheet
SHEET_NAME_EN = "questions"    # 若有人用英文 sheet，我們也容錯


# -----------------------------
# Data structures
# -----------------------------
@dataclass
class ExamItem:
    order_index: int
    label: str
    section: str
    score: Optional[float]
    stem: str  # 題幹全文（不截斷）


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
    return pd.DataFrame(columns=TEACHER_COLS_EN)


def set_q_df(df: Any) -> None:
    st.session_state["q_df"] = df if isinstance(df, pd.DataFrame) else safe_q_df()


def file_signature(uploaded_file) -> str:
    return f"{uploaded_file.name}:{uploaded_file.size}"


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "sheet1") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


def q_df_to_teacher_df_zh(q_df_en: pd.DataFrame) -> pd.DataFrame:
    """內部英文欄位 → 老師版中文欄位（固定欄位順序）。"""
    df = q_df_en.copy()
    for c in TEACHER_COLS_EN:
        if c not in df.columns:
            df[c] = "" if c != "score" else None
    df = df[TEACHER_COLS_EN]
    return df.rename(columns=COL_MAP_EN2ZH)


def teacher_df_to_q_df_en(df_any: pd.DataFrame) -> pd.DataFrame:
    """
    老師上傳（可能中文欄名/英文欄名）→ 轉成內部英文欄名 + 補齊欄位。
    """
    df = df_any.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # 若含中文欄名，轉英文
    rename_map = {}
    for c in df.columns:
        if c in COL_MAP_ZH2EN:
            rename_map[c] = COL_MAP_ZH2EN[c]
    if rename_map:
        df = df.rename(columns=rename_map)

    # 補齊欄位
    for c in TEACHER_COLS_EN:
        if c not in df.columns:
            df[c] = "" if c != "score" else None

    df = df[TEACHER_COLS_EN].copy()

    # 型別容錯
    df["order_index"] = pd.to_numeric(df["order_index"], errors="coerce")
    # label 保持字串（避免 01 變 1）
    df["label"] = df["label"].astype(str).str.strip()

    return df


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
    只判斷兩種題號格式作為作答點：
      - 1. 2. 3. ...
      - 1、2、3、...
    題號必須出現在「行首」。
    題幹：從 anchor 到下一題 anchor 的全文（不截斷）。
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

    # 去掉連續重複題號
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
    """由拆題 items 建立內部英文題目表（老師可回填欄位預設空）。"""
    return pd.DataFrame(
        [
            {
                "order_index": it.order_index,
                "label": it.label,
                "section": it.section,
                "score": it.score,
                "difficulty": "",
                "learning_code": "",
                "note": "",
                "stem": it.stem,
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
                "題序": int(dfq.loc[i, "order_index"]) if pd.notna(dfq.loc[i, "order_index"]) else i + 1,
                "題號": str(dfq.loc[i, "label"]),
                "難易度": dfq.loc[i, "difficulty"],
                "學習表現代碼": dfq.loc[i, "learning_code"],
                "結果": "對" if correctness[i] else "錯",
                "題幹": dfq.loc[i, "stem"],
            }
        )
    return pd.DataFrame(rows)


# -----------------------------
# 回填 Excel 套用（核心）
# -----------------------------
FILL_COLS_EN = ["section", "score", "difficulty", "learning_code", "note"]


def read_teacher_excel(uploaded_xlsx, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(uploaded_xlsx, sheet_name=sheet_name, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def apply_teacher_fill(current_q_en: pd.DataFrame, uploaded_df_any: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    """
    把老師回填套回 current_q（內部英文欄位）。
    對齊策略：
    1) 優先用 order_index（題序）
    2) 否則用 label（題號）
    """
    cur = teacher_df_to_q_df_en(current_q_en)
    fil = teacher_df_to_q_df_en(uploaded_df_any)

    # 決定 join key
    use_order = cur["order_index"].notna().all() and fil["order_index"].notna().all()
    key = "order_index" if use_order else "label"

    fil_small = fil[[key] + FILL_COLS_EN].copy()

    merged = cur.merge(fil_small, on=key, how="left", suffixes=("", "_new"))

    for c in FILL_COLS_EN:
        newc = f"{c}_new"
        if newc not in merged.columns:
            continue
        merged[c] = merged[newc].combine_first(merged[c])
        merged.drop(columns=[newc], inplace=True)

    merged = merged.sort_values("order_index").reset_index(drop=True)

    applied_rows = int(merged[FILL_COLS_EN].notna().any(axis=1).sum())
    return merged, f"✅ 已套用回填資料（對齊欄位：{COL_MAP_EN2ZH.get(key, key)}）。題目數 {len(merged)}，至少有回填資料的題數約 {applied_rows}。"


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
st.caption("上傳 PDF → 拆題 → 下載中文題目 Excel（老師可回填）→ 回填後上傳套用 → 作答分析用最新回填資料")

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

# 題目表 + 回填套用
st.subheader("2) 題目表（老師可回填）與回填套用")

q_df_en = safe_q_df()

if st.session_state.get("full_text") and not q_df_en.empty:
    st.write(f"題目數：**{len(q_df_en)}**（題號格式僅支援行首 `1.` 或 `1、`）")

    # (A) 下載中文題目表
    q_df_zh = q_df_to_teacher_df_zh(q_df_en)
    st.download_button(
        label="下載題目 Excel（中文欄名，可回填）",
        data=to_excel_bytes(q_df_zh, sheet_name=SHEET_NAME_ZH),
        file_name="題目表_老師回填.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # (B) 上傳回填 Excel 套用
    st.markdown("### 上傳老師回填 Excel → 套用到系統")
    colA, colB, colC = st.columns([2, 1, 1])
    with colA:
        filled_file = st.file_uploader(
            "上傳老師回填後的 Excel（可用本系統下載的模板回填）",
            type=["xlsx"],
            key="filled_uploader",
        )
    with colB:
        # 預設中文 sheet，也容錯英文
        sheet_name = st.text_input("Sheet 名稱", value=SHEET_NAME_ZH)
    with colC:
        apply_btn = st.button("套用回填", type="primary", disabled=(filled_file is None))

    if apply_btn and filled_file is not None:
        try:
            # 先嘗試使用使用者輸入的 sheet；失敗則容錯嘗試另一個
            try:
                uploaded_df = read_teacher_excel(filled_file, sheet_name=sheet_name)
            except Exception:
                fallback = SHEET_NAME_EN if sheet_name == SHEET_NAME_ZH else SHEET_NAME_ZH
                uploaded_df = read_teacher_excel(filled_file, sheet_name=fallback)

            merged, msg = apply_teacher_fill(q_df_en, uploaded_df)
            set_q_df(merged)
            st.session_state["apply_msg"] = msg
        except Exception as e:
            st.session_state["apply_msg"] = f"❌ 套用失敗：請確認檔案格式與 sheet 名稱。({type(e).__name__})"

    if st.session_state.get("apply_msg"):
        if st.session_state["apply_msg"].startswith("✅"):
            st.success(st.session_state["apply_msg"])
        else:
            st.error(st.session_state["apply_msg"])

    # (C) Web 端回填（顯示中文標題，但內部欄位仍英文）
    st.markdown("### 直接在網頁回填（可選）")
    edited_df = st.data_editor(
        safe_q_df(),
        use_container_width=True,
        height=420,
        num_rows="fixed",
        column_config={
            "order_index": st.column_config.NumberColumn("題序", disabled=True),
            "label": st.column_config.TextColumn("題號", disabled=True),
            "section": st.column_config.TextColumn("題型"),
            "score": st.column_config.NumberColumn("配分", min_value=0, step=1),
            "difficulty": st.column_config.SelectboxColumn("難易度", options=["", "易", "中", "難"]),
            "learning_code": st.column_config.TextColumn("學習表現代碼"),
            "note": st.column_config.TextColumn("備註"),
            "stem": st.column_config.TextColumn("題幹", disabled=True),
        },
        key="q_df_editor",
    )
    set_q_df(edited_df)

    # (D) 下載最新版（中文欄名）
    latest_zh = q_df_to_teacher_df_zh(safe_q_df())
    st.download_button(
        label="下載（含目前回填內容）Excel（中文欄名）",
        data=to_excel_bytes(latest_zh, sheet_name=SHEET_NAME_ZH),
        file_name="題目表_老師回填_最新版.xlsx",
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

    st.download_button(
        label="下載作答結果 Excel（中文欄名）",
        data=to_excel_bytes(df, sheet_name="作答結果"),
        file_name="作答分析結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
