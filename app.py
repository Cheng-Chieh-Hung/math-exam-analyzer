import io
import re
from dataclasses import dataclass
from typing import List, Tuple, Optional, Any, Dict

import pandas as pd
import pdfplumber
import streamlit as st


# =========================================================
# 中文欄名對照（題目表：下載/上傳）
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
SHEET_NAME_ZH = "題目表"
SHEET_NAME_EN = "questions"


# =========================================================
# 全班作答匯入（中文欄名）
# =========================================================
ATT_COL_SEAT = "座號"
ATT_COL_NAME = "姓名"
ATT_COL_ANS = "作答字串"
ATT_SHEET_DEFAULT = "作答匯入"


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
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()


def to_excel_bytes_multi(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=str(name)[:31])
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
    """老師上傳（可能中文/英文欄名）→ 內部英文欄名 + 補齊欄位。"""
    df = df_any.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # 中文 → 英文
    rename_map = {c: COL_MAP_ZH2EN[c] for c in df.columns if c in COL_MAP_ZH2EN}
    if rename_map:
        df = df.rename(columns=rename_map)

    # 補欄位
    for c in TEACHER_COLS_EN:
        if c not in df.columns:
            df[c] = "" if c != "score" else None

    df = df[TEACHER_COLS_EN].copy()
    df["order_index"] = pd.to_numeric(df["order_index"], errors="coerce")
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


# =============================
# 單一學生（sidebar）作答：仍用原本 -/X 模式（不變）
# =============================
def parse_answer_string_sidebar(ans: str) -> List[bool]:
    cleaned = [c for c in (ans or "").strip() if c in ["-", "X", "x"]]
    return [c == "-" for c in cleaned]


def build_teacher_fill_df(items: List[ExamItem]) -> pd.DataFrame:
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


def build_results_df_single(q_df: pd.DataFrame, correctness: List[bool]) -> pd.DataFrame:
    dfq = q_df.copy().sort_values("order_index").reset_index(drop=True)
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


# =============================
# 全班匯入作答（你指定的判定規則）
# - 只有 '-' 算對
# - 其餘（任何符號）算錯
# - 強制一致：字串長度必須 = 題目數
# =============================
def normalize_class_answer_str(ans: str) -> str:
    """
    全班匯入用：移除空白字元，只保留其餘字元原樣。
    之後規則：'-' => 對，其他 => 錯。
    """
    if ans is None:
        return ""
    s = str(ans)
    # 移除空白（空格、tab、換行）
    s = re.sub(r"\s+", "", s)
    return s


def read_class_answers_excel(uploaded_xlsx, sheet_name: str = ATT_SHEET_DEFAULT) -> pd.DataFrame:
    df = pd.read_excel(uploaded_xlsx, sheet_name=sheet_name, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    # 必填欄位檢查（中文）
    missing = [c for c in [ATT_COL_SEAT, ATT_COL_ANS] if c not in df.columns]
    if missing:
        raise ValueError(f"缺少必要欄位：{', '.join(missing)}（請使用範本欄名）")

    if ATT_COL_NAME not in df.columns:
        df[ATT_COL_NAME] = ""

    out = df[[ATT_COL_SEAT, ATT_COL_NAME, ATT_COL_ANS]].copy()
    out[ATT_COL_SEAT] = out[ATT_COL_SEAT].astype(str).str.strip()
    out[ATT_COL_NAME] = out[ATT_COL_NAME].astype(str).replace({"nan": ""}).str.strip()
    out[ATT_COL_ANS] = out[ATT_COL_ANS].apply(normalize_class_answer_str)
    return out


def validate_class_answers_length(class_df: pd.DataFrame, n_questions: int) -> Tuple[bool, pd.DataFrame]:
    """
    強制一致：
    - 每個學生作答字串長度必須等於題目數 n_questions
    - 回傳 (ok, error_df)
    """
    df = class_df.copy()
    df["作答長度"] = df[ATT_COL_ANS].astype(str).apply(lambda x: len(x))
    bad = df[df["作答長度"] != n_questions].copy()
    if bad.empty:
        return True, pd.DataFrame()
    return False, bad[[ATT_COL_SEAT, ATT_COL_NAME, "作答長度"]]


def answer_str_to_correctness(ans_str: str) -> List[bool]:
    """
    規則：只有 '-' 算對，其餘一律算錯。
    """
    s = normalize_class_answer_str(ans_str)
    return [ch == "-" for ch in s]


def build_class_matrix_and_summary(q_df_en: pd.DataFrame, class_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    輸出兩張表（中文欄名）：
    1) 逐題矩陣：第1~6欄為題目資訊，第7欄起每位學生一欄（對/錯）
    2) 學生總表：每生答對/答錯/正確率/易中難錯題數
    """
    dfq = q_df_en.copy().sort_values("order_index").reset_index(drop=True)

    # 逐題矩陣 base 欄位（依你指定順序）
    matrix = pd.DataFrame({
        "題序": dfq["order_index"].apply(lambda x: int(x) if pd.notna(x) else None),
        "題號": dfq["label"].astype(str),
        "配分": dfq["score"],
        "難易度": dfq["difficulty"],
        "學習表現代碼": dfq["learning_code"],
        "題幹": dfq["stem"],
    })

    # 學生欄（依匯入順序）
    student_cols = []
    for _, row in class_df.iterrows():
        seat = str(row[ATT_COL_SEAT]).strip()
        name = str(row.get(ATT_COL_NAME, "")).strip()
        col_name = f"{seat}學生對錯"  # 你指定：1號學生對錯、2號學生對錯…
        # 若重名，補序號避免覆蓋
        base_col = col_name
        k = 2
        while col_name in student_cols:
            col_name = f"{base_col}_{k}"
            k += 1
        student_cols.append(col_name)

        correctness = answer_str_to_correctness(row[ATT_COL_ANS])
        # 長度已驗證一致，所以一定 == 題目數
        matrix[col_name] = ["對" if ok else "錯" for ok in correctness]

    # 學生總表
    summary_rows = []
    for idx, row in class_df.iterrows():
        seat = str(row[ATT_COL_SEAT]).strip()
        name = str(row.get(ATT_COL_NAME, "")).strip()
        correctness = answer_str_to_correctness(row[ATT_COL_ANS])

        total = len(dfq)
        correct_n = sum(1 for x in correctness if x)
        wrong_n = total - correct_n
        acc = correct_n / total if total > 0 else 0.0

        # 易/中/難錯題數（依題目表欄位）
        diff_series = dfq["difficulty"].fillna("").astype(str).tolist()
        diff_wrong = {"易": 0, "中": 0, "難": 0}
        for d, ok in zip(diff_series, correctness):
            if not ok and d in diff_wrong:
                diff_wrong[d] += 1

        summary_rows.append({
            "座號": seat,
            "姓名": name,
            "題目數": total,
            "答對題數": correct_n,
            "答錯題數": wrong_n,
            "正確率": acc,
            "易_錯題數": diff_wrong["易"],
            "中_錯題數": diff_wrong["中"],
            "難_錯題數": diff_wrong["難"],
        })

    summary_df = pd.DataFrame(summary_rows)
    return matrix, summary_df


def build_class_import_template(n_rows: int = 40) -> pd.DataFrame:
    """
    老師匯入範本：
    - 座號（字串，允許 01、1號）
    - 姓名（可空）
    - 作答字串（必填，長度需=題目數）
    """
    df = pd.DataFrame({
        ATT_COL_SEAT: ["" for _ in range(n_rows)],
        ATT_COL_NAME: ["" for _ in range(n_rows)],
        ATT_COL_ANS:  ["" for _ in range(n_rows)],
    })
    return df


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
        "class_msg": "",
        "class_matrix_df": None,
        "class_summary_df": None,
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
    st.session_state["class_msg"] = ""
    st.session_state["class_matrix_df"] = None
    st.session_state["class_summary_df"] = None


def run_analysis_single():
    q_df = safe_q_df()
    if q_df.empty:
        st.session_state["last_message"] = "❌ 尚未建立題目表：請先上傳 PDF 並完成解析。"
        st.session_state["last_result_df"] = None
        return

    correctness = parse_answer_string_sidebar(st.session_state.get("ans_str", ""))
    if len(correctness) == 0:
        st.session_state["last_message"] = "❌ 作答字串沒有讀到 '-' 或 'X'，請確認輸入格式（只接受 - 或 X）。"
        st.session_state["last_result_df"] = None
        return

    msg = "✅ 已完成作答分析。"
    if len(correctness) != len(q_df):
        msg += f"（提醒：作答長度 {len(correctness)} ≠ 題目數 {len(q_df)}，目前只分析前 {min(len(correctness), len(q_df))} 題）"

    df = build_results_df_single(q_df, correctness)
    st.session_state["last_result_df"] = df
    st.session_state["last_message"] = msg


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="考卷分析 MVP", layout="wide")
init_state()

st.title("考卷分析 MVP（可選字 PDF）")
st.caption("上傳 PDF → 拆題 → 題目表回填 → 單人/全班作答匯入分析")

# Sidebar：單一學生輸入（保留原功能）
st.sidebar.header("單一學生作答輸入（側欄）")
st.session_state["ans_str"] = st.sidebar.text_input(
    "作答字串（- 對 / X 錯）",
    value=st.session_state.get("ans_str", ""),
    placeholder="例：-------X-X-----XX-XXX--X",
)
st.sidebar.button("分析作答（單人）", type="primary", on_click=run_analysis_single)
st.sidebar.divider()
st.sidebar.button("Reset（清空）", on_click=reset_all)

# ============ 1) 上傳解析 ============
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

# 新檔案：只讀一次 bytes
if pdf_file is not None:
    sig = file_signature(pdf_file)
    if st.session_state.get("uploaded_sig") != sig:
        st.session_state["uploaded_sig"] = sig
        st.session_state["parsed_sig"] = None
        st.session_state["full_text"] = ""
        st.session_state["last_message"] = ""
        st.session_state["last_result_df"] = None
        st.session_state["apply_msg"] = ""
        st.session_state["class_msg"] = ""
        st.session_state["class_matrix_df"] = None
        st.session_state["class_summary_df"] = None

        st.session_state["pdf_bytes"] = pdf_file.getvalue()
        set_items([])
        set_q_df(safe_q_df())
        st.success("已上傳新檔案。")

# 解析（只做一次或手動）
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

# ============ 2) 題目表回填 ============
st.subheader("2) 題目表（老師可回填）")

q_df_en = safe_q_df()

if st.session_state.get("full_text") and not q_df_en.empty:
    st.write(f"題目數：**{len(q_df_en)}**（題號格式僅支援行首 `1.` 或 `1、`）")

    # 下載題目表（中文欄名）
    q_df_zh = q_df_to_teacher_df_zh(q_df_en)
    st.download_button(
        label="下載題目表 Excel（中文欄名，可回填）",
        data=to_excel_bytes(q_df_zh, sheet_name=SHEET_NAME_ZH),
        file_name="題目表_老師回填.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("可直接在下表回填：配分/難易度/學習表現代碼/備註（題幹不可改）。")

    edited_df = st.data_editor(
        q_df_en,
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

    # 回填後下載最新版（中文欄名）
    latest_zh = q_df_to_teacher_df_zh(safe_q_df())
    st.download_button(
        label="下載（含目前回填內容）題目表 Excel（中文欄名）",
        data=to_excel_bytes(latest_zh, sheet_name=SHEET_NAME_ZH),
        file_name="題目表_老師回填_最新版.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("文字預覽（前 1200 字）", expanded=False):
        text = st.session_state["full_text"]
        st.text(text[:1200] + ("…" if len(text) > 1200 else ""))

else:
    st.info("尚未解析：請先上傳 PDF 並完成解析，才會產生題目表。")

# ============ 3) 單人作答分析結果 ============
st.divider()
st.subheader("3) 單一學生作答分析結果（側欄按鈕觸發）")

msg = st.session_state.get("last_message", "")
df_single = st.session_state.get("last_result_df", None)

if msg:
    if df_single is None:
        st.error(msg)
    else:
        st.success(msg)

if df_single is not None:
    st.dataframe(df_single, use_container_width=True, height=360)
    st.download_button(
        label="下載單人作答結果 Excel（中文欄名）",
        data=to_excel_bytes(df_single, sheet_name="作答結果"),
        file_name="單人作答分析結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ============ 4) 全班作答匯入（你指定格式） ============
st.divider()
st.subheader("4) 全班作答匯入（每位學生一列、一格作答字串）")

q_df_en = safe_q_df()
if q_df_en.empty:
    st.info("請先完成：上傳 PDF → 解析拆題 → 建立題目表（必要）。")
else:
    st.markdown("### A) 下載老師匯入範本（中文欄名）")
    n_rows = st.number_input("範本預設列數（可留更多空行給老師填）", min_value=5, max_value=200, value=40, step=5)
    template_df = build_class_import_template(int(n_rows))
    st.download_button(
        label="下載全班作答匯入範本.xlsx",
        data=to_excel_bytes(template_df, sheet_name=ATT_SHEET_DEFAULT),
        file_name="全班作答匯入範本.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("規則：作答字串『只有 - 算對』，其餘任何符號都算錯。作答字串長度必須 = 題目數，否則不分析並回報座號。")

    st.markdown("### B) 上傳老師回填的全班作答 Excel → 產出兩張表（逐題矩陣 + 學生總表）")

    colA, colB, colC = st.columns([2, 1, 1])
    with colA:
        class_file = st.file_uploader("上傳全班作答 Excel（每位學生一列）", type=["xlsx"], key="class_uploader")
    with colB:
        sheet_name = st.text_input("Sheet 名稱", value=ATT_SHEET_DEFAULT)
    with colC:
        run_class = st.button("開始分析（全班）", type="primary", disabled=(class_file is None))

    if run_class and class_file is not None:
        try:
            class_df = read_class_answers_excel(class_file, sheet_name=sheet_name)

            # 移除完全空白列（座號與作答都空）
            class_df = class_df[
                ~((class_df[ATT_COL_SEAT].astype(str).str.strip() == "") & (class_df[ATT_COL_ANS].astype(str).str.strip() == ""))
            ].copy()

            if class_df.empty:
                st.error("❌ 讀不到有效資料：請確認至少有『座號』與『作答字串』。")
            else:
                n_q = len(q_df_en)
                ok, bad_df = validate_class_answers_length(class_df, n_q)
                if not ok:
                    st.error(f"❌ 作答字串長度不一致（必須等於題目數 {n_q}）。以下座號有問題：")
                    st.dataframe(bad_df, use_container_width=True, height=260)
                else:
                    matrix_df, summary_df = build_class_matrix_and_summary(q_df_en, class_df)
                    st.session_state["class_matrix_df"] = matrix_df
                    st.session_state["class_summary_df"] = summary_df
                    st.success(f"✅ 全班分析完成：學生數 {len(summary_df)}，題數 {len(matrix_df)}")

        except Exception as e:
            st.error(f"❌ 全班分析失敗：請確認 Excel 格式或 Sheet 名稱。（{type(e).__name__}）")

    # 顯示與下載
    matrix_df = st.session_state.get("class_matrix_df", None)
    summary_df = st.session_state.get("class_summary_df", None)

    if isinstance(matrix_df, pd.DataFrame) and isinstance(summary_df, pd.DataFrame):
        st.markdown("### 逐題矩陣（依你指定欄位順序）")
        st.dataframe(matrix_df, use_container_width=True, height=420)

        with st.expander("學生總表", expanded=False):
            st.dataframe(summary_df, use_container_width=True, height=320)

        out_xls = to_excel_bytes_multi({
            "逐題矩陣": matrix_df,
            "學生總表": summary_df,
        })
        st.download_button(
            label="下載全班分析結果 Excel（兩個 Sheet）",
            data=out_xls,
            file_name="全班作答分析結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
