import io
import re
import math
import os
from dataclasses import dataclass
from typing import List, Tuple, Optional, Any, Dict

import pandas as pd
import pdfplumber
import streamlit as st


# =========================================================
# 題目表（中文欄名對照）：下載 / 上傳回填 / 系統內部
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

# 老師回填時「會套用」的欄位（不覆蓋題序/題號/題幹）
FILL_COLS_EN = ["section", "score", "difficulty", "learning_code", "note"]

# 學習表現代碼對照表（repo 內檔案）
CODE_MAP_XLSX = "learning_code_map.xlsx"
CODE_MAP_CSV = "learning_code_map.csv"


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
# Utilities
# -----------------------------
def file_signature(uploaded_file) -> str:
    return f"{uploaded_file.name}:{uploaded_file.size}"


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "sheet1") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=str(sheet_name)[:31])
    return output.getvalue()


def to_excel_bytes_multi(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=str(name)[:31])
    return output.getvalue()


def safe_q_df() -> pd.DataFrame:
    df = st.session_state.get("q_df", None)
    if isinstance(df, pd.DataFrame):
        return df
    return pd.DataFrame(columns=TEACHER_COLS_EN)


def set_q_df(df: Any) -> None:
    st.session_state["q_df"] = df if isinstance(df, pd.DataFrame) else safe_q_df()


def q_df_to_teacher_df_zh(q_df_en: pd.DataFrame) -> pd.DataFrame:
    df = q_df_en.copy()
    for c in TEACHER_COLS_EN:
        if c not in df.columns:
            df[c] = "" if c != "score" else None
    df = df[TEACHER_COLS_EN]
    return df.rename(columns=COL_MAP_EN2ZH)


def teacher_df_to_q_df_en(df_any: pd.DataFrame) -> pd.DataFrame:
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
    df["score"] = pd.to_numeric(df["score"], errors="coerce")
    df["learning_code"] = df["learning_code"].astype(str).replace({"nan": ""}).str.strip()
    df["difficulty"] = df["difficulty"].astype(str).replace({"nan": ""}).str.strip()
    df["stem"] = df["stem"].astype(str).replace({"nan": ""})
    return df


# -----------------------------
# Load learning code map (external file)
# -----------------------------
@st.cache_data(show_spinner=False)
def load_learning_code_map_from_repo() -> pd.DataFrame:
    """
    從 repo 根目錄讀取學習表現代碼對照表（xlsx 優先，其次 csv）。
    必須包含欄位：學習表現代碼、學習表現
    """
    if os.path.exists(CODE_MAP_XLSX):
        df = pd.read_excel(CODE_MAP_XLSX, engine="openpyxl")
    elif os.path.exists(CODE_MAP_CSV):
        df = pd.read_csv(CODE_MAP_CSV, encoding="utf-8")
    else:
        return pd.DataFrame(columns=["學習表現代碼", "學習表現"])

    df.columns = [str(c).strip() for c in df.columns]
    if "學習表現代碼" not in df.columns or "學習表現" not in df.columns:
        return pd.DataFrame(columns=["學習表現代碼", "學習表現"])

    df = df[["學習表現代碼", "學習表現"]].copy()
    df["學習表現代碼"] = df["學習表現代碼"].astype(str).str.strip()
    df["學習表現"] = df["學習表現"].astype(str).replace({"nan": ""})
    df = df[df["學習表現代碼"] != ""].drop_duplicates("學習表現代碼")
    return df.reset_index(drop=True)


def learning_code_to_desc_map(code_map_df: pd.DataFrame) -> Dict[str, str]:
    if code_map_df is None or code_map_df.empty:
        return {}
    return dict(zip(code_map_df["學習表現代碼"].astype(str), code_map_df["學習表現"].astype(str)))


# -----------------------------
# PDF parse & split
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


# -----------------------------
# 題目表：回填上傳套用
# -----------------------------
def read_teacher_excel(uploaded_xlsx, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(uploaded_xlsx, sheet_name=sheet_name, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def apply_teacher_fill(current_q_en: pd.DataFrame, uploaded_df_any: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    cur = teacher_df_to_q_df_en(current_q_en)
    fil = teacher_df_to_q_df_en(uploaded_df_any)

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
    key_zh = COL_MAP_EN2ZH.get(key, key)
    return merged, f"✅ 已套用回填資料（對齊欄位：{key_zh}）。題目數 {len(merged)}，至少有回填資料的題數約 {applied_rows}。"


# -----------------------------
# 全班作答：只有 '-' 算對，其餘算錯；強制一致
# -----------------------------
def normalize_class_answer_str(ans: str) -> str:
    if ans is None:
        return ""
    s = str(ans)
    s = re.sub(r"\s+", "", s)
    return s


def read_class_answers_excel(uploaded_xlsx, sheet_name: str = ATT_SHEET_DEFAULT) -> pd.DataFrame:
    df = pd.read_excel(uploaded_xlsx, sheet_name=sheet_name, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

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
    df = class_df.copy()
    df["作答長度"] = df[ATT_COL_ANS].astype(str).apply(lambda x: len(x))
    bad = df[df["作答長度"] != n_questions].copy()
    if bad.empty:
        return True, pd.DataFrame()
    return False, bad[[ATT_COL_SEAT, ATT_COL_NAME, "作答長度"]]


def answer_str_to_correctness(ans_str: str) -> List[bool]:
    s = normalize_class_answer_str(ans_str)
    return [ch == "-" for ch in s]


def build_class_matrix_and_summary(q_df_en: pd.DataFrame, class_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    dfq = teacher_df_to_q_df_en(q_df_en).sort_values("order_index").reset_index(drop=True)

    score_list = dfq["score"].fillna(0.0).astype(float).tolist()
    total_score = float(sum(score_list))

    # 逐題矩陣 base（欄位順序固定）
    matrix = pd.DataFrame({
        "題序": dfq["order_index"].apply(lambda x: int(x) if pd.notna(x) else None),
        "題號": dfq["label"].astype(str),
        "配分": dfq["score"],
        "難易度": dfq["difficulty"],
        "學習表現代碼": dfq["learning_code"],
        "題幹": dfq["stem"],
    })

    student_cols = []
    summary_rows = []
    diff_series = dfq["difficulty"].fillna("").astype(str).tolist()

    for _, row in class_df.iterrows():
        seat = str(row[ATT_COL_SEAT]).strip()
        name = str(row.get(ATT_COL_NAME, "")).strip()

        base_col = f"{seat}號學生對錯"
        col_name = base_col
        k = 2
        while col_name in student_cols:
            col_name = f"{base_col}_{k}"
            k += 1
        student_cols.append(col_name)

        correctness = answer_str_to_correctness(row[ATT_COL_ANS])
        matrix[col_name] = ["對" if ok else "錯" for ok in correctness]

        # 成績（依配分加總）
        student_score = 0.0
        for ok, sc in zip(correctness, score_list):
            if ok:
                student_score += float(sc)

        total = len(dfq)
        correct_n = sum(1 for x in correctness if x)
        wrong_n = total - correct_n
        acc = correct_n / total if total > 0 else 0.0

        diff_wrong = {"易": 0, "中": 0, "難": 0}
        for d, ok in zip(diff_series, correctness):
            if not ok and d in diff_wrong:
                diff_wrong[d] += 1

        summary_rows.append({
            "座號": seat,
            "姓名": name,
            "題目數": total,
            "總分": total_score,
            "成績": student_score,
            "排名": None,  # 先占位
            "答對題數": correct_n,
            "答錯題數": wrong_n,
            "正確率": acc,
            "易_錯題數": diff_wrong["易"],
            "中_錯題數": diff_wrong["中"],
            "難_錯題數": diff_wrong["難"],
        })

    summary_df = pd.DataFrame(summary_rows)
    if not summary_df.empty:
        summary_df["排名"] = summary_df["成績"].rank(method="min", ascending=False).astype(int)
        summary_df = summary_df.sort_values(["排名", "座號"]).reset_index(drop=True)

    return matrix, summary_df


def build_class_import_template_fixed() -> pd.DataFrame:
    # 固定提供 40 列空白，老師自行複製列即可
    n_rows = 40
    return pd.DataFrame({
        ATT_COL_SEAT: ["" for _ in range(n_rows)],
        ATT_COL_NAME: ["" for _ in range(n_rows)],
        ATT_COL_ANS:  ["" for _ in range(n_rows)],
    })


# -----------------------------
# 班級總體分析：1)組距 2)五標 3)各題答對率 4)學習表現代碼分析
# -----------------------------
def build_score_distribution(scores: pd.Series, total_score: float) -> pd.DataFrame:
    scores = scores.fillna(0.0).astype(float)
    max_score = float(scores.max()) if len(scores) else 0.0
    upper = max(total_score, max_score)
    upper = math.ceil(upper / 10.0) * 10.0
    edges = [x for x in range(0, int(upper) + 10, 10)]
    if len(edges) < 2:
        edges = [0, 10]

    cats = pd.cut(scores, bins=edges, right=False, include_lowest=True)
    dist = cats.value_counts().sort_index()

    rows = []
    for interval, cnt in dist.items():
        left = int(interval.left)
        right = int(interval.right) - 1
        label = f"{left}-{right}"
        rows.append({"成績組距(10分一組)": label, "人數": int(cnt)})
    return pd.DataFrame(rows)


def build_five_standards(scores: pd.Series) -> pd.DataFrame:
    s = scores.fillna(0.0).astype(float)
    if len(s) == 0:
        return pd.DataFrame([{"項目": x, "分數": 0.0} for x in ["頂標(88%)", "前標(75%)", "均標(50%)", "後標(25%)", "底標(12%)", "平均分"]])

    return pd.DataFrame([
        {"項目": "頂標(88%)", "分數": float(s.quantile(0.88))},
        {"項目": "前標(75%)", "分數": float(s.quantile(0.75))},
        {"項目": "均標(50%)", "分數": float(s.quantile(0.50))},
        {"項目": "後標(25%)", "分數": float(s.quantile(0.25))},
        {"項目": "底標(12%)", "分數": float(s.quantile(0.12))},
        {"項目": "平均分", "分數": float(s.mean())},
    ])


def build_question_correct_rate(matrix_df: pd.DataFrame) -> pd.DataFrame:
    if matrix_df is None or matrix_df.empty:
        return pd.DataFrame(columns=["題序", "題號", "配分", "難易度", "學習表現代碼", "答對率", "答對人數", "作答人數"])

    base_cols = ["題序", "題號", "配分", "難易度", "學習表現代碼"]
    student_cols = [c for c in matrix_df.columns if c not in ["題序", "題號", "配分", "難易度", "學習表現代碼", "題幹"]]
    if not student_cols:
        return pd.DataFrame(columns=base_cols + ["答對率", "答對人數", "作答人數"])

    correct_cnt = (matrix_df[student_cols] == "對").sum(axis=1).astype(int)
    total_cnt = matrix_df[student_cols].notna().sum(axis=1).astype(int)
    rate = (correct_cnt / total_cnt.replace(0, pd.NA)).fillna(0.0)

    out = matrix_df[base_cols].copy()
    out["答對率"] = rate
    out["答對人數"] = correct_cnt
    out["作答人數"] = total_cnt
    return out


def split_learning_codes(s: str) -> List[str]:
    """
    題目表的學習表現代碼欄可能是：
      "n-1-1" 或 "n-1-1;n-2-3"
    統一拆成 list（去空白、去空項）
    """
    if s is None:
        return []
    text = str(s).strip()
    if text == "" or text.lower() == "nan":
        return []
    parts = [p.strip() for p in text.split(";")]
    return [p for p in parts if p != ""]


def build_learning_code_analysis(
    q_df_en: pd.DataFrame,
    qrate_df: pd.DataFrame,
    code_desc_map: Dict[str, str],
) -> pd.DataFrame:
    """
    第4區段：學習表現代碼分析
    欄位：學習表現代碼、學習表現（完整文字）、題目數、答對率
    答對率：該代碼下所有題目的「各題答對率」取平均
    """
    dfq = teacher_df_to_q_df_en(q_df_en).sort_values("order_index").reset_index(drop=True)
    if dfq.empty or qrate_df is None or qrate_df.empty:
        return pd.DataFrame(columns=["學習表現代碼", "學習表現", "題目數", "答對率"])

    # 建立：題序 → 該題答對率
    rate_by_order = dict(zip(qrate_df["題序"], qrate_df["答對率"]))

    code_to_orders: Dict[str, List[int]] = {}
    for _, row in dfq.iterrows():
        order = int(row["order_index"]) if pd.notna(row["order_index"]) else None
        if order is None:
            continue
        codes = split_learning_codes(row.get("learning_code", ""))
        for c in codes:
            code_to_orders.setdefault(c, []).append(order)

    rows = []
    for code, orders in sorted(code_to_orders.items(), key=lambda x: x[0]):
        rates = []
        for o in orders:
            if o in rate_by_order:
                rates.append(float(rate_by_order[o]))
        avg_rate = float(sum(rates) / len(rates)) if rates else 0.0

        desc = code_desc_map.get(code, "")
        rows.append({
            "學習表現代碼": code,
            "學習表現": desc,          # 必須完整文字（由對照表提供）
            "題目數": len(orders),
            "答對率": avg_rate,
        })

    out = pd.DataFrame(rows)
    # 讓老師更容易看：題目數多的排前面
    if not out.empty:
        out = out.sort_values(["題目數", "學習表現代碼"], ascending=[False, True]).reset_index(drop=True)
    return out


def stack_sections_with_headers(sections: List[Tuple[str, pd.DataFrame]]) -> pd.DataFrame:
    """
    把多段表堆疊成單一 DataFrame，段落間插入空白列，
    並在每段前插入：
      1) 區段標題列
      2) 欄位標題列（用原本 df.columns）
      3) 資料列
    """
    blocks = []
    max_cols = 1
    for _, df in sections:
        max_cols = max(max_cols, len(df.columns) if df is not None else 1)

    cols = [f"欄{i}" for i in range(1, max_cols + 1)]

    def make_row(values: List[Any]) -> pd.DataFrame:
        row = {cols[i]: values[i] for i in range(min(len(values), max_cols))}
        return pd.DataFrame([row], columns=cols)

    blank = pd.DataFrame([{}], columns=cols)

    for idx, (title, df) in enumerate(sections):
        blocks.append(make_row([title]))

        if df is None or df.empty:
            blocks.append(make_row(["（無資料）"]))
        else:
            blocks.append(make_row(df.columns.tolist()))
            for _, r in df.iterrows():
                blocks.append(make_row(r.tolist()))

        if idx != len(sections) - 1:
            blocks.append(blank)

    return pd.concat(blocks, ignore_index=True).fillna("")


# -----------------------------
# Session init
# -----------------------------
def init_state():
    defaults = {
        "pdf_bytes": None,
        "uploaded_sig": None,
        "parsed_sig": None,
        "full_text": "",
        "q_df": safe_q_df(),
        "apply_msg": "",
        "class_matrix_df": None,
        "class_summary_df": None,
        "class_overall_df": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    if not isinstance(st.session_state.get("q_df", None), pd.DataFrame):
        st.session_state["q_df"] = safe_q_df()


def hard_reset_page():
    """
    需求：按下 Reset = 整個頁面回到最初始狀態
    作法：清空 session_state + rerun
    """
    st.session_state.clear()
    st.rerun()


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="考卷分析 MVP", layout="wide")
init_state()

st.title("考卷分析 MVP（可選字 PDF）")
st.caption("上傳 PDF → 拆題 → 題目表回填/套用 → 全班作答匯入 → 下載報表（含班級總體分析＋學習表現分析）")

# Reset：整頁回到初始狀態
if st.button("Reset（重新整理，回到初始狀態）"):
    hard_reset_page()

# 讀取「學習表現代碼對照表」（repo 外部檔案）
code_map_df = load_learning_code_map_from_repo()
code_desc_map = learning_code_to_desc_map(code_map_df)

with st.expander("學習表現代碼對照表狀態（從 repo 讀取）", expanded=False):
    if code_map_df.empty:
        st.warning("找不到 learning_code_map.xlsx 或 learning_code_map.csv，或欄位不正確。第4區段『學習表現』會是空白。")
        st.info("請在 repo 根目錄新增對照表檔案，欄位必須包含：學習表現代碼、學習表現。")
    else:
        st.success(f"已載入對照表：{len(code_map_df)} 筆代碼")
        st.dataframe(code_map_df.head(20), use_container_width=True, height=260)

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

if pdf_file is not None:
    sig = file_signature(pdf_file)
    if st.session_state.get("uploaded_sig") != sig:
        st.session_state["uploaded_sig"] = sig
        st.session_state["parsed_sig"] = None
        st.session_state["full_text"] = ""
        st.session_state["apply_msg"] = ""
        st.session_state["class_matrix_df"] = None
        st.session_state["class_summary_df"] = None
        st.session_state["class_overall_df"] = None
        st.session_state["pdf_bytes"] = pdf_file.getvalue()
        st.success("已上傳新檔案。")

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
        q_df = build_teacher_fill_df(items)
        set_q_df(q_df)

        st.session_state["parsed_sig"] = st.session_state.get("uploaded_sig")
    st.success("解析完成！已建立『題目表（老師可回填）』。")

# ============ 2) 題目表：下載 / 網頁回填 / 上傳回填套用 ============
st.subheader("2) 題目表（老師可回填）")

q_df_en = safe_q_df()
if st.session_state.get("full_text") and not q_df_en.empty:
    st.write(f"題目數：**{len(q_df_en)}**（題號格式僅支援行首 `1.` 或 `1、`）")

    q_df_zh = q_df_to_teacher_df_zh(q_df_en)
    st.download_button(
        label="下載題目表 Excel（中文欄名，可回填）",
        data=to_excel_bytes(q_df_zh, sheet_name=SHEET_NAME_ZH),
        file_name="題目表_老師回填.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("### 上傳老師回填題目表 → 套用到系統")
    colA, colB, colC = st.columns([2, 1, 1])
    with colA:
        fill_file = st.file_uploader("上傳題目表回填 Excel", type=["xlsx"], key="q_fill_uploader")
    with colB:
        fill_sheet = st.text_input("Sheet 名稱", value=SHEET_NAME_ZH)
    with colC:
        apply_btn = st.button("套用回填", type="primary", disabled=(fill_file is None))

    if apply_btn and fill_file is not None:
        try:
            try:
                uploaded_df = read_teacher_excel(fill_file, sheet_name=fill_sheet)
            except Exception:
                fallback = SHEET_NAME_EN if fill_sheet == SHEET_NAME_ZH else SHEET_NAME_ZH
                uploaded_df = read_teacher_excel(fill_file, sheet_name=fallback)

            merged, msg = apply_teacher_fill(q_df_en, uploaded_df)
            set_q_df(merged)
            st.session_state["apply_msg"] = msg
        except Exception as e:
            st.session_state["apply_msg"] = f"❌ 套用失敗：請確認檔案格式與 sheet 名稱。（{type(e).__name__}）"

    if st.session_state.get("apply_msg"):
        if str(st.session_state["apply_msg"]).startswith("✅"):
            st.success(st.session_state["apply_msg"])
        else:
            st.error(st.session_state["apply_msg"])

    st.markdown("### 直接在網頁回填（可選）")
    edited_df = st.data_editor(
        teacher_df_to_q_df_en(safe_q_df()),
        use_container_width=True,
        height=420,
        num_rows="fixed",
        column_config={
            "order_index": st.column_config.NumberColumn("題序", disabled=True),
            "label": st.column_config.TextColumn("題號", disabled=True),
            "section": st.column_config.TextColumn("題型"),
            "score": st.column_config.NumberColumn("配分", min_value=0, step=1),
            "difficulty": st.column_config.SelectboxColumn("難易度", options=["", "易", "中", "難"]),
            "learning_code": st.column_config.TextColumn("學習表現代碼（多個用 ;）"),
            "note": st.column_config.TextColumn("備註"),
            "stem": st.column_config.TextColumn("題幹", disabled=True),
        },
        key="q_df_editor",
    )
    set_q_df(edited_df)

    latest_zh = q_df_to_teacher_df_zh(safe_q_df())
    st.download_button(
        label="下載（含目前回填內容）題目表 Excel（中文欄名）",
        data=to_excel_bytes(latest_zh, sheet_name=SHEET_NAME_ZH),
        file_name="題目表_老師回填_最新版.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("尚未解析：請先上傳 PDF 並完成解析，才會產生題目表。")

# ============ 3) 全班作答匯入 ============
st.divider()
st.subheader("3) 全班作答匯入（每位學生一列、一格作答字串）")

q_df_en = safe_q_df()
if q_df_en.empty:
    st.info("請先完成：上傳 PDF → 解析拆題 → 建立題目表（必要）。")
else:
    st.markdown("### A) 下載全班作答匯入範本（固定提供）")
    template_df = build_class_import_template_fixed()
    st.download_button(
        label="下載全班作答匯入範本.xlsx",
        data=to_excel_bytes(template_df, sheet_name=ATT_SHEET_DEFAULT),
        file_name="全班作答匯入範本.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("規則：作答字串『只有 - 算對』，其餘任何符號都算錯。作答字串長度必須 = 題目數，否則不分析並回報座號。")

    st.markdown("### B) 上傳老師回填的全班作答 Excel → 產出：逐題矩陣 / 學生總表 / 班級總體分析（含學習表現分析）")

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

            # 移除完全空白列
            class_df = class_df[
                ~((class_df[ATT_COL_SEAT].astype(str).str.strip() == "") & (class_df[ATT_COL_ANS].astype(str).str.strip() == ""))
            ].copy()

            if class_df.empty:
                st.error("❌ 讀不到有效資料：請確認至少有『座號』與『作答字串』。")
            else:
                qdf_norm = teacher_df_to_q_df_en(q_df_en).sort_values("order_index").reset_index(drop=True)
                n_q = len(qdf_norm)

                ok, bad_df = validate_class_answers_length(class_df, n_q)
                if not ok:
                    st.error(f"❌ 作答字串長度不一致（必須等於題目數 {n_q}）。以下座號有問題：")
                    st.dataframe(bad_df, use_container_width=True, height=260)
                else:
                    matrix_df, summary_df = build_class_matrix_and_summary(qdf_norm, class_df)

                    # 班級總體分析
                    total_score = float(qdf_norm["score"].fillna(0.0).sum())
                    scores = summary_df["成績"].astype(float) if "成績" in summary_df.columns else pd.Series(dtype=float)

                    dist_df = build_score_distribution(scores, total_score)
                    five_df = build_five_standards(scores)
                    qrate_df = build_question_correct_rate(matrix_df)

                    # 第4段：學習表現代碼分析（需要 q_df + qrate + 對照表）
                    learn_df = build_learning_code_analysis(qdf_norm, qrate_df, code_desc_map)

                    overall_df = stack_sections_with_headers([
                        ("1) 班級成績組距（10分一組）", dist_df),
                        ("2) 班級五標（頂/前/均/後/底）", five_df),
                        ("3) 各題目班級答對率", qrate_df),
                        ("4) 學習表現代碼分析", learn_df),
                    ])

                    st.session_state["class_matrix_df"] = matrix_df
                    st.session_state["class_summary_df"] = summary_df
                    st.session_state["class_overall_df"] = overall_df

                    st.success(f"✅ 全班分析完成：學生數 {len(summary_df)}，題數 {len(matrix_df)}")

        except Exception as e:
            st.error(f"❌ 全班分析失敗：請確認 Excel 格式或 Sheet 名稱。（{type(e).__name__}）")

    matrix_df = st.session_state.get("class_matrix_df", None)
    summary_df = st.session_state.get("class_summary_df", None)
    overall_df = st.session_state.get("class_overall_df", None)

    if isinstance(matrix_df, pd.DataFrame) and isinstance(summary_df, pd.DataFrame) and isinstance(overall_df, pd.DataFrame):
        st.markdown("### 逐題矩陣")
        st.dataframe(matrix_df, use_container_width=True, height=420)

        with st.expander("學生總表（含：成績、排名）", expanded=False):
            st.dataframe(summary_df, use_container_width=True, height=340)

        with st.expander("班級總體分析（含：學習表現代碼分析）", expanded=False):
            st.dataframe(overall_df, use_container_width=True, height=420)

        out_xls = to_excel_bytes_multi({
            "逐題矩陣": matrix_df,
            "學生總表": summary_df,
            "班級總體分析": overall_df,
        })
        st.download_button(
            label="下載全班分析結果 Excel（3個 Sheet）",
            data=out_xls,
            file_name="全班作答分析結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
