import re
import io
import csv
import zipfile
import datetime

import streamlit as st
import openpyxl
import pdfplumber
import pandas as pd

# ----------------------------------------------------------------------
# 页面设置
# ----------------------------------------------------------------------
st.set_page_config(page_title="SOA 对账小助手", page_icon="💗", layout="wide")

# ----------------------------------------------------------------------
# 简洁粉色主题 CSS（干净、留白多、只用一个粉色色系）
# ----------------------------------------------------------------------
PINK = "#FF6FA5"
PINK_DARK = "#E24E88"
PINK_TEXT = "#7A3B58"
PINK_BG = "#FFF7FA"
PINK_LIGHT = "#FFE7F0"
PINK_BORDER = "#FFD6E6"

GREEN = "#2E9E6B"
GREEN_BG = "#E9F9F1"
ROSE = "#E24E70"
ROSE_BG = "#FFE7EC"
AMBER = "#C98A1B"
AMBER_BG = "#FFF3D9"

CUTE_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Baloo+2:wght@600;700;800&family=Quicksand:wght@400;500;600;700&display=swap');

html, body, [class*="css"] {{
    font-family: 'Quicksand', sans-serif;
}}

.stApp {{
    background: {PINK_BG};
}}

h1, h2, h3, h4 {{
    font-family: 'Baloo 2', sans-serif !important;
    color: {PINK_TEXT} !important;
}}

/* 顶部标题 */
.hero-card {{
    background: {PINK};
    border-radius: 20px;
    padding: 20px 26px;
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 16px;
}}
.hero-title {{
    font-family: 'Baloo 2', sans-serif;
    font-size: 24px;
    font-weight: 800;
    color: #FFFFFF;
    margin: 0;
}}
.hero-sub {{
    font-family: 'Quicksand', sans-serif;
    font-size: 13px;
    color: #FFE7F0;
    margin-top: 4px;
}}

/* 卡片 */
.plain-card {{
    background: #FFFFFF;
    border: 1px solid {PINK_BORDER};
    border-radius: 16px;
    padding: 18px 20px;
    margin-bottom: 16px;
}}

.summary-line {{
    font-size: 15px;
    font-weight: 700;
    color: {PINK_TEXT};
    background: {PINK_LIGHT};
    border-radius: 12px;
    padding: 12px 16px;
    margin-bottom: 14px;
}}

/* 统计方块：只用三种颜色（绿=没问题 / 玫红=有差异 / 琥珀=未对账） */
.metric-box {{
    border-radius: 14px;
    padding: 14px 8px;
    text-align: center;
    font-family: 'Baloo 2', sans-serif;
    border: 1px solid transparent;
}}
.metric-box .num {{ font-size: 24px; font-weight: 800; }}
.metric-box .lbl {{ font-size: 12px; margin-top: 2px; font-family: 'Quicksand', sans-serif; font-weight: 600; }}

.mb-ok    {{ background: {GREEN_BG}; color: {GREEN}; border-color: #CFF0DF; }}
.mb-issue {{ background: {ROSE_BG}; color: {ROSE}; border-color: #FFD2DC; }}
.mb-miss  {{ background: {AMBER_BG}; color: {AMBER}; border-color: #FBE6B8; }}

/* 按钮：单一粉色，简洁 */
.stButton>button, .stDownloadButton>button {{
    border-radius: 10px !important;
    border: none !important;
    background: {PINK} !important;
    color: white !important;
    font-family: 'Baloo 2', sans-serif !important;
    font-weight: 700 !important;
    padding: 8px 20px !important;
}}
.stButton>button:hover, .stDownloadButton>button:hover {{
    background: {PINK_DARK} !important;
}}

/* 文件上传框 */
[data-testid="stFileUploader"] {{
    background: {PINK_LIGHT};
    border: 1.5px dashed {PINK};
    border-radius: 14px;
    padding: 8px;
}}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {{ gap: 4px; }}
.stTabs [data-baseweb="tab"] {{
    background: {PINK_LIGHT};
    border-radius: 10px 10px 0 0;
    padding: 6px 14px;
    font-family: 'Baloo 2', sans-serif;
    font-weight: 700;
    color: {PINK_DARK};
}}
.stTabs [aria-selected="true"] {{
    background: {PINK} !important;
    color: white !important;
}}

/* 表格 */
[data-testid="stDataFrame"] {{
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid {PINK_BORDER};
}}

hr {{ border-top: 1px solid {PINK_BORDER}; }}
</style>
"""
st.markdown(CUTE_CSS, unsafe_allow_html=True)

# 简单Q版女孩头像（纯SVG矢量，小尺寸，不喧宾夺主）
GIRL_SVG = (
    '<svg width="52" height="58" viewBox="0 0 120 136" xmlns="http://www.w3.org/2000/svg">'
    '<ellipse cx="16" cy="74" rx="15" ry="27" fill="#8B5A46"/>'
    '<ellipse cx="104" cy="74" rx="15" ry="27" fill="#8B5A46"/>'
    '<path d="M8 52 L17 60 L8 68 Z" fill="#FFFFFF"/>'
    '<path d="M26 52 L17 60 L26 68 Z" fill="#FFFFFF"/>'
    '<circle cx="17" cy="60" r="3.2" fill="#FFE3EC"/>'
    '<path d="M94 52 L103 60 L94 68 Z" fill="#FFFFFF"/>'
    '<path d="M112 52 L103 60 L112 68 Z" fill="#FFFFFF"/>'
    '<circle cx="103" cy="60" r="3.2" fill="#FFE3EC"/>'
    '<path d="M18 128 Q60 98 102 128 L102 136 L18 136 Z" fill="#FFFFFF"/>'
    '<rect x="51" y="88" width="18" height="16" rx="7" fill="#FFDCC2"/>'
    '<circle cx="60" cy="63" r="41" fill="#FFE3CC"/>'
    '<path d="M18 56 Q60 4 102 56 Q102 28 60 22 Q18 28 18 56 Z" fill="#8B5A46"/>'
    '<path d="M28 44 Q60 27 92 44 L90 62 Q60 45 30 62 Z" fill="#8B5A46"/>'
    '<path d="M16 56 Q7 78 16 100 Q25 82 25 60 Z" fill="#8B5A46"/>'
    '<path d="M104 56 Q113 78 104 100 Q95 82 95 60 Z" fill="#8B5A46"/>'
    '<circle cx="44" cy="68" r="6.5" fill="#5A3E4D"/>'
    '<circle cx="76" cy="68" r="6.5" fill="#5A3E4D"/>'
    '<circle cx="46.5" cy="65.5" r="2.2" fill="#FFFFFF"/>'
    '<circle cx="78.5" cy="65.5" r="2.2" fill="#FFFFFF"/>'
    '<circle cx="33" cy="78" r="6.5" fill="#FFB0C8" opacity="0.75"/>'
    '<circle cx="87" cy="78" r="6.5" fill="#FFB0C8" opacity="0.75"/>'
    '<path d="M53 83 Q60 89 67 83" stroke="#C9758F" stroke-width="2.5" fill="none" stroke-linecap="round"/>'
    '</svg>'
)


def cute_header():
    st.markdown(
        '<div class="hero-card">'
        + GIRL_SVG
        + '<div>'
        + '<p class="hero-title">SOA 对账小助手</p>'
        + '<p class="hero-sub">上传 Excel + 一份或多份 SOA PDF，按发票号码和金额自动核对</p>'
        + '</div>'
        + '</div>',
        unsafe_allow_html=True,
    )


# ----------------------------------------------------------------------
# 密码保护
# ----------------------------------------------------------------------
try:
    SITE_PASSWORD = st.secrets["site_password"]
except Exception:
    SITE_PASSWORD = "changeme123"


def check_password():
    if st.session_state.get("authed"):
        return True

    cute_header()
    st.markdown('<div class="plain-card">', unsafe_allow_html=True)
    st.markdown("#### 🔒 请输入通行密码")
    st.caption("这个小工具只给被邀请的小伙伴使用哦")
    pw = st.text_input("密码", type="password", label_visibility="collapsed", placeholder="输入密码 Password")
    if st.button("进入"):
        if pw == SITE_PASSWORD:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("密码不对，再试一次")
    st.markdown('</div>', unsafe_allow_html=True)
    return False


if not check_password():
    st.stop()

# ----------------------------------------------------------------------
# 解析逻辑
# ----------------------------------------------------------------------
INVOICE_NO_PAREN_RE = re.compile(r"\((\d+)\)")
INVOICE_NO_DIGITS_RE = re.compile(r"\d{4,}")

PDF_LINE_RE = re.compile(
    r"^(\d{2}-[A-Z]{3}-\d{2,4})\s+"          # date
    r"([A-Z]{2})\s+"                          # type: IV / OR / CN / AD
    r"(\S+)\s+"                               # doc no
    r"(\d{2}-[A-Z]{3}-\d{2,4})\s+"           # due date
    r"(\S+)\s+"                               # ref / DO no
    r"([A-Z]{3})\s+"                          # currency
    r"([\d,]+\.\d{2})\s+"                    # debit
    r"([\d,]+\.\d{2})"                        # credit
)

# SOA 里的四种单据类型：IV=Invoice, OR=Official Receipt（付款/收据）,
# CN=Credit Note, AD=Adjustment。
# 只有 IV / CN / AD 会计入某张发票的金额；OR 是付款记录，不是发票金额的一部分，
# 一律忽略（不然同一张发票的余额会被payment冲销，误判成金额不一致）。
# CN 如果跟某张 invoice number 一样，会在这里自动被减掉（debit=0时用 -credit）。
INVOICE_TYPES = {"IV", "CN", "AD"}


def _to_float(x):
    try:
        if x is None or x == "":
            return None
        return float(str(x).replace(",", ""))
    except (TypeError, ValueError):
        return None


def normalize_invoice_no(raw):
    """把发票号码标准化成统一的比对key：
    只保留字母数字，去掉空格/括号/连字符等符号，纯数字的话再去掉前导0。
    这样 'TB C2 (20059764)' 和 '20059764' 会被认成同一张发票。"""
    if raw is None:
        return ""
    s = re.sub(r"[^0-9A-Za-z]", "", str(raw))
    if s.isdigit():
        s = str(int(s)) if s else "0"
    return s.upper()


def extract_invoice_no(ref):
    """从一段Reference文本里找出真正的发票号码，兼容各种写法：
    'TB C2 (20057646)'（括号内数字优先）、'TB C2-20057294'（连字符）、
    'TB C2 20057294'（空格）、纯数字 '20057294' 等。
    找不到明显数字串时，返回原文本去掉多余空格后的结果。"""
    if ref is None:
        return ""
    s = str(ref).strip()

    m = INVOICE_NO_PAREN_RE.findall(s)
    if m:
        return m[-1]

    m = INVOICE_NO_DIGITS_RE.findall(s)
    if m:
        return m[-1]

    return s


def parse_excel(file_obj):
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    header_idx = None
    col_map = {}
    for i, row in enumerate(rows):
        if not row:
            continue
        row_lower = [str(c).strip().lower() if c else "" for c in row]
        if "invoice date" in row_lower:
            header_idx = i
            for j, val in enumerate(row_lower):
                if val == "invoice date":
                    col_map["date"] = j
                elif val == "reference":
                    col_map["ref"] = j
                elif val == "gross":
                    col_map["amount"] = j
            break

    if header_idx is None or "date" not in col_map or "ref" not in col_map or "amount" not in col_map:
        raise ValueError(
            "无法在Excel中找到 'Invoice Date' / 'Reference' / 'Gross' 表头，"
            "请确认上传的是Xero导出的Invoice报表。"
        )

    groups = {}
    for row in rows[header_idx + 1:]:
        if not row or len(row) <= max(col_map.values()):
            continue
        inv_date = row[col_map["date"]]
        if not isinstance(inv_date, (datetime.datetime, datetime.date)):
            continue
        ref = row[col_map["ref"]]
        amount = _to_float(row[col_map["amount"]])
        if ref is None or amount is None:
            continue

        display_no = extract_invoice_no(ref)
        norm_key = normalize_invoice_no(display_no)
        if not norm_key:
            continue

        if norm_key not in groups:
            groups[norm_key] = {"date": inv_date, "amount": 0.0, "display": display_no}
        groups[norm_key]["amount"] += amount

    return groups


def parse_pdf(file_obj):
    groups = {}
    total_lines_seen = 0

    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw_line in text.split("\n"):
                line = raw_line.strip()
                m = PDF_LINE_RE.match(line)
                if not m:
                    continue
                total_lines_seen += 1
                date_str, typ, docno, due_str, ref, cur, debit_s, credit_s = m.groups()
                if typ not in INVOICE_TYPES:
                    continue
                debit = _to_float(debit_s) or 0.0
                credit = _to_float(credit_s) or 0.0
                amount = debit if debit > 0 else -credit

                try:
                    dt = datetime.datetime.strptime(date_str, "%d-%b-%y")
                except ValueError:
                    try:
                        dt = datetime.datetime.strptime(date_str, "%d-%b-%Y")
                    except ValueError:
                        continue

                norm_key = normalize_invoice_no(docno)
                if not norm_key:
                    continue

                if norm_key not in groups:
                    groups[norm_key] = {"date": dt, "amount": 0.0, "display": docno, "type": typ}
                groups[norm_key]["amount"] += amount

    if total_lines_seen == 0:
        raise ValueError(
            "无法从PDF中提取任何表格行。该PDF可能是扫描/图片版，"
            "本工具目前仅支持文字型PDF。"
        )

    return groups


def reconcile(excel_groups, pdf_groups):
    """只比对发票号码和金额，不比对日期。
    发票号码经过标准化后再匹配，格式上的小差异（比如 'TB C2 (20059764)' vs '20059764'）
    不影响匹配结果。"""
    all_keys = sorted(set(excel_groups) | set(pdf_groups))
    matched, mismatched, excel_only, pdf_only = [], [], [], []

    for k in all_keys:
        e = excel_groups.get(k)
        p = pdf_groups.get(k)

        if e and p:
            amt_diff = round(e["amount"] - p["amount"], 2)
            amt_match = abs(amt_diff) < 0.01
            display_no = e["display"] if e["display"] else p["display"]

            row = {
                "Invoice No": display_no,
                "Excel Date": e["date"].strftime("%Y-%m-%d"),
                "SOA Date": p["date"].strftime("%Y-%m-%d"),
                "Excel Amount": round(e["amount"], 2),
                "SOA Amount": round(p["amount"], 2),
                "Diff": amt_diff,
            }
            if amt_match:
                matched.append(row)
            else:
                row["问题"] = "金额不一致"
                mismatched.append(row)
        elif e and not p:
            excel_only.append({
                "Invoice No": e["display"],
                "Excel Date": e["date"].strftime("%Y-%m-%d"),
                "Excel Amount": round(e["amount"], 2),
            })
        elif p and not e:
            pdf_only.append({
                "Invoice No": p["display"],
                "SOA Date": p["date"].strftime("%Y-%m-%d"),
                "SOA Amount": round(p["amount"], 2),
            })

    return {
        "matched": matched,
        "mismatched": mismatched,
        "excel_only": excel_only,
        "pdf_only": pdf_only,
    }


STATUS_MATCHED = "✅ 匹配"
STATUS_MISMATCH = "⚠️ 金额不一致"
STATUS_EXCEL_ONLY = "📗 仅Excel有"
STATUS_PDF_ONLY = "📄 仅SOA有"

# 排序优先级：有问题的排最前面，最容易被看到
STATUS_PRIORITY = {
    STATUS_MISMATCH: 0,
    STATUS_EXCEL_ONLY: 1,
    STATUS_PDF_ONLY: 1,
    STATUS_MATCHED: 2,
}

STATUS_ROW_COLOR = {
    STATUS_MISMATCH: ROSE_BG,
    STATUS_EXCEL_ONLY: AMBER_BG,
    STATUS_PDF_ONLY: AMBER_BG,
    STATUS_MATCHED: GREEN_BG,
}


def build_master_table(results):
    """把matched/mismatched/excel_only/pdf_only合并成一张总表，只保留
    Invoice No / 金额 / 差额 / 状态，不比对日期，有问题的排最前面。"""
    rows = []
    for row in results["mismatched"]:
        rows.append({
            "状态": STATUS_MISMATCH,
            "Invoice No": row["Invoice No"],
            "Excel Amount": row["Excel Amount"],
            "SOA Amount": row["SOA Amount"],
            "差额 (Excel - SOA)": row["Diff"],
        })
    for row in results["excel_only"]:
        rows.append({
            "状态": STATUS_EXCEL_ONLY,
            "Invoice No": row["Invoice No"],
            "Excel Amount": row["Excel Amount"],
            "SOA Amount": None,
            "差额 (Excel - SOA)": None,
        })
    for row in results["pdf_only"]:
        rows.append({
            "状态": STATUS_PDF_ONLY,
            "Invoice No": row["Invoice No"],
            "Excel Amount": None,
            "SOA Amount": row["SOA Amount"],
            "差额 (Excel - SOA)": None,
        })
    for row in results["matched"]:
        rows.append({
            "状态": STATUS_MATCHED,
            "Invoice No": row["Invoice No"],
            "Excel Amount": row["Excel Amount"],
            "SOA Amount": row["SOA Amount"],
            "差额 (Excel - SOA)": row["Diff"],
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["_sort"] = df["状态"].map(STATUS_PRIORITY)
    df = df.sort_values(["_sort", "Invoice No"]).drop(columns="_sort").reset_index(drop=True)
    return df


def style_master_table(df):
    def _row_style(row):
        color = STATUS_ROW_COLOR.get(row["状态"], "#FFFFFF")
        return [f"background-color: {color}"] * len(row)

    styler = df.style.apply(_row_style, axis=1)
    fmt = {}
    for col in ("Excel Amount", "SOA Amount", "差额 (Excel - SOA)"):
        if col in df.columns:
            fmt[col] = "{:,.2f}"
    return styler.format(fmt, na_rep="—")


def human_summary(results, name=""):
    n_issue = len(results["mismatched"])
    n_missing = len(results["excel_only"]) + len(results["pdf_only"])
    n_ok = len(results["matched"])
    prefix = f"「{name}」：" if name else ""
    if n_issue == 0 and n_missing == 0:
        return f"🎉 {prefix}太棒了！全部 {n_ok} 张发票金额都对上了"
    parts = []
    if n_issue:
        parts.append(f"{n_issue} 张金额对不上")
    if n_missing:
        parts.append(f"{n_missing} 张未对账（只在一边出现）")
    return f"⚠️ {prefix}有 " + "，".join(parts) + "，需要你看一下"


def to_csv_bytes(results):
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Invoice No", "Status", "Excel Amount", "SOA Amount", "Diff", "Excel Date", "SOA Date"])
    for row in results["mismatched"]:
        writer.writerow([row["Invoice No"], row["问题"], row["Excel Amount"], row["SOA Amount"], row["Diff"], row["Excel Date"], row["SOA Date"]])
    for row in results["excel_only"]:
        writer.writerow([row["Invoice No"], "仅Excel有", row["Excel Amount"], "", "", row["Excel Date"], ""])
    for row in results["pdf_only"]:
        writer.writerow([row["Invoice No"], "仅SOA有", "", row["SOA Amount"], "", "", row["SOA Date"]])
    for row in results["matched"]:
        writer.writerow([row["Invoice No"], "Matched", row["Excel Amount"], row["SOA Amount"], row["Diff"], row["Excel Date"], row["SOA Date"]])
    return output.getvalue().encode("utf-8-sig")


def metric_box(col, number, label, css_class):
    col.markdown(
        f'<div class="metric-box {css_class}">'
        f'<div class="num">{number}</div>'
        f'<div class="lbl">{label}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


# ----------------------------------------------------------------------
# 主界面
# ----------------------------------------------------------------------
cute_header()

col_logout = st.columns([6, 1])[1]
with col_logout:
    if st.button("退出"):
        st.session_state["authed"] = False
        st.rerun()

st.markdown('<div class="plain-card">', unsafe_allow_html=True)
st.markdown("#### 📤 上传文件")
st.caption(
    "Excel 需为 Xero 导出的 Invoice 报表（含 Invoice Date / Reference / Gross 表头，"
    "其他栏位比如 Project 会自动忽略，不影响比对）。"
    "PDF 是供应商的 SOA，可以一次选多个（比如每个月一份），逐份跟同一份 Excel 对比。"
    "对比只看 **发票号码** 和 **金额**，不看日期。\n\n"
    "SOA 里的单据类型：只有 **IV（发票）** 和 **AD（调整）** 会计入金额，"
    "**CN（credit note）** 如果发票号码跟某张 invoice 一样，会自动从该发票金额里扣掉，"
    "**OR（收据/付款记录）** 会被完全忽略（不是发票金额的一部分）。"
)

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("Excel 文件 (.xlsx)", type=["xlsx", "xls"])
with col2:
    pdf_files = st.file_uploader(
        "SOA PDF 文件（可多选）",
        type=["pdf"],
        accept_multiple_files=True,
    )

if pdf_files:
    st.caption(f"已选择 {len(pdf_files)} 份 PDF：" + "、".join(f.name for f in pdf_files))

st.markdown('</div>', unsafe_allow_html=True)

if excel_file and pdf_files:
    if st.button("🔍 开始对比", type="primary"):
        with st.spinner("正在核对中..."):
            try:
                excel_groups = parse_excel(excel_file)
            except Exception as e:
                st.error(f"Excel解析失败: {e}")
                st.stop()

            results_by_pdf = {}
            errors = {}
            for pf in pdf_files:
                try:
                    pdf_groups = parse_pdf(pf)
                    results_by_pdf[pf.name] = reconcile(excel_groups, pdf_groups)
                except Exception as e:
                    errors[pf.name] = str(e)

            st.session_state["results_by_pdf"] = results_by_pdf
            st.session_state["parse_errors"] = errors

if "results_by_pdf" in st.session_state:
    results_by_pdf = st.session_state["results_by_pdf"]
    errors = st.session_state.get("parse_errors", {})

    if errors:
        for name, msg in errors.items():
            st.warning(f"⚠️ {name} 解析失败：{msg}")

    if results_by_pdf:
        st.markdown("<hr/>", unsafe_allow_html=True)
        st.markdown("### 总览")

        total_matched = sum(len(r["matched"]) for r in results_by_pdf.values())
        total_mismatch = sum(len(r["mismatched"]) for r in results_by_pdf.values())
        total_excel_only = sum(len(r["excel_only"]) for r in results_by_pdf.values())
        total_pdf_only = sum(len(r["pdf_only"]) for r in results_by_pdf.values())
        total_missing = total_excel_only + total_pdf_only
        total_all = total_matched + total_mismatch + total_missing

        overview_summary = human_summary({
            "matched": [None] * total_matched,
            "mismatched": [None] * total_mismatch,
            "excel_only": [None] * total_excel_only,
            "pdf_only": [None] * total_pdf_only,
        })
        st.markdown(
            f'<div class="summary-line">{overview_summary}　'
            f'（共处理 {len(results_by_pdf)} 份 SOA，{total_all} 张发票）</div>',
            unsafe_allow_html=True,
        )

        c1, c2, c3 = st.columns(3)
        metric_box(c1, total_matched, "✅ 匹配", "mb-ok")
        metric_box(c2, total_mismatch, "⚠️ 金额不一致", "mb-issue")
        metric_box(c3, total_missing, "📌 未对账", "mb-miss")

        # 汇总所有需要关注的发票（跨所有PDF），不用逐个点进去找
        all_issue_rows = []
        for name, r in results_by_pdf.items():
            mt = build_master_table(r)
            if mt.empty:
                continue
            issues = mt[mt["状态"] != STATUS_MATCHED].copy()
            if not issues.empty:
                issues.insert(0, "SOA 文件", name)
                all_issue_rows.append(issues)

        if all_issue_rows:
            st.markdown("#### 需要你关注的发票（所有 SOA 汇总）")
            combined = pd.concat(all_issue_rows, ignore_index=True)
            combined["_sort"] = combined["状态"].map(STATUS_PRIORITY)
            combined = combined.sort_values(["_sort", "SOA 文件", "Invoice No"]).drop(columns="_sort")
            st.dataframe(style_master_table(combined.reset_index(drop=True)), use_container_width=True, hide_index=True)
        else:
            st.success("🎉 所有 SOA 都对上了，完全没有需要担心的地方！")

        # 每份PDF的小结表
        with st.expander("每份 SOA 的比对小结"):
            summary_rows = []
            for name, r in results_by_pdf.items():
                summary_rows.append({
                    "SOA 文件": name,
                    "✅ 匹配": len(r["matched"]),
                    "⚠️ 金额不一致": len(r["mismatched"]),
                    "📗 仅Excel有": len(r["excel_only"]),
                    "📄 仅SOA有": len(r["pdf_only"]),
                })
            st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        # 打包下载全部报告
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for name, r in results_by_pdf.items():
                safe_name = re.sub(r"[^\w\-.]", "_", name.rsplit(".", 1)[0])
                zf.writestr(f"{safe_name}_对账结果.csv", to_csv_bytes(r))
        st.download_button(
            "📦 打包下载全部对账报告 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="soa_对账报告_全部.zip",
            mime="application/zip",
        )

        st.markdown("<hr/>", unsafe_allow_html=True)
        st.markdown("### 每份 PDF 的详细结果")

        tabs = st.tabs([f"📄 {name}" for name in results_by_pdf.keys()])
        for tab, (name, r) in zip(tabs, results_by_pdf.items()):
            with tab:
                st.markdown(
                    f'<div class="summary-line">{human_summary(r)}</div>',
                    unsafe_allow_html=True,
                )

                cc1, cc2, cc3 = st.columns(3)
                metric_box(cc1, len(r["matched"]), "✅ 匹配", "mb-ok")
                metric_box(cc2, len(r["mismatched"]), "金额不一致", "mb-issue")
                metric_box(cc3, len(r["excel_only"]) + len(r["pdf_only"]), "未对账", "mb-miss")

                st.write("")
                fc1, fc2, fc3 = st.columns([2, 2, 2])
                with fc1:
                    only_issues = st.checkbox(
                        "只看有问题的",
                        value=bool(r["mismatched"] or r["excel_only"] or r["pdf_only"]),
                        key=f"only_issues_{name}",
                    )
                with fc2:
                    search_term = st.text_input(
                        "🔎 搜索发票号码", key=f"search_{name}", placeholder="输入 Invoice No..."
                    )
                with fc3:
                    st.download_button(
                        "📥 导出CSV",
                        data=to_csv_bytes(r),
                        file_name=f"{name}_对账结果.csv",
                        mime="text/csv",
                        key=f"dl_{name}",
                    )

                master_df = build_master_table(r)
                if master_df.empty:
                    st.info("这份PDF没有解析出任何发票行，可能格式不一样，欢迎把样本发过来看看～")
                else:
                    view_df = master_df
                    if only_issues:
                        view_df = view_df[view_df["状态"] != STATUS_MATCHED]
                    if search_term:
                        view_df = view_df[view_df["Invoice No"].astype(str).str.contains(search_term, case=False, na=False)]

                    if view_df.empty:
                        st.caption("没有符合条件的发票，换个搜索词或取消勾选试试？")
                    else:
                        st.dataframe(
                            style_master_table(view_df.reset_index(drop=True)),
                            use_container_width=True,
                            hide_index=True,
                            height=min(42 * (len(view_df) + 1) + 10, 560),
                        )
                        st.caption(
                            f"图例：粉红=金额不一致　琥珀=只在一边出现　绿色=匹配没问题　"
                            f"（共显示 {len(view_df)} / {len(master_df)} 条）"
                        )

st.markdown(
    "<div style='text-align:center; margin-top:26px; color:#B98BA3; font-size:12px;'>"
    "made with 💗"
    "</div>",
    unsafe_allow_html=True,
)
