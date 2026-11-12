"""
Combine Partner LFs  -  standalone Streamlit app (single file)

Run:    streamlit run combine_app.py
Needs:  streamlit, pandas, openpyxl

What it does
* Matches rows by indicator code (OC 1.2, OP 2.1.3 ...) - not by row position.
* Fills one achievement column per partner (creates the columns if the overall LF has none).
* Combines every partner's "Reason for Programmatic deviation" into one column:
      [Partner A]: text ...
* Total: SUM for count indicators; pooled numerator/denominator for % indicators,
  or a MANUAL flag in "Total Check" when a partner typed a % without a formula.
"""

from __future__ import annotations

import io
import re
from copy import copy
from dataclasses import dataclass, field

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

CODE_RE = re.compile(r"^(OC|OP|OI|IM|G)\s*\d+(\.\d+)*$", re.IGNORECASE)
REF = r"\$?[A-Z]{1,3}\$?\d+"
NUM = r"\d+(?:\.\d+)?"
FRACTION_RE = re.compile(
    rf"^=\s*\(?\s*({REF}|{NUM})\s*\)?\s*/\s*\(?\s*({REF}|{NUM})\s*\)?\s*$", re.IGNORECASE
)
PCT_WORDS = ("percentage", "proportion", "%", "rate of", "ratio")
DEFAULT_REASON_HEADER = "Reason for Programmatic deviation"


# --------------------------------------------------------------------------- helpers
def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def _norm_code(s) -> str:
    t = re.sub(r"\s+", " ", str(s or "")).strip().upper()
    return re.sub(r"^([A-Z]{1,3})\s*(\d)", r"\1 \2", t)      # OC2.1.6 -> OC 2.1.6


def _is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _clean_text(v) -> str:
    if v is None:
        return ""
    t = str(v).replace("\r\n", "\n").strip()
    t = re.sub(r"\n{3,}", "\n\n", t)
    return "" if t in ("-", "0") else t


def _load(data: bytes, data_only: bool):
    return openpyxl.load_workbook(io.BytesIO(data), data_only=data_only)


def find_lf_sheet(wb, marker="key indicators"):
    """Return (sheet, header_row) for the first sheet that has a 'Key Indicators' header."""
    for ws in wb.worksheets:
        for r in range(1, min(ws.max_row, 15) + 1):
            for c in range(1, min(ws.max_column, 40) + 1):
                if _norm(ws.cell(r, c).value) == marker:
                    return ws, r
    raise ValueError("Could not find a sheet with a 'Key Indicators' header.")


def find_code_col(ws, header_row) -> int:
    """Column holding indicator codes = the column with the most OC/OP-style values."""
    best, best_n = None, 0
    for c in range(1, min(ws.max_column, 10) + 1):
        n = sum(
            1 for r in range(header_row + 1, ws.max_row + 1)
            if CODE_RE.match(_norm_code(ws.cell(r, c).value))
        )
        if n > best_n:
            best, best_n = c, n
    if not best:
        raise ValueError("Could not find the indicator code column (OC 1.1, OP 2.1.3 ...).")
    return best


def headers(ws, header_row) -> dict[str, int]:
    out = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None and str(v).strip():
            out.setdefault(str(v).strip(), c)
    return out


def find_col(hdrs: dict[str, int], name: str):
    for h, c in hdrs.items():
        if _norm(h) == _norm(name):
            return c
    return None


# --------------------------------------------------------------------------- partner side
@dataclass
class PartnerEntry:
    value: object = None          # number or text ('-') as reported
    num: float | None = None      # numerator if value is a fraction formula
    den: float | None = None      # denominator
    reason: str = ""
    no_cached_value: bool = False  # formula present but file never recalculated in Excel


@dataclass
class PartnerData:
    label: str
    entries: dict[str, PartnerEntry] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


def partner_text_headers(data: bytes) -> list[str]:
    """All header names in the partner LF sheet (to choose achievement / reason columns)."""
    ws, hr = find_lf_sheet(_load(data, data_only=True))
    return list(headers(ws, hr).keys())


def _resolve(token: str, ws_vals):
    token = token.replace("$", "")
    if re.fullmatch(NUM, token):
        return float(token)
    v = ws_vals[token].value
    return float(v) if _is_number(v) else None


def read_partner(data: bytes, label: str, value_header: str,
                 reason_header: str = DEFAULT_REASON_HEADER) -> PartnerData:
    wb_f = _load(data, data_only=False)
    wb_v = _load(data, data_only=True)
    ws_f, hr = find_lf_sheet(wb_f)
    ws_v = wb_v[ws_f.title]

    hdrs = headers(ws_f, hr)
    val_col = find_col(hdrs, value_header)
    if val_col is None:
        raise ValueError(f"[{label}] column '{value_header}' not found in sheet '{ws_f.title}'.")
    rsn_col = find_col(hdrs, reason_header) if reason_header else None
    code_col = find_code_col(ws_f, hr)

    pdata = PartnerData(label=label)
    if reason_header and rsn_col is None:
        pdata.warnings.append(f"[{label}] no '{reason_header}' column - reasons skipped.")

    for r in range(hr + 1, ws_f.max_row + 1):
        code = _norm_code(ws_f.cell(r, code_col).value)
        if not CODE_RE.match(code):
            continue
        e = PartnerEntry()
        f = ws_f.cell(r, val_col).value
        v = ws_v.cell(r, val_col).value
        if isinstance(f, str) and f.startswith("="):
            m = FRACTION_RE.match(f.replace(" ", ""))
            if m:
                e.num, e.den = _resolve(m.group(1), ws_v), _resolve(m.group(2), ws_v)
                if e.num is None or e.den in (None, 0):
                    e.num = e.den = None
            if v is None:
                if e.num is not None:
                    v = e.num / e.den          # we can calculate a simple fraction ourselves
                else:
                    e.no_cached_value = True
                    pdata.warnings.append(
                        f"[{label}] {code}: formula {f} has no saved value - open the file in "
                        "Excel, save, and upload again.")
        e.value = v
        if rsn_col:
            e.reason = _clean_text(ws_v.cell(r, rsn_col).value)
        if code in pdata.entries:
            pdata.warnings.append(f"[{label}] {code} appears twice - first row used.")
            continue
        pdata.entries[code] = e
    return pdata


# --------------------------------------------------------------------------- master side
def _partner_cols_in_sheet(ws, hr) -> dict[str, int]:
    out = {}
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(hr, c).value or "").strip()
        m = re.search(r"achievement\s*[-:]?\s+(.+)$", h, re.IGNORECASE)
        if m and not re.search(r"total", h, re.IGNORECASE):
            out[re.sub(r"^partner", "Partner", m.group(1).strip(), flags=re.I)] = c
    return out


def master_partner_columns(data: bytes) -> dict[str, int]:
    """{'Partner A': col, ...} from headers like 'Q3 2026 Achievement Partner A'.
    Empty dict if the overall LF has no partner columns yet (they will be created)."""
    ws, hr = find_lf_sheet(_load(data, data_only=False))
    return _partner_cols_in_sheet(ws, hr)


def suggest_partner_label(filename: str, existing: list[str], index: int = 0) -> str:
    """Existing column if the filename matches one, else a new label from the filename
    (CHAIN_PartnerA_M&E Report... -> 'Partner A'), else 'Partner <n>'."""
    found = guess_partner_label(filename, existing)
    if found:
        return found
    m = re.search(r"partner[\s_-]*([A-Za-z0-9]+?)(?=[\s_\-.&]|$)", filename, re.IGNORECASE)
    if m:
        return f"Partner {m.group(1).upper() if len(m.group(1)) <= 2 else m.group(1)}"
    return f"Partner {index + 1}"


def guess_partner_label(filename: str, labels: list[str]):
    m = re.search(r"partner[\s_-]*([A-Za-z0-9]+?)(?=[\s_\-.&]|$)", filename, re.IGNORECASE)
    if m:
        want = f"partner {m.group(1)}".lower()
        for lb in labels:
            if lb.lower() == want:
                return lb
    for lb in labels:   # partner name (e.g. BMA) inside filename
        if lb.lower().replace("partner ", "") in filename.lower():
            return lb
    return None


def _is_pct_indicator(text) -> bool:
    t = _norm(text)
    return any(t.startswith(w) or w == "%" and "%" in t for w in PCT_WORDS)


def combine(master_bytes: bytes, partners: list[PartnerData],
            reason_col_name: str = DEFAULT_REASON_HEADER,
            clear_uploaded_columns: bool = True, pooled_percent: bool = True,
            period_label: str = "Q3"):
    """Returns (xlsx_bytes, report_rows, warnings).
    Partner / Total / Reason columns that do not exist in the overall LF are created,
    named e.g. '<period_label> Achievement Partner A'."""
    wb = _load(master_bytes, data_only=False)
    ws, hr = find_lf_sheet(wb)
    hdrs = headers(ws, hr)
    code_col = find_code_col(ws, hr)
    ind_col = find_col(hdrs, "Key Indicators")
    pcols = _partner_cols_in_sheet(ws, hr)
    total_col = next((c for h, c in hdrs.items() if re.search(r"total achievement", h, re.I)), None)
    warnings = [w for p in partners for w in p.warnings]
    period = (period_label or "").strip()

    def last_header_col():
        return max((c for c in range(1, ws.max_column + 1)
                    if str(ws.cell(hr, c).value or "").strip()), default=0)

    style_src = ws.cell(hr, last_header_col() or 1)

    def add_col(name, width=16):
        c = last_header_col() + 1
        cell = ws.cell(hr, c, name)
        cell.font, cell.fill, cell.border = copy(style_src.font), copy(style_src.fill), copy(style_src.border)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(c)].width = width
        return c

    def ensure_col(name, width=16):
        return find_col(headers(ws, hr), name) or add_col(name, width)

    # Partner columns that are missing -> create (in upload order)
    for p in partners:
        if p.label not in pcols:
            pcols[p.label] = add_col(f"{period} Achievement {p.label}".strip())
            warnings.append(f"Added new column '{period} Achievement {p.label}'.".replace("  ", " "))
    if total_col is None:
        total_col = add_col(f"{period} Total Achievement".strip())

    rsn_col = ensure_col(reason_col_name)
    chk_col = ensure_col("Total Check")
    ws.column_dimensions[get_column_letter(rsn_col)].width = 70
    ws.column_dimensions[get_column_letter(chk_col)].width = 40

    master_codes, report = set(), []
    pc_letters = [get_column_letter(c) for c in sorted(pcols.values())]

    for r in range(hr + 1, ws.max_row + 1):
        code = _norm_code(ws.cell(r, code_col).value)
        if not CODE_RE.match(code):
            continue
        master_codes.add(code)
        row = {"Row": r, "Code": code,
               "Indicator": str(ws.cell(r, ind_col).value or "")[:80] if ind_col else ""}

        # 1) partner values
        for p in partners:
            cell = ws.cell(r, pcols[p.label])
            if isinstance(cell, MergedCell):
                continue
            e = p.entries.get(code)
            if e is not None:
                cell.value = e.value
            elif clear_uploaded_columns:
                cell.value = None

        # 2) combined reasons (uploaded partners, in master column order)
        parts = [f"[{p.label}]: {p.entries[code].reason}"
                 for p in sorted(partners, key=lambda p: pcols[p.label])
                 if code in p.entries and p.entries[code].reason]
        rc = ws.cell(r, rsn_col)
        rc.value = "\n\n".join(parts) if parts else None
        rc.alignment = Alignment(wrap_text=True, vertical="top")

        # 3) total
        by_label = {p.label: p for p in partners}
        vals, fracs, missing = [], [], []
        for lb, c in pcols.items():
            v = ws.cell(r, c).value
            if not _is_number(v):
                continue
            vals.append(v)
            e = by_label[lb].entries.get(code) if lb in by_label else None
            if e and e.num is not None:
                fracs.append((e.num, e.den))
            else:
                missing.append(lb)
        is_pct = bool(fracs) or (ind_col and _is_pct_indicator(ws.cell(r, ind_col).value))
        check = ""
        if total_col and not isinstance(ws.cell(r, total_col), MergedCell):
            tc = ws.cell(r, total_col)
            if is_pct and pooled_percent:
                if fracs and not missing:
                    n, d = sum(f[0] for f in fracs), sum(f[1] for f in fracs)
                    tc.value = n / d
                    tc.number_format = "0.0%"
                    check = f"Pooled: {n:g}/{d:g} from {len(fracs)} partner(s)"
                elif vals:
                    tc.value = None
                    check = ("MANUAL: % without numerator/denominator from "
                             + ", ".join(missing))
                else:
                    tc.value = None
            elif not is_pct:
                tc.value = "=SUM(" + ",".join(f"{L}{r}" for L in pc_letters) + ")"
                check = "Sum" if vals else ""
            row["Total"] = (sum(f[0] for f in fracs) / sum(f[1] for f in fracs)
                            if is_pct and fracs and not missing else
                            (sum(vals) if vals and not is_pct else None))
        ws.cell(r, chk_col).value = check or None
        for lb, c in pcols.items():
            row[lb] = ws.cell(r, c).value
        row["Total Check"] = check
        row["Reasons"] = len(parts)
        report.append(row)

    for p in partners:
        extra = sorted(set(p.entries) - master_codes)
        if extra:
            warnings.append(f"[{p.label}] codes not found in overall LF: {', '.join(extra)}")

    wb.calculation.fullCalcOnLoad = True   # Excel recalculates SUM formulas on open
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), report, warnings


# =============================================================================
#                                   USER INTERFACE
# =============================================================================
import hashlib
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Combine Partner LFs", page_icon="🗃️", layout="wide")

st.markdown("""
<style>
  #MainMenu, footer, header [data-testid="stToolbar"] {visibility: hidden;}
  .block-container {padding-top: 2rem; max-width: 1200px;}
  .app-title {font-size: 1.9rem; font-weight: 700; color: #1f3b57; margin-bottom: .1rem;}
  .app-sub {color: #5b6b7b; margin-bottom: 1.4rem;}
  .step {font-size: 1.05rem; font-weight: 700; color: #1f3b57; margin: .2rem 0 .6rem;}
  .step span {display:inline-block; width:1.6rem; height:1.6rem; line-height:1.6rem; text-align:center;
              border-radius:50%; background:#1f6fb2; color:#fff; font-size:.85rem; margin-right:.5rem;}
  div[data-testid="stMetricValue"] {font-size: 1.8rem;}
</style>
""", unsafe_allow_html=True)


def step(n, title):
    st.markdown(f'<div class="step"><span>{n}</span>{title}</div>', unsafe_allow_html=True)


def fmt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float):
        return f"{v:.1%}" if 0 < v < 1 else f"{v:g}"
    return str(v)


st.markdown('<div class="app-title"> Combine Partner LFs</div>', unsafe_allow_html=True)
st.markdown('<div class="app-sub">Merge every partner\'s logframe report into the overall LF — '
            'matched by indicator code, with all deviation reasons in one column.</div>',
            unsafe_allow_html=True)

# ---------------------------------------------------------------- 1. upload
with st.container(border=True):
    step(1, "Upload files")
    c1, c2 = st.columns(2)
    overall_file = c1.file_uploader("Overall LF", type=["xlsx"])
    partner_files = c2.file_uploader("Partner LF reports (select all)", type=["xlsx"],
                                     accept_multiple_files=True)

if not (overall_file and partner_files):
    st.caption("Upload the overall LF and at least one partner LF to continue.")
    st.stop()

master_bytes = overall_file.getvalue()
files = {f.name: f.getvalue() for f in partner_files}
try:
    existing = list(master_partner_columns(master_bytes).keys())
    all_headers = []
    for name, data in files.items():
        for h in partner_text_headers(data):
            if h not in all_headers:
                all_headers.append(h)
except ValueError as e:
    st.error(str(e))
    st.stop()

# ---------------------------------------------------------------- 2. settings
with st.container(border=True):
    step(2, "Check settings")

    ach = [h for h in all_headers if "achievement" in h.lower()] or all_headers
    a1, a2, a3 = st.columns([1, 1.4, 1])
    value_header = a1.selectbox("Period to read", ach)
    reason_idx = next((i for i, h in enumerate(all_headers) if "reason" in h.lower()), 0)
    reason_header = a2.selectbox("Reason column", all_headers, index=reason_idx)
    period_label = a3.text_input("Name for new columns",
                                 value=f"{value_header.split()[0]} {datetime.now().year}",
                                 help="New columns are named '<this> Achievement <Partner>'.")

    st.markdown("**Partner name for each file** — taken from the filename, edit if needed")
    key = "names_" + hashlib.md5("|".join(files).encode()).hexdigest()[:8]
    names_df = pd.DataFrame({
        "File": list(files),
        "Partner": [suggest_partner_label(n, existing, i) for i, n in enumerate(files)],
    })
    edited = st.data_editor(
        names_df, key=key, hide_index=True, use_container_width=True,
        disabled=["File"],
        column_config={"File": st.column_config.TextColumn(width="large"),
                       "Partner": st.column_config.TextColumn(width="medium", required=True)},
    )
    assign = dict(zip(edited["File"], edited["Partner"].fillna("").astype(str).str.strip()))
    new_cols = [p for p in assign.values() if p and p not in existing]
    if existing:
        st.caption(f"Existing partner columns in overall LF: {', '.join(existing)}"
                   + (f" · will add: {', '.join(new_cols)}" if new_cols else ""))
    else:
        st.caption(f"No partner columns in overall LF yet — will add: {', '.join(new_cols)}")

    with st.expander("Options"):
        clear_cols = st.checkbox("Blank out indicators a partner did not report", value=True)
        pooled = st.checkbox("Pooled total for % indicators (Σ numerators ÷ Σ denominators)", value=True)

    names = list(assign.values())
    problem = ("Every file needs a partner name." if any(not n for n in names) else
               "Two files have the same partner name." if len(set(names)) < len(names) else None)
    if problem:
        st.error(problem)
    run = st.button("Combine", type="primary", use_container_width=True, disabled=bool(problem))

if run:
    try:
        partners = [read_partner(files[n], assign[n], value_header, reason_header) for n in files]
        st.session_state.result = combine(master_bytes, partners,
                                          clear_uploaded_columns=clear_cols,
                                          pooled_percent=pooled, period_label=period_label)
        st.session_state.result_label = period_label
    except ValueError as e:
        st.session_state.pop("result", None)
        st.error(str(e))

# ---------------------------------------------------------------- 3. results
if "result" in st.session_state:
    out, report, warns = st.session_state.result
    rep = pd.DataFrame(report)
    partner_cols = [c for c in rep.columns if c not in
                    ("Row", "Code", "Indicator", "Total", "Total Check", "Reasons")]
    manual = rep[rep["Total Check"].str.startswith("MANUAL", na=False)]
    added = [w for w in warns if w.startswith("Added new column")]
    issues = [w for w in warns if not w.startswith("Added new column")]

    with st.container(border=True):
        step(3, "Review & download")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Indicators", len(rep))
        m2.metric("Partners", len(partner_cols))
        m3.metric("With reasons", int((rep["Reasons"] > 0).sum()))
        m4.metric("Manual totals", len(manual))

        if issues:
            st.warning("**Please check:**\n\n" + "\n".join(f"- {w}" for w in issues))
        if added:
            st.info(" · ".join(added))

        show = rep.drop(columns=["Row"]).copy()
        for c in partner_cols + ["Total"]:
            show[c] = show[c].map(fmt)
        t1, t2 = st.tabs(["All indicators", f"Needs manual total ({len(manual)})"])
        t1.dataframe(show, hide_index=True, use_container_width=True, height=420)
        t2.dataframe(show[show["Total Check"].str.startswith("MANUAL", na=False)],
                     hide_index=True, use_container_width=True)

        label = "".join(ch for ch in st.session_state.result_label if ch.isalnum() or ch in " _-")
        st.download_button("📥  Download combined overall LF", data=out, type="primary",
                           use_container_width=True,
                           file_name=f"Overall_LF_{label.strip().replace(' ', '_')}_"
                                     f"{datetime.now():%Y%m%d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
