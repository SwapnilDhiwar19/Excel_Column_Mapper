#!/usr/bin/env python3
"""
Streamlit app: Excel Column Mapper

Lets a user:
  1. Enter/edit their database column names (and optional known aliases) in a
     table right in the browser.
  2. Upload a messy Excel file.
  3. See a live mapping report (Exact / Alias / Fuzzy / Unmatched, with
     confidence scores) and MANUALLY override the "Mapped DB Column" for any
     row right in the table — useful when no automatic match was found, or
     the automatic match picked the wrong column.
  4. Pick exactly which final columns to include in the preview (and
     optionally the downloaded file) via a multiselect.
  5. Download a final .xlsx with two sheets:
       - "Mapped Data"           -> original rows, headers renamed to DB names
       - "Column Mapping Report" -> the audit trail (incl. manual overrides)

Run with:
    streamlit run app.py
"""

import io
import json
import re

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    from rapidfuzz import fuzz
    FUZZ_BACKEND = "rapidfuzz"
except ImportError:
    from difflib import SequenceMatcher
    FUZZ_BACKEND = "difflib"

    class _FuzzShim:
        @staticmethod
        def token_sort_ratio(a, b):
            return SequenceMatcher(None, a, b).ratio() * 100

    fuzz = _FuzzShim()


# ============================================================================
# Core matching logic (pure functions, no Streamlit calls -> easy to test)
# ============================================================================

def normalize(text: str) -> str:
    """Lowercase, strip, collapse whitespace/underscores/hyphens so
    'Customer_Name', 'Customer Name', 'CustomerName' all compare fairly."""
    text = str(text).strip().lower()
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def build_targets(db_columns: list[dict]) -> list[dict]:
    """db_columns: list of {'db_column': str, 'aliases': list[str]}"""
    targets = []
    for entry in db_columns:
        db_column = entry["db_column"].strip()
        if not db_column:
            continue
        aliases = [a.strip() for a in entry.get("aliases", []) if a.strip()]
        targets.append({
            "db_column": db_column,
            "aliases": aliases,
            "norm_db": normalize(db_column),
            "norm_aliases": [normalize(a) for a in aliases],
        })
    return targets


def match_column(excel_col: str, targets: list[dict], threshold: float) -> dict:
    """Priority: exact db_column match > exact alias match > best fuzzy match."""
    norm_col = normalize(excel_col)

    for t in targets:
        if norm_col == t["norm_db"]:
            return {"db_column": t["db_column"], "match_type": "Exact", "score": 100.0}

    for t in targets:
        if norm_col in t["norm_aliases"]:
            return {"db_column": t["db_column"], "match_type": "Alias", "score": 100.0}

    best = {"db_column": None, "match_type": "Unmatched", "score": 0.0}
    for t in targets:
        for cand in [t["norm_db"]] + t["norm_aliases"]:
            score = fuzz.token_sort_ratio(norm_col, cand)
            if score > best["score"]:
                best = {"db_column": t["db_column"], "match_type": "Fuzzy", "score": score}

    if best["score"] >= threshold:
        return best
    return {"db_column": None, "match_type": "Unmatched", "score": best["score"]}


def map_columns(excel_columns: list[str], targets: list[dict], threshold: float) -> list[dict]:
    """Map every excel column, then resolve conflicts if two excel columns
    both want the same DB column (keep the higher-confidence one)."""
    raw_results = [
        {"excel_column": col, **match_column(col, targets, threshold)}
        for col in excel_columns
    ]

    claimed = {}
    for r in raw_results:
        if r["db_column"] is None:
            continue
        key = r["db_column"]
        if key not in claimed or r["score"] > claimed[key]["score"]:
            claimed[key] = r
    winners = {id(r) for r in claimed.values()}

    final_results = []
    for r in raw_results:
        if r["db_column"] is not None and id(r) not in winners:
            final_results.append({**r, "db_column": None,
                                   "match_type": "Unmatched (duplicate claim)"})
        else:
            final_results.append(r)
    return final_results


def build_output_workbook_bytes(
    mapped_df: pd.DataFrame,
    mapping_results: list[dict],
    restrict_report_to: list[str] | None = None,
) -> bytes:
    """Builds the two-sheet output workbook in memory and returns raw bytes.

    `mapped_df` should already have its headers renamed to DB column names
    (and already filtered down to whichever columns should ship, if the user
    chose to restrict the download to a subset).

    `restrict_report_to`, if given, limits the "Column Mapping Report" sheet
    to rows whose final column name (db_column or original excel_column) is
    in that list, so the report matches what's actually in "Mapped Data".
    """
    report_rows = mapping_results
    if restrict_report_to is not None:
        allowed = set(restrict_report_to)
        report_rows = [
            r for r in mapping_results
            if (r["db_column"] or r["excel_column"]) in allowed
        ]

    report_df = pd.DataFrame([
        {
            "Excel Column (original)": r["excel_column"],
            "Mapped DB Column": r["db_column"] if r["db_column"] else "-- NOT MATCHED --",
            "Match Type": r["match_type"],
            "Confidence Score": round(r["score"], 1),
        }
        for r in report_rows
    ])

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        mapped_df.to_excel(writer, sheet_name="Mapped Data", index=False)
        report_df.to_excel(writer, sheet_name="Column Mapping Report", index=False)
    buffer.seek(0)

    styled_bytes = _style_report_sheet(buffer.read(), len(report_df))
    return styled_bytes


def _style_report_sheet(workbook_bytes: bytes, n_rows: int) -> bytes:
    wb = load_workbook(io.BytesIO(workbook_bytes))
    ws = wb["Column Mapping Report"]

    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    unmatched_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fuzzy_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    exact_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for row_idx in range(2, n_rows + 2):
        match_type = ws.cell(row=row_idx, column=3).value
        if match_type and "Unmatched" in str(match_type):
            fill = unmatched_fill
        elif match_type == "Fuzzy":
            fill = fuzzy_fill
        else:
            fill = exact_fill
        for cell in ws[row_idx]:
            cell.fill = fill

    for col_cells in ws.columns:
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 50)
    ws.freeze_panes = "A2"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _badge(match_type: str) -> str:
    """Small emoji prefix so match quality is visible even in an editable
    (unstyled) data_editor table."""
    if match_type is None:
        return "Unmatched"
    if match_type == "Manual":
        return "\U0001F535 Manual"
    if "Unmatched" in match_type:
        return f"\U0001F534 {match_type}"
    if match_type == "Fuzzy":
        return f"\U0001F7E1 {match_type}"
    return f"\U0001F7E2 {match_type}"  # Exact / Alias


# ============================================================================
# Streamlit UI
# ============================================================================

st.set_page_config(
    page_title="Excel Column Mapper",
    page_icon="\U0001F517",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---- Global styling ----------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"]  {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* Hide default Streamlit chrome for a cleaner look */
    #MainMenu, footer {visibility: hidden;}

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        max-width: 1100px;
    }

    /* Hero banner */
    .ecm-hero {
        background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 50%, #4338CA 100%);
        border-radius: 18px;
        padding: 2.1rem 2.4rem;
        margin-bottom: 1.8rem;
        box-shadow: 0 8px 24px rgba(79, 70, 229, 0.25);
    }
    .ecm-hero h1 {
        color: #FFFFFF;
        font-size: 1.85rem;
        font-weight: 800;
        margin: 0 0 0.35rem 0;
        letter-spacing: -0.02em;
    }
    .ecm-hero p {
        color: #E0E7FF;
        font-size: 1rem;
        margin: 0;
        max-width: 640px;
    }

    /* Step tracker */
    .ecm-steps {
        display: flex;
        gap: 0.6rem;
        margin-bottom: 1.6rem;
    }
    .ecm-step {
        flex: 1;
        display: flex;
        align-items: center;
        gap: 0.55rem;
        background: #F8FAFC;
        border: 1.5px solid #E2E8F0;
        border-radius: 12px;
        padding: 0.55rem 0.85rem;
        transition: all 0.2s ease;
    }
    .ecm-step.active {
        background: #EEF2FF;
        border-color: #4F46E5;
        box-shadow: 0 2px 8px rgba(79, 70, 229, 0.15);
    }
    .ecm-step.done {
        background: #F0FDF4;
        border-color: #86EFAC;
    }
    .ecm-step-num {
        width: 26px; height: 26px;
        border-radius: 50%;
        background: #CBD5E1;
        color: white;
        font-size: 0.8rem;
        font-weight: 700;
        display: flex; align-items: center; justify-content: center;
        flex-shrink: 0;
    }
    .ecm-step.active .ecm-step-num { background: #4F46E5; }
    .ecm-step.done .ecm-step-num { background: #22C55E; }
    .ecm-step-label {
        font-size: 0.82rem;
        font-weight: 600;
        color: #64748B;
    }
    .ecm-step.active .ecm-step-label { color: #4338CA; }
    .ecm-step.done .ecm-step-label { color: #15803D; }

    /* Section card headers */
    .ecm-section-title {
        display: flex;
        align-items: center;
        gap: 0.6rem;
        font-size: 1.15rem;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 0.2rem;
    }
    .ecm-section-badge {
        width: 30px; height: 30px;
        border-radius: 9px;
        background: linear-gradient(135deg, #4F46E5, #7C3AED);
        color: white;
        display: flex; align-items: center; justify-content: center;
        font-size: 1rem;
        flex-shrink: 0;
    }
    .ecm-section-sub {
        color: #64748B;
        font-size: 0.9rem;
        margin: 0.1rem 0 0.9rem 2.6rem;
    }

    /* Legend chips */
    .ecm-legend-chip {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 999px;
        padding: 0.25rem 0.7rem;
        font-size: 0.8rem;
        font-weight: 600;
        color: #475569;
        margin-right: 0.4rem;
    }

    /* Buttons */
    div.stButton > button, div.stDownloadButton > button {
        border-radius: 10px;
        font-weight: 600;
        padding: 0.55rem 1.4rem;
        border: none;
        transition: all 0.15s ease;
    }
    div.stButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] {
        background: linear-gradient(135deg, #4F46E5, #7C3AED);
        box-shadow: 0 4px 12px rgba(79, 70, 229, 0.3);
    }
    div.stButton > button[kind="primary"]:hover, div.stDownloadButton > button[kind="primary"]:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(79, 70, 229, 0.4);
    }

    /* Metric cards */
    div[data-testid="stMetric"] {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 0.9rem 1rem 0.6rem 1rem;
    }

    hr { margin: 1.6rem 0; }
    </style>
    """,
    unsafe_allow_html=True,
)


def section_title(number: str, icon: str, title: str, subtitle: str = "") -> None:
    st.markdown(
        f"""
        <div class="ecm-section-title">
            <div class="ecm-section-badge">{icon}</div>
            {number}. {title}
        </div>
        {f'<div class="ecm-section-sub">{subtitle}</div>' if subtitle else ''}
        """,
        unsafe_allow_html=True,
    )


with st.sidebar:
    st.markdown("### \U0001F517 Quick guide")
    st.markdown(
        """
        1. **Define columns** \u2014 list your real database column names, with
           optional aliases for variants you've already seen.
        2. **Upload** your messy Excel file.
        3. **Adjust the threshold** if you want fuzzy matches to be stricter
           or looser.
        4. **Run mapping**, then review the report \u2014 fix anything wrong or
           missing right in the table.
        5. **Pick columns to preview**, then download the finished file.
        """
    )
    st.divider()
    st.markdown("### How matching works")
    st.markdown(
        """
        - \U0001F7E2 **Exact** \u2014 header matches a DB column name exactly
        - \U0001F7E2 **Alias** \u2014 header matches a known alias exactly
        - \U0001F7E1 **Fuzzy** \u2014 best similarity score above threshold
        - \U0001F534 **Unmatched** \u2014 nothing cleared the threshold
        - \U0001F535 **Manual** \u2014 you picked it yourself
        """
    )
    st.divider()
    st.caption("Runs entirely in memory \u2014 nothing is stored on a server.")


# ---- Hero ---------------------------------------------------------------------
st.markdown(
    """
    <div class="ecm-hero">
        <h1>\U0001F517 Excel Column Mapper</h1>
        <p>Map messy Excel headers &mdash; "Customer Name", "CustomerName", "Cust_Name" &mdash;
        to your standardized database columns automatically, review and fix the mapping by hand,
        then export a clean, audited file.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---- Step tracker ---------------------------------------------------------------
if "db_columns_df" in st.session_state:
    _db_ready = any(
        str(r.get("DB Column Name", "")).strip()
        for _, r in st.session_state["db_columns_df"].iterrows()
    )
else:
    _db_ready = False
_file_ready = "mm_pending_file_name" in st.session_state
_mapped_ready = "mm_results" in st.session_state

_steps = [
    ("1", "Define columns", _db_ready),
    ("2", "Upload file", _file_ready),
    ("3", "Match settings", _file_ready),
    ("4", "Review & download", _mapped_ready),
]
_step_html = '<div class="ecm-steps">'
for i, (num, label, done) in enumerate(_steps):
    cls = "done" if done else ("active" if (i == 0 or _steps[i - 1][2]) and not done else "")
    marker = "\u2713" if done else num
    _step_html += f'<div class="ecm-step {cls}"><div class="ecm-step-num">{marker}</div><div class="ecm-step-label">{label}</div></div>'
_step_html += "</div>"
st.markdown(_step_html, unsafe_allow_html=True)

if "db_columns_df" not in st.session_state:
    st.session_state.db_columns_df = pd.DataFrame(
        [
            {"DB Column Name": "Customer_Name", "Known Aliases (comma-separated, optional)": "Customer Name, CustomerName, Cust Name"},
            {"DB Column Name": "Order_Date", "Known Aliases (comma-separated, optional)": "OrderDate, Order Dt"},
            {"DB Column Name": "Total_Amount", "Known Aliases (comma-separated, optional)": ""},
        ]
    )

# ---- Step 1: database column names -----------------------------------------
with st.container(border=True):
    section_title("1", "\U0001F4CB", "Define your database columns",
                  "Add one row per column. Aliases are optional but boost confidence on variants you already know about.")

    col_input_mode = st.radio(
        "How do you want to provide the column list?",
        ["Enter manually", "Upload JSON config", "Upload CSV (db_column, aliases)"],
        horizontal=True,
        label_visibility="collapsed",
    )

    if col_input_mode == "Enter manually":
        edited_df = st.data_editor(
            st.session_state.db_columns_df,
            num_rows="dynamic",
            width='stretch',
            key="db_columns_editor",
        )
        st.session_state.db_columns_df = edited_df
        db_columns_input = [
            {
                "db_column": str(row["DB Column Name"]).strip(),
                "aliases": [a.strip() for a in str(row.get("Known Aliases (comma-separated, optional)", "") or "").split(",") if a.strip()],
            }
            for _, row in edited_df.iterrows()
            if str(row["DB Column Name"]).strip()
        ]

    elif col_input_mode == "Upload JSON config":
        json_file = st.file_uploader(
            "Upload a JSON file: "
            '[{"db_column": "Customer_Name", "aliases": ["Customer Name", "CustomerName"]}, ...]',
            type=["json"],
        )
        db_columns_input = []
        if json_file is not None:
            try:
                db_columns_input = json.load(json_file)
                st.success(f"Loaded {len(db_columns_input)} target column(s) from JSON.")
            except Exception as e:
                st.error(f"Couldn't parse JSON: {e}")

    else:  # Upload CSV
        st.caption("CSV must have columns `db_column` and `aliases` (aliases pipe- or comma-separated).")
        csv_file = st.file_uploader("Upload CSV", type=["csv"])
        db_columns_input = []
        if csv_file is not None:
            try:
                csv_df = pd.read_csv(csv_file)
                for _, row in csv_df.iterrows():
                    aliases_raw = str(row.get("aliases", "") or "")
                    aliases = [a.strip() for a in re.split(r"[|,]", aliases_raw) if a.strip()]
                    db_columns_input.append({"db_column": str(row["db_column"]).strip(), "aliases": aliases})
                st.success(f"Loaded {len(db_columns_input)} target column(s) from CSV.")
            except Exception as e:
                st.error(f"Couldn't parse CSV: {e}")

st.write("")

# ---- Step 2: upload excel ----------------------------------------------------
with st.container(border=True):
    section_title("2", "\U0001F4C4", "Upload your Excel file")
    excel_file = st.file_uploader("Excel file (.xlsx)", type=["xlsx", "xls"], label_visibility="collapsed")

    sheet_name = 0
    if excel_file is not None:
        st.session_state["mm_pending_file_name"] = excel_file.name
        try:
            xls = pd.ExcelFile(excel_file)
            if len(xls.sheet_names) > 1:
                sheet_name = st.selectbox("Which sheet has your data?", xls.sheet_names)
            else:
                sheet_name = xls.sheet_names[0]
                st.caption(f"Using sheet: **{sheet_name}**")
        except Exception as e:
            st.error(f"Couldn't read the Excel file: {e}")
            excel_file = None
    elif "mm_pending_file_name" in st.session_state:
        del st.session_state["mm_pending_file_name"]

st.write("")

# ---- Step 3: options ----------------------------------------------------------
with st.container(border=True):
    section_title("3", "\U0001F3AF", "Match settings")
    threshold = st.slider(
        "Minimum fuzzy match confidence to auto-accept (%)",
        min_value=50, max_value=100, value=75, step=5,
        help="Below this score, a column is left unmatched rather than guessed.",
    )

st.write("")

# ---- Step 4: run + results -----------------------------------------------------
with st.container(border=True):
    section_title("4", "\U0001F680", "Run mapping")

    run = st.button("\U0001F680  Map Columns", type="primary", disabled=(excel_file is None or not db_columns_input),
                     width='content')

    if excel_file is None:
        st.info("Upload an Excel file to continue.")
    elif not db_columns_input:
        st.info("Add at least one database column above to continue.")

# Run the automatic matcher and stash results in session_state. Everything
# below reads from session_state (not from `run`) so that editing the table
# or the column-preview picker on a later rerun doesn't wipe out the results.
if run:
    excel_file.seek(0)
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    targets = build_targets(db_columns_input)

    if not targets:
        st.error("No valid database columns were provided.")
    else:
        mapping_results = map_columns(list(df.columns), targets, threshold)
        st.session_state["mm_df"] = df
        st.session_state["mm_results"] = mapping_results
        st.session_state["mm_available_db_columns"] = [t["db_column"] for t in targets]
        st.session_state["mm_source_name"] = excel_file.name

if "mm_results" in st.session_state:
    df = st.session_state["mm_df"]
    mapping_results = st.session_state["mm_results"]
    available_db_columns = st.session_state["mm_available_db_columns"]
    dropdown_options = available_db_columns + ["-- NOT MATCHED --"]

    matched = sum(1 for r in mapping_results if r["db_column"] is not None)
    total = len(mapping_results)

    st.write("")
    section_title("5", "\U0001F4CA", "Results")

    c1, c2, c3 = st.columns(3)
    c1.metric("Columns found in file", total)
    c2.metric("Successfully mapped", matched)
    c3.metric("Unmatched", total - matched)

    st.markdown("<div style='height:0.9rem'></div>", unsafe_allow_html=True)
    st.markdown("**Column Mapping Report**")
    st.markdown(
        """
        <span class="ecm-legend-chip">\U0001F7E2 Exact / Alias</span>
        <span class="ecm-legend-chip">\U0001F7E1 Fuzzy</span>
        <span class="ecm-legend-chip">\U0001F534 Unmatched</span>
        <span class="ecm-legend-chip">\U0001F535 Manual override</span>
        """,
        unsafe_allow_html=True,
    )
    st.caption("Edit **\u2018Mapped DB Column\u2019** directly on any row to fix an unmatched or wrong match.")
    st.write("")

    report_df = pd.DataFrame([
        {
            "Excel Column (original)": r["excel_column"],
            "Mapped DB Column": r["db_column"] if r["db_column"] else "-- NOT MATCHED --",
            "Match Quality": _badge(r["match_type"]),
            "Confidence Score": round(r["score"], 1),
        }
        for r in mapping_results
    ])

    edited_report = st.data_editor(
        report_df,
        width='stretch',
        hide_index=True,
        key="mm_report_editor",
        disabled=["Excel Column (original)", "Match Quality", "Confidence Score"],
        column_config={
            "Mapped DB Column": st.column_config.SelectboxColumn(
                "Mapped DB Column",
                options=dropdown_options,
                required=True,
                help="Change this if the automatic match is wrong or missing.",
            ),
        },
    )

    # Reconcile edits back into final mapping results: anything the user
    # changed from what the matcher originally picked becomes a "Manual" entry.
    final_results = []
    for orig, edited_val in zip(mapping_results, edited_report["Mapped DB Column"]):
        new_db = None if edited_val == "-- NOT MATCHED --" else edited_val
        if new_db != orig["db_column"]:
            final_results.append({
                "excel_column": orig["excel_column"],
                "db_column": new_db,
                "match_type": "Manual" if new_db else "Unmatched (manual)",
                "score": 100.0 if new_db else 0.0,
            })
        else:
            final_results.append(orig)

    final_matched = sum(1 for r in final_results if r["db_column"] is not None)
    if final_matched != matched:
        st.info(f"After manual edits: {final_matched}/{total} columns mapped.")

    # ---- Column selection for preview (and optionally the download) --------
    st.write("")
    section_title("6", "\U0001F441\uFE0F", "Preview & download",
                  "Choose which mapped columns to preview. Optionally apply the same filter to your download.")
    final_names = [r["db_column"] if r["db_column"] else r["excel_column"] for r in final_results]
    selected_columns = st.multiselect(
        "Only these columns will show in the preview below",
        options=final_names,
        default=final_names,
        label_visibility="collapsed",
    )
    apply_to_download = st.checkbox(
        "Only include the selected columns above in the downloaded file too",
        value=False,
    )

    rename_map = {r["excel_column"]: r["db_column"] for r in final_results if r["db_column"] is not None}
    full_mapped_df = df.rename(columns=rename_map)

    st.markdown("**Preview: Mapped Data** (first 10 rows)")
    preview_cols = [c for c in selected_columns if c in full_mapped_df.columns]
    if preview_cols:
        st.dataframe(full_mapped_df[preview_cols].head(10), width='stretch')
    else:
        st.info("No columns selected to preview.")

    download_df = full_mapped_df[preview_cols] if (apply_to_download and preview_cols) else full_mapped_df
    restrict_to = preview_cols if (apply_to_download and preview_cols) else None

    output_bytes = build_output_workbook_bytes(download_df, final_results, restrict_report_to=restrict_to)
    src_name = st.session_state.get("mm_source_name", "output.xlsx")
    out_name = (src_name.rsplit(".", 1)[0] if "." in src_name else src_name) + "_mapped.xlsx"

    st.write("")
    st.download_button(
        "\U0001F4E5  Download mapped Excel file",
        data=output_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    if total - final_matched > 0:
        st.warning(
            f"{total - final_matched} column(s) are still unmatched. "
            "Pick a value in the 'Mapped DB Column' dropdown above to map them manually."
        )

st.write("")
st.divider()
st.caption(f"\U0001F529 Matching engine: {FUZZ_BACKEND}")
