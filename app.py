#!/usr/bin/env python3
"""
Streamlit app: Excel Column Mapper (Modern UI Edition)
"""

import io
import json
import re
from collections import Counter

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
        def token_sort_ratio(a: str, b: str) -> float:
            return SequenceMatcher(None, a, b).ratio() * 100

    fuzz = _FuzzShim()


# ============================================================================
# Core Logic
# ============================================================================

def normalize(text: str) -> str:
    """Handle CamelCase, PascalCase, separators, and collapse whitespace."""
    text = str(text).strip()
    text = re.sub(r"([a-z])([A-Z])", r"\1 \2", text)
    text = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", text)
    text = text.lower()
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def build_targets(db_columns: list[dict]) -> list[dict]:
    targets = []
    for entry in db_columns:
        db_column = str(entry.get("db_column", "")).strip()
        if not db_column:
            continue
        aliases = [str(a).strip() for a in entry.get("aliases", []) if str(a).strip()]
        targets.append({
            "db_column": db_column,
            "aliases": aliases,
            "norm_db": normalize(db_column),
            "norm_aliases": [normalize(a) for a in aliases],
        })
    return targets


def match_column(excel_col: str, targets: list[dict], threshold: float) -> dict:
    norm_col = normalize(excel_col)

    # 1. Exact DB name match
    for t in targets:
        if norm_col == t["norm_db"]:
            return {"db_column": t["db_column"], "match_type": "Exact", "score": 100.0}

    # 2. Exact Alias match
    for t in targets:
        if norm_col in t["norm_aliases"]:
            return {"db_column": t["db_column"], "match_type": "Alias", "score": 100.0}

    # 3. Fuzzy match
    best = {"db_column": None, "match_type": "Unmatched", "score": 0.0}
    for t in targets:
        for cand in [t["norm_db"]] + t["norm_aliases"]:
            score = fuzz.token_sort_ratio(norm_col, cand)
            if score > best["score"]:
                best = {"db_column": t["db_column"], "match_type": "Fuzzy", "score": float(score)}

    if best["score"] >= threshold:
        return best
    return {"db_column": None, "match_type": "Unmatched", "score": best["score"]}


def map_columns(excel_columns: list[str], targets: list[dict], threshold: float) -> list[dict]:
    raw_results = [
        {"excel_column": col, **match_column(col, targets, threshold)}
        for col in excel_columns
    ]

    # Resolve automatic conflicts (greedy on highest score)
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
            final_results.append({
                "excel_column": r["excel_column"],
                "db_column": None,
                "match_type": "Unmatched (duplicate)",
                "score": r["score"],
            })
        else:
            final_results.append(r)
    return final_results


def build_output_workbook_bytes(
    mapped_df: pd.DataFrame,
    mapping_results: list[dict],
    restrict_report_to: list[str] | None = None,
) -> bytes:
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

    return _style_report_sheet(buffer.read(), len(report_df))


def _style_report_sheet(workbook_bytes: bytes, n_rows: int) -> bytes:
    wb = load_workbook(io.BytesIO(workbook_bytes))
    ws = wb["Column Mapping Report"]

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    unmatched_fill = PatternFill(start_color="FFE4E6", end_color="FFE4E6", fill_type="solid")
    fuzzy_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    exact_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for row_idx in range(2, n_rows + 2):
        match_type = ws.cell(row=row_idx, column=3).value
        if match_type and ("Unmatched" in str(match_type) or "Duplicate" in str(match_type)):
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
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 50)
    ws.freeze_panes = "A2"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _badge(match_type: str) -> str:
    if not match_type or "Unmatched" in match_type:
        return f"❌ {match_type or 'Unmatched'}"
    if match_type == "Manual":
        return "✏️ Manual Override"
    if match_type == "Fuzzy":
        return "⚡ Fuzzy Match"
    if match_type == "Alias":
        return "🏷️ Known Alias"
    return "✅ Exact Match"


# ============================================================================
# Streamlit Interface & Custom Styling
# ============================================================================

st.set_page_config(
    page_title="AutoSchema Mapper",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Modern SaaS Styling
st.markdown("""
    <style>
        /* Main background & fonts */
        .main {
            background-color: #f8fafc;
        }
        
        /* Hero Banner */
        .hero-container {
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
            padding: 2.2rem 2.5rem;
            border-radius: 16px;
            color: #ffffff;
            margin-bottom: 2rem;
            box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.15);
        }
        .hero-title {
            font-size: 2.2rem;
            font-weight: 700;
            letter-spacing: -0.02em;
            margin-bottom: 0.4rem;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .hero-desc {
            color: #94a3b8;
            font-size: 1.05rem;
            margin: 0;
            max-width: 700px;
        }
        
        /* Metric cards */
        .metric-card {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 1.25rem 1.5rem;
            box-shadow: 0 2px 4px rgba(0,0,0,0.02);
            transition: all 0.2s ease;
        }
        .metric-card:hover {
            border-color: #cbd5e1;
            box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        }
        .metric-num {
            font-size: 1.85rem;
            font-weight: 700;
            color: #0f172a;
            line-height: 1.2;
        }
        .metric-label {
            font-size: 0.85rem;
            font-weight: 600;
            color: #64748b;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }
        
        /* Section cards */
        .step-header {
            font-size: 1.25rem;
            font-weight: 600;
            color: #1e293b;
            margin-bottom: 1rem;
            display: flex;
            align-items: center;
            gap: 8px;
        }
    </style>
""", unsafe_allow_html=True)

# Hero Header
st.markdown("""
    <div class="hero-container">
        <div class="hero-title">⚡ AutoSchema Mapper</div>
        <p class="hero-desc">Instantly normalize raw Excel spreadsheets into your target database schema with intelligent fuzzy matching, alias detection, and one-click export.</p>
    </div>
""", unsafe_allow_html=True)

SAMPLE_SCHEMA = [
    {"DB Column Name": "customer_id", "Known Aliases (comma-separated, optional)": "cust_id, client_num, id, account_no"},
    {"DB Column Name": "customer_name", "Known Aliases (comma-separated, optional)": "CustomerName, client name, full_name, user_name"},
    {"DB Column Name": "order_date", "Known Aliases (comma-separated, optional)": "OrderDate, dt, purchase_date, txn_date"},
    {"DB Column Name": "total_amount", "Known Aliases (comma-separated, optional)": "Total, total_usd, revenue, sales, net_amount"},
]

if "db_columns_df" not in st.session_state:
    st.session_state.db_columns_df = pd.DataFrame(SAMPLE_SCHEMA)

# --- Sidebar Configuration ---
with st.sidebar:
    st.markdown("### ⚙️ Engine Settings")
    threshold = st.slider(
        "Fuzzy Match Sensitivity",
        min_value=50, max_value=100, value=75, step=5,
        format="%d%%",
        help="Higher values require closer spelling matches before accepting an automatic match."
    )
    
    st.markdown("---")
    st.markdown("### ℹ️ Engine Info")
    st.caption(f"**Backend:** `{FUZZ_BACKEND}`")
    st.caption("**Status:** Engine Ready")

# --- 1. Target Schema Configuration ---
with st.container():
    st.markdown('<div class="step-header">🎯 1. Define Target DB Schema</div>', unsafe_allow_html=True)
    
    col_mode, col_reset = st.columns([3, 1])
    with col_mode:
        col_input_mode = st.segmented_control(
            "Schema Input Mode",
            ["Interactive Table", "Upload JSON", "Upload CSV"],
            default="Interactive Table",
            label_visibility="collapsed"
        )
    with col_reset:
        if st.button("🔄 Reset Template", use_container_width=True):
            st.session_state.db_columns_df = pd.DataFrame(SAMPLE_SCHEMA)
            st.rerun()

    db_columns_input = []
    if col_input_mode == "Interactive Table":
        edited_df = st.data_editor(
            st.session_state.db_columns_df,
            num_rows="dynamic",
            use_container_width=True,
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
    elif col_input_mode == "Upload JSON":
        json_file = st.file_uploader("Upload schema JSON", type=["json"], label_visibility="collapsed")
        if json_file:
            try:
                db_columns_input = json.load(json_file)
                st.success(f"✓ Loaded {len(db_columns_input)} target columns successfully.")
            except Exception as e:
                st.error(f"JSON Parsing Error: {e}")
    else:
        csv_file = st.file_uploader("Upload schema CSV with 'db_column' column", type=["csv"], label_visibility="collapsed")
        if csv_file:
            try:
                csv_df = pd.read_csv(csv_file)
                for _, row in csv_df.iterrows():
                    aliases_raw = str(row.get("aliases", "") or "")
                    aliases = [a.strip() for a in re.split(r"[|,]", aliases_raw) if a.strip()]
                    db_columns_input.append({"db_column": str(row["db_column"]).strip(), "aliases": aliases})
                st.success(f"✓ Loaded {len(db_columns_input)} target columns successfully.")
            except Exception as e:
                st.error(f"CSV Parsing Error: {e}")

st.markdown("<div style='margin-top: 2rem;'></div>", unsafe_allow_html=True)

# --- 2. File Upload ---
with st.container():
    st.markdown('<div class="step-header">📂 2. Upload Source Spreadsheet</div>', unsafe_allow_html=True)
    excel_file = st.file_uploader("Upload your incoming spreadsheet", type=["xlsx", "xls"], label_visibility="collapsed")

# --- 3. Processing, Visual Metrics & Mapping ---
if excel_file is not None and db_columns_input:
    targets = build_targets(db_columns_input)
    available_db_columns = [t["db_column"] for t in targets]
    dropdown_options = ["-- NOT MATCHED --"] + available_db_columns

    try:
        xls = pd.ExcelFile(excel_file)
        if len(xls.sheet_names) > 1:
            sheet = st.selectbox("📑 Select Active Sheet", xls.sheet_names)
        else:
            sheet = xls.sheet_names[0]
        raw_df = pd.read_excel(excel_file, sheet_name=sheet)
    except Exception as e:
        st.error(f"Could not read spreadsheet: {e}")
        st.stop()

    base_results = map_columns(list(raw_df.columns), targets, threshold)

    st.markdown("<div style='margin-top: 2rem;'></div>", unsafe_allow_html=True)
    st.markdown('<div class="step-header">🔍 3. Review & Validate Mappings</div>', unsafe_allow_html=True)

    editor_data = pd.DataFrame([
        {
            "Excel Column (Original)": r["excel_column"],
            "Target DB Column": r["db_column"] if r["db_column"] else "-- NOT MATCHED --",
            "Match Status": _badge(r["match_type"]),
            "Confidence": f"{round(r['score'])}%",
        }
        for r in base_results
    ])

    edited_table = st.data_editor(
        editor_data,
        use_container_width=True,
        hide_index=True,
        disabled=["Excel Column (Original)", "Match Status", "Confidence"],
        column_config={
            "Target DB Column": st.column_config.SelectboxColumn(
                "Target DB Column",
                options=dropdown_options,
                required=True,
                help="Select target column to link or override manually.",
            )
        },
    )

    # Reconcile overrides
    reconciled_results = []
    assigned_targets = []
    for orig, edited_val in zip(base_results, edited_table["Target DB Column"]):
        target = None if edited_val == "-- NOT MATCHED --" else edited_val
        is_modified = target != orig["db_column"]

        reconciled_results.append({
            "excel_column": orig["excel_column"],
            "db_column": target,
            "match_type": "Manual" if is_modified and target else (orig["match_type"] if not is_modified else "Unmatched (manual)"),
            "score": 100.0 if is_modified and target else (orig["score"] if not is_modified else 0.0),
        })
        if target:
            assigned_targets.append(target)

    # Calculate real-time stats
    total_cols = len(reconciled_results)
    matched_count = sum(1 for r in reconciled_results if r["db_column"])
    unmatched_count = total_cols - matched_count
    match_rate = round((matched_count / total_cols * 100), 1) if total_cols else 0

    # Display KPI Metrics
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Source Columns</div><div class="metric-num">{total_cols}</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Mapped</div><div class="metric-num" style="color:#16a34a;">{matched_count}</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Unmapped</div><div class="metric-num" style="color:#dc2626;">{unmatched_count}</div></div>', unsafe_allow_html=True)
    with m4:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Coverage</div><div class="metric-num" style="color:#2563eb;">{match_rate}%</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='margin-top: 1.5rem;'></div>", unsafe_allow_html=True)

    # Duplicate collisions
    duplicates = [col for col, count in Counter(assigned_targets).items() if count > 1]
    has_collision = len(duplicates) > 0

    if has_collision:
        st.error(f"⛔ **Mapping Collision:** Target column `{', '.join(duplicates)}` is assigned multiple times. Please adjust assignments to proceed.")

    # --- 4. Export Section ---
    st.markdown('<div class="step-header">🚀 4. Preview & Export</div>', unsafe_allow_html=True)
    
    rename_dict = {r["excel_column"]: r["db_column"] for r in reconciled_results if r["db_column"]}
    preview_df = raw_df.rename(columns=rename_dict)
    all_exportable_cols = list(preview_df.columns)

    with st.expander("👁️ Data Preview & Column Filtering", expanded=True):
        selected_cols = st.multiselect(
            "Included Columns in Output:",
            options=all_exportable_cols,
            default=all_exportable_cols,
        )
        if selected_cols:
            st.dataframe(preview_df[selected_cols].head(8), use_container_width=True)
        else:
            st.warning("Select at least one column to export.")

    filter_download = st.checkbox("Export only the selected column subset", value=False)
    final_export_df = preview_df[selected_cols] if filter_download and selected_cols else preview_df
    restrict_report = selected_cols if filter_download and selected_cols else None

    if not has_collision and len(final_export_df.columns) > 0:
        output_bytes = build_output_workbook_bytes(final_export_df, reconciled_results, restrict_report_to=restrict_report)
        out_name = f"{excel_file.name.rsplit('.', 1)[0]}_standardized.xlsx"

        st.download_button(
            label="📥 Download Standardized Spreadsheet",
            data=output_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    else:
        st.button("📥 Download Standardized Spreadsheet", disabled=True, use_container_width=True)

elif not db_columns_input:
    st.info("💡 Add at least one target DB column in **Step 1** to get started.")
else:
    st.info("💡 Upload an Excel spreadsheet in **Step 2** to begin matching.")
