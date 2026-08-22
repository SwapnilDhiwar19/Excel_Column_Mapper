#!/usr/bin/env python3
"""
Streamlit app: Excel Column Mapper (Refactored)
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
    # Insert space before capital letters preceded by lowercase (CamelCase -> Camel Case)
    text = re.sub(r"([a-z])([A-Z])", r"\1 \2", text)
    # Insert space before consecutive capitals followed by lowercase (e.g., 'OrderID' -> 'Order ID')
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
                "match_type": "Unmatched (duplicate claim)",
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
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)
    ws.freeze_panes = "A2"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _badge(match_type: str) -> str:
    if not match_type or "Unmatched" in match_type:
        return f"🔴 {match_type or 'Unmatched'}"
    if match_type == "Manual":
        return "🔵 Manual"
    if match_type == "Fuzzy":
        return f"🟡 {match_type}"
    return f"🟢 {match_type}"


# ============================================================================
# Streamlit Interface
# ============================================================================

st.set_page_config(page_title="Excel Column Mapper", page_icon="🔗", layout="wide")

st.title("🔗 Excel Column Mapper")
st.caption("Standardize messy incoming Excel headers to target database schemas.")

SAMPLE_SCHEMA = [
    {"DB Column Name": "customer_id", "Known Aliases (comma-separated, optional)": "cust_id, client_num, id"},
    {"DB Column Name": "customer_name", "Known Aliases (comma-separated, optional)": "CustomerName, client name, full_name"},
    {"DB Column Name": "order_date", "Known Aliases (comma-separated, optional)": "OrderDate, dt, purchase_date"},
    {"DB Column Name": "total_amount", "Known Aliases (comma-separated, optional)": "Total, total_usd, revenue, sales"},
]

if "db_columns_df" not in st.session_state:
    st.session_state.db_columns_df = pd.DataFrame(SAMPLE_SCHEMA)

# --- 1. Target Schema Configuration ---
st.header("1. Target DB Schema")

c_mode, c_preset = st.columns([3, 1])
with c_mode:
    col_input_mode = st.radio(
        "Schema Input Mode",
        ["Manual Entry", "Upload JSON", "Upload CSV"],
        horizontal=True,
        label_visibility="collapsed",
    )
with c_preset:
    if st.button("🔄 Reset to Default Schema"):
        st.session_state.db_columns_df = pd.DataFrame(SAMPLE_SCHEMA)
        st.rerun()

db_columns_input = []
if col_input_mode == "Manual Entry":
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
    json_file = st.file_uploader("Upload schema JSON", type=["json"])
    if json_file:
        try:
            db_columns_input = json.load(json_file)
            st.success(f"Loaded {len(db_columns_input)} target columns.")
        except Exception as e:
            st.error(f"JSON Error: {e}")
else:
    csv_file = st.file_uploader("Upload schema CSV (must have 'db_column' column)", type=["csv"])
    if csv_file:
        try:
            csv_df = pd.read_csv(csv_file)
            for _, row in csv_df.iterrows():
                aliases_raw = str(row.get("aliases", "") or "")
                aliases = [a.strip() for a in re.split(r"[|,]", aliases_raw) if a.strip()]
                db_columns_input.append({"db_column": str(row["db_column"]).strip(), "aliases": aliases})
            st.success(f"Loaded {len(db_columns_input)} target columns.")
        except Exception as e:
            st.error(f"CSV Error: {e}")

# --- 2. File Upload & Settings ---
st.header("2. Upload & Settings")
col_upload, col_settings = st.columns([2, 1])

with col_upload:
    excel_file = st.file_uploader("Excel file (.xlsx, .xls)", type=["xlsx", "xls"])

with col_settings:
    threshold = st.slider(
        "Fuzzy Match Sensitivity (%)",
        min_value=50, max_value=100, value=75, step=5,
        help="Confidence cutoff to accept fuzzy matches.",
    )

# --- 3. Processing & Mapping ---
if excel_file is not None and db_columns_input:
    targets = build_targets(db_columns_input)
    available_db_columns = [t["db_column"] for t in targets]
    dropdown_options = ["-- NOT MATCHED --"] + available_db_columns

    try:
        xls = pd.ExcelFile(excel_file)
        sheet = xls.sheet_names[0] if len(xls.sheet_names) == 1 else st.selectbox("Select Sheet", xls.sheet_names)
        raw_df = pd.read_excel(excel_file, sheet_name=sheet)
    except Exception as e:
        st.error(f"Could not read Excel file: {e}")
        st.stop()

    st.header("3. Review & Override Mappings")
    base_results = map_columns(list(raw_df.columns), targets, threshold)

    editor_data = pd.DataFrame([
        {
            "Excel Column (original)": r["excel_column"],
            "Mapped DB Column": r["db_column"] if r["db_column"] else "-- NOT MATCHED --",
            "Auto Match Status": _badge(r["match_type"]),
            "Auto Score": round(r["score"], 1),
        }
        for r in base_results
    ])

    edited_table = st.data_editor(
        editor_data,
        use_container_width=True,
        hide_index=True,
        disabled=["Excel Column (original)", "Auto Match Status", "Auto Score"],
        column_config={
            "Mapped DB Column": st.column_config.SelectboxColumn(
                "Mapped DB Column",
                options=dropdown_options,
                required=True,
                help="Override the assigned target column.",
            )
        },
    )

    # Reconcile overrides & track manual modifications
    reconciled_results = []
    assigned_targets = []
    for orig, edited_val in zip(base_results, edited_table["Mapped DB Column"]):
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

    # Validate target duplicates
    duplicates = [col for col, count in Counter(assigned_targets).items() if count > 1]
    has_collision = len(duplicates) > 0

    if has_collision:
        st.error(f"⛔ **Duplicate Mapping Conflict:** The database column(s) `{', '.join(duplicates)}` have been assigned to multiple Excel headers. Fix them before downloading.")

    # --- 4. Preview & Export ---
    st.header("4. Preview & Download")
    rename_dict = {r["excel_column"]: r["db_column"] for r in reconciled_results if r["db_column"]}
    
    # Construct preview DF (keep original unmapped name if unassigned)
    preview_df = raw_df.rename(columns=rename_dict)
    
    all_exportable_cols = list(preview_df.columns)
    selected_cols = st.multiselect(
        "Select output columns:",
        options=all_exportable_cols,
        default=all_exportable_cols,
    )

    if selected_cols:
        st.dataframe(preview_df[selected_cols].head(10), use_container_width=True)
    else:
        st.info("No columns selected for preview.")

    filter_download = st.checkbox("Apply column subset selection to final file", value=False)
    
    final_export_df = preview_df[selected_cols] if filter_download and selected_cols else preview_df
    restrict_report = selected_cols if filter_download and selected_cols else None

    if not has_collision:
        output_bytes = build_output_workbook_bytes(final_export_df, reconciled_results, restrict_report_to=restrict_report)
        out_name = f"{excel_file.name.rsplit('.', 1)[0]}_standardized.xlsx"

        st.download_button(
            label="📥 Download Processed Excel",
            data=output_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    else:
        st.button("📥 Download Processed Excel", disabled=True, help="Resolve duplicate column mappings above to enable download.")

elif not db_columns_input:
    st.info("Define at least one target database column in Step 1.")
else:
    st.info("Upload an Excel file in Step 2 to begin mapping.")

st.divider()
st.caption(f"Fuzzy Engine: `{FUZZ_BACKEND}` | Streamlit Native")
