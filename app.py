import io
import json
import re
from collections import Counter

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Direct Databricks native connector
try:
    from databricks import sql as dbsql
    HAS_DATABRICKS_DRIVER = True
except ImportError:
    HAS_DATABRICKS_DRIVER = False

# Fuzzy matching engine setup
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

from sqlalchemy import create_engine, text


# ============================================================================
# Core Normalization & Matching Logic
# ============================================================================

def normalize(text_val: str) -> str:
    """Handle CamelCase, PascalCase, separators, and collapse whitespace."""
    text_val = str(text_val).strip()
    text_val = re.sub(r"([a-z])([A-Z])", r"\1 \2", text_val)
    text_val = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", text_val)
    text_val = text_val.lower()
    text_val = re.sub(r"[_\-]+", " ", text_val)
    text_val = re.sub(r"\s+", " ", text_val)
    return text_val.strip()


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

    # 1. Exact DB match
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
            if (r.get("db_column") or r["excel_column"]) in allowed
        ]

    report_df = pd.DataFrame([
        {
            "Excel Column (original)": r["excel_column"],
            "Mapped DB Column": r.get("db_column") if r.get("db_column") else "-- NOT MATCHED --",
            "Match Type": r.get("match_type", "Manual"),
            "Confidence Score": round(float(r.get("score", 0.0)), 1),
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


def load_uploaded_df(file) -> pd.DataFrame:
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    xls = pd.ExcelFile(file)
    return pd.read_excel(file, sheet_name=xls.sheet_names[0])


# ============================================================================
# Databricks Native Driver Helpers
# ============================================================================

def get_databricks_connection(server_hostname: str, http_path: str, access_token: str, catalog: str = "", schema: str = ""):
    clean_host = server_hostname.replace("https://", "").replace("http://", "").strip("/")
    clean_path = http_path.strip()
    if not clean_path.startswith("/"):
        clean_path = "/" + clean_path

    clean_token = access_token.strip().strip("'").strip('"')
    if clean_token.lower().startswith("bearer "):
        clean_token = clean_token[7:].strip()

    conn_kwargs = {
        "server_hostname": clean_host,
        "http_path": clean_path,
        "access_token": clean_token,
    }
    if catalog and catalog.strip():
        conn_kwargs["catalog"] = catalog.strip()
    if schema and schema.strip():
        conn_kwargs["schema"] = schema.strip()

    return dbsql.connect(**conn_kwargs)


def format_table_identifier(tbl: str, creds_dict: dict) -> str:
    """Safely quotes 3-part or 2-part Databricks table identifiers."""
    parts = [p.strip().replace("`", "") for p in tbl.strip().split(".") if p.strip()]
    if len(parts) == 3:
        return f"`{parts[0]}`.`{parts[1]}`.`{parts[2]}`"
    elif len(parts) == 2:
        cat = creds_dict.get("catalog", "").strip()
        if cat:
            return f"`{cat}`.`{parts[0]}`.`{parts[1]}`"
        return f"`{parts[0]}`.`{parts[1]}`"
    else:
        cat = creds_dict.get("catalog", "").strip()
        sch = creds_dict.get("schema", "").strip()
        if cat and sch:
            return f"`{cat}`.`{sch}`.`{parts[0]}`"
        elif sch:
            return f"`{sch}`.`{parts[0]}`"
        return f"`{parts[0]}`"


def init_mock_metadata_db():
    """Local SQLite mock database simulating a metadata table."""
    eng = create_engine("sqlite:///:memory:")
    with eng.connect() as conn:
        conn.execute(text("""
            CREATE TABLE columns_master (
                report_name TEXT,
                column_name TEXT,
                aliases TEXT
            );
        """))
        conn.execute(text("""
            INSERT INTO columns_master VALUES
            ('Sales Summary', 'customer_id', 'cust_id, client_code, id'),
            ('Sales Summary', 'order_date', 'date, txn_date, purchase_date'),
            ('Sales Summary', 'total_amount', 'sales, revenue, net_amount'),
            ('Sales Summary', 'sales_rep', 'agent, representative, employee'),
            ('Inventory Ledger', 'sku_code', 'item_code, product_id, part_number'),
            ('Inventory Ledger', 'warehouse_location', 'wh_id, site, facility'),
            ('Inventory Ledger', 'stock_quantity', 'qty, available_qty, count');
        """))
        conn.commit()
    return eng


# ============================================================================
# Streamlit Interface
# ============================================================================

st.set_page_config(
    page_title="AutoSchema Mapper - Databricks Edition",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
        .main { background-color: #f8fafc; }
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
        }
        .hero-desc {
            color: #94a3b8;
            font-size: 1.05rem;
            margin: 0;
            max-width: 850px;
        }
        .metric-card {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 1.25rem 1.5rem;
            box-shadow: 0 2px 4px rgba(0,0,0,0.02);
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
        .step-header {
            font-size: 1.25rem;
            font-weight: 600;
            color: #1e293b;
            margin-bottom: 1rem;
        }
        .col-pill {
            display: inline-block;
            background-color: #e2e8f0;
            color: #0f172a;
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 0.85rem;
            font-family: monospace;
            margin: 3px;
        }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="hero-container">
        <div class="hero-title">⚡ AutoSchema Mapper (Databricks)</div>
        <p class="hero-desc">Extract target schema columns directly from your Databricks metadata registry table by selecting your table, filter column, and report name.</p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Engine Settings")
    threshold = st.slider(
        "Fuzzy Match Sensitivity",
        min_value=50, max_value=100, value=75, step=5,
        format="%d%%",
        help="Higher values require closer spelling matches before auto-accepting."
    )
    st.markdown("---")
    st.markdown("### ℹ️ Engine Info")
    st.caption(f"**Fuzzy Backend:** `{FUZZ_BACKEND}`")
    st.caption(f"**Databricks Driver:** `{'Installed' if HAS_DATABRICKS_DRIVER else 'Missing'}`")

# --- 1. Target Schema Configuration ---
with st.container():
    st.markdown('<div class="step-header">🎯 1. Target Schema Definition (Databricks Metadata)</div>', unsafe_allow_html=True)

    mode = st.radio(
        "Connection Type:",
        ["Connect to Databricks SQL Warehouse", "Built-in Demo (Metadata Table Simulation)"],
        horizontal=True
    )

    db_columns_input = []

    if mode == "Built-in Demo (Metadata Table Simulation)":
        if "mock_meta_eng" not in st.session_state:
            st.session_state.mock_meta_eng = init_mock_metadata_db()
        eng = st.session_state.mock_meta_eng

        with eng.connect() as conn:
            demo_reports = [r[0] for r in conn.execute(text("SELECT DISTINCT report_name FROM columns_master ORDER BY 1")).fetchall()]

        selected_report = st.selectbox("Select Report Name from columns_master:", demo_reports)
        if selected_report:
            with eng.connect() as conn:
                res = conn.execute(text(f"SELECT column_name, aliases FROM columns_master WHERE report_name = '{selected_report}'")).fetchall()
                db_columns_input = [
                    {
                        "db_column": row[0],
                        "aliases": [a.strip() for a in str(row[1] or "").split(",") if a.strip()]
                    }
                    for row in res
                ]
            st.markdown(f"**Loaded {len(db_columns_input)} Columns for `{selected_report}`:**")
            st.markdown(" ".join([f"<span class='col-pill'>{c['db_column']}</span>" for c in db_columns_input]), unsafe_allow_html=True)

    else:
        if not HAS_DATABRICKS_DRIVER:
            st.error("Missing dependency: Run `pip install databricks-sql-connector` in your terminal.")
            st.stop()

        # Safely pull secrets from Streamlit Cloud or local .streamlit/secrets.toml
        db_secrets = st.secrets.get("databricks", {})
        default_host = db_secrets.get("host", "")
        default_http_path = db_secrets.get("http_path", "")
        default_token = db_secrets.get("token", "")
        default_catalog = db_secrets.get("catalog", "workspace")
        default_schema = db_secrets.get("schema", "excel_column_mapping_utility")

        has_secrets = bool(default_host and default_http_path and default_token)

        with st.expander("🔑 Databricks Warehouse Credentials", expanded=not has_secrets and "db_creds" not in st.session_state):
            if has_secrets:
                st.caption("🔒 Credentials detected from Streamlit Secrets.")

            col_c1, col_c2 = st.columns(2)
            with col_c1:
                db_host = st.text_input("Server Hostname", value=default_host, placeholder="adb-xxxx.xx.azuredatabricks.net")
                db_http_path = st.text_input("HTTP Path", value=default_http_path, placeholder="/sql/1.0/warehouses/xxxxxxxxxxxx")
            with col_c2:
                db_token = st.text_input("Personal Access Token (PAT)", value=default_token, type="password")
                col_cat, col_sch = st.columns(2)
                with col_cat:
                    db_catalog = st.text_input("Catalog", value=default_catalog)
                with col_sch:
                    db_schema = st.text_input("Schema / Database", value=default_schema)

            connect_btn = st.button("🔗 Connect to Databricks")

        should_connect = connect_btn or (has_secrets and "db_creds" not in st.session_state)

        if should_connect and db_host and db_http_path and db_token:
            try:
                with st.spinner("Connecting and loading tables from Databricks..."):
                    conn = get_databricks_connection(db_host, db_http_path, db_token, db_catalog, db_schema)
                    with conn.cursor() as cursor:
                        cursor.execute("SHOW TABLES")
                        res = cursor.fetchall()
                        found_tables = [row["tableName"] if isinstance(row, dict) else row[1] for row in res]
                        st.session_state["databricks_tables"] = sorted(found_tables)
                        st.session_state["db_creds"] = {
                            "host": db_host, "path": db_http_path, "token": db_token,
                            "catalog": db_catalog, "schema": db_schema
                        }
                    conn.close()
                st.success(f"✓ Connected successfully! Found {len(found_tables)} table(s) in `{db_catalog}.{db_schema}`.")
            except Exception as e:
                st.error(f"Databricks Connection Failed: {e}")

        if "db_creds" in st.session_state:
            creds = st.session_state["db_creds"]

            config_method = st.radio(
                "How do you want to extract columns from your metadata table?",
                ["Filter by Report Name (Cascading Dropdowns)", "Write Custom SQL Query"],
                horizontal=True
            )

            if config_method == "Filter by Report Name (Cascading Dropdowns)":
                tables_available = st.session_state.get("databricks_tables", [])

                # 1️⃣ Table Selection
                c_tbl, c_refresh = st.columns([4, 1])
                with c_tbl:
                    if tables_available:
                        default_idx = 0
                        for idx, t_name in enumerate(tables_available):
                            if "columns_master" in t_name.lower():
                                default_idx = idx
                                break
                        selected_table = st.selectbox("1️⃣ Select Metadata Table:", options=tables_available, index=default_idx)
                    else:
                        selected_table = st.text_input("1️⃣ Metadata Table Name:", value="columns_master")

                with c_refresh:
                    st.write("")
                    st.write("")
                    if st.button("🔄 Reload Columns"):
                        st.session_state.pop(f"cols_{selected_table}", None)
                        st.session_state.pop(f"filter_vals_{selected_table}", None)

                # Fetch table column definitions
                table_cols_key = f"cols_{selected_table}"
                if table_cols_key not in st.session_state and selected_table:
                    try:
                        conn = get_databricks_connection(creds["host"], creds["path"], creds["token"], creds["catalog"], creds["schema"])
                        qual_tbl = format_table_identifier(selected_table, creds)
                        with conn.cursor() as cursor:
                            cursor.execute(f"DESCRIBE TABLE {qual_tbl}")
                            desc_res = cursor.fetchall()
                            loaded_fields = [
                                (row["col_name"] if isinstance(row, dict) else row[0])
                                for row in desc_res
                                if row and not str(row[0]).startswith("#")
                            ]
                            st.session_state[table_cols_key] = loaded_fields
                        conn.close()
                    except Exception as e:
                        st.error(f"Could not read columns from `{selected_table}`: {e}")

                table_fields = st.session_state.get(table_cols_key, [])

                if table_fields:
                    col_step2, col_step3 = st.columns(2)

                    # 2️⃣ Filter Column
                    def find_best_index(options, candidates):
                        for c in candidates:
                            for idx, opt in enumerate(options):
                                if c in opt.lower():
                                    return idx
                        return 0

                    with col_step2:
                        filter_col_idx = find_best_index(table_fields, ["report", "type", "category"])
                        filter_column = st.selectbox("2️⃣ Filter Column (Report Name Field):", options=table_fields, index=filter_col_idx)

                    # Dynamic distinct values for the chosen filter column
                    filter_cache_key = f"filter_vals_{selected_table}_{filter_column}"
                    if filter_cache_key not in st.session_state:
                        try:
                            conn = get_databricks_connection(creds["host"], creds["path"], creds["token"], creds["catalog"], creds["schema"])
                            qual_tbl = format_table_identifier(selected_table, creds)
                            with conn.cursor() as cursor:
                                cursor.execute(f"SELECT DISTINCT `{filter_column}` FROM {qual_tbl} WHERE `{filter_column}` IS NOT NULL ORDER BY 1")
                                distinct_rows = cursor.fetchall()
                                st.session_state[filter_cache_key] = [str(r[0]) for r in distinct_rows if r[0] is not None]
                            conn.close()
                        except Exception as e:
                            st.error(f"Could not fetch distinct values for `{filter_column}`: {e}")

                    distinct_filter_values = st.session_state.get(filter_cache_key, [])

                    # 3️⃣ Filter Value Selection
                    with col_step3:
                        if distinct_filter_values:
                            chosen_report = st.selectbox(
                                f"3️⃣ Select Value to Filter on (`{filter_column}`):",
                                options=distinct_filter_values
                            )
                        else:
                            chosen_report = None
                            st.warning(f"No values found in `{filter_column}`.")

                    # Query all other columns automatically for the selected report
                    if chosen_report:
                        try:
                            conn = get_databricks_connection(creds["host"], creds["path"], creds["token"], creds["catalog"], creds["schema"])
                            qual_tbl = format_table_identifier(selected_table, creds)

                            remaining_cols = [c for c in table_fields if c != filter_column]
                            target_candidates = [c for c in remaining_cols if any(k in c.lower() for k in ["column", "target", "col", "field"])]
                            target_col = target_candidates[0] if target_candidates else (remaining_cols[0] if remaining_cols else None)

                            alias_candidates = [c for c in remaining_cols if c != target_col and any(k in c.lower() for k in ["alias", "synonym"])]
                            alias_col = alias_candidates[0] if alias_candidates else None

                            with conn.cursor() as cursor:
                                if target_col and alias_col:
                                    query = f"SELECT `{target_col}`, `{alias_col}` FROM {qual_tbl} WHERE `{filter_column}` = '{chosen_report}'"
                                    cursor.execute(query)
                                    results = cursor.fetchall()
                                    db_columns_input = [
                                        {
                                            "db_column": str(r[0]).strip(),
                                            "aliases": [a.strip() for a in str(r[1] or "").split(",") if a.strip()]
                                        }
                                        for r in results if r[0]
                                    ]
                                elif target_col:
                                    query = f"SELECT `{target_col}` FROM {qual_tbl} WHERE `{filter_column}` = '{chosen_report}'"
                                    cursor.execute(query)
                                    results = cursor.fetchall()
                                    db_columns_input = [{"db_column": str(r[0]).strip(), "aliases": []} for r in results if r[0]]
                                else:
                                    query = f"SELECT * FROM {qual_tbl} WHERE `{filter_column}` = '{chosen_report}'"
                                    cursor.execute(query)
                                    results = cursor.fetchall()
                                    col_names = [col[0] for col in cursor.description]
                                    t_idx = 1 if len(col_names) > 1 else 0
                                    db_columns_input = [{"db_column": str(r[t_idx]).strip(), "aliases": []} for r in results if r[t_idx]]

                            conn.close()

                            st.markdown(f"**Loaded {len(db_columns_input)} Target Columns for `{chosen_report}`:**")
                            st.markdown(" ".join([f"<span class='col-pill'>{c['db_column']}</span>" for c in db_columns_input]), unsafe_allow_html=True)
                        except Exception as e:
                            st.error(f"Error loading report columns: {e}")

            else:
                default_query = f"SELECT column_name, aliases FROM {creds.get('catalog', 'workspace')}.{creds.get('schema', 'excel_column_mapping_utility')}.columns_master WHERE report_name = 'Sales_Report'"
                custom_sql = st.text_area(
                    "✏️ Custom SQL Query:",
                    value=st.session_state.get("last_custom_meta_sql", default_query),
                    help="First column must return target column names. Second column (optional) can return comma-separated aliases.",
                    height=100
                )
                if st.button("🚀 Execute SQL Query"):
                    st.session_state["last_custom_meta_sql"] = custom_sql
                    try:
                        with st.spinner("Executing query..."):
                            conn = get_databricks_connection(creds["host"], creds["path"], creds["token"], creds["catalog"], creds["schema"])
                            with conn.cursor() as cursor:
                                cursor.execute(custom_sql)
                                rows = cursor.fetchall()
                                if len(cursor.description) > 1:
                                    db_columns_input = [
                                        {
                                            "db_column": str(r[0]).strip(),
                                            "aliases": [a.strip() for a in str(r[1] or "").split(",") if a.strip()]
                                        }
                                        for r in rows if r[0]
                                    ]
                                else:
                                    db_columns_input = [{"db_column": str(r[0]).strip(), "aliases": []} for r in rows if r[0]]
                                st.session_state["custom_sql_cols"] = db_columns_input
                            conn.close()
                            st.success(f"✓ Retrieved {len(db_columns_input)} column names.")
                    except Exception as e:
                        st.error(f"SQL execution error: {e}")

                if "custom_sql_cols" in st.session_state and not db_columns_input:
                    db_columns_input = st.session_state["custom_sql_cols"]
                    st.markdown(f"**Active Target Columns ({len(db_columns_input)}):**")
                    st.markdown(" ".join([f"<span class='col-pill'>{c['db_column']}</span>" for c in db_columns_input]), unsafe_allow_html=True)

st.markdown("<div style='margin-top: 1.5rem;'></div>", unsafe_allow_html=True)

# --- 2. File Uploads (Source & Up to 5 Masters) ---
with st.container():
    col_src, col_mst = st.columns(2)
    with col_src:
        st.markdown('<div class="step-header">📂 2A. Source Spreadsheet (Required)</div>', unsafe_allow_html=True)
        excel_file = st.file_uploader("Upload incoming spreadsheet", type=["xlsx", "xls", "csv"], key="source_file")
    with col_mst:
        st.markdown('<div class="step-header">📚 2B. Master Datasets (Max 5, Optional)</div>', unsafe_allow_html=True)
        master_files = st.file_uploader(
            "Upload up to 5 master tables to join",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="master_files"
        )
        if master_files and len(master_files) > 5:
            st.warning("⚠️ Only the first 5 uploaded master files will be processed.")
            master_files = master_files[:5]

# --- 3. Column Matching & Review ---
if excel_file is not None and db_columns_input:
    targets = build_targets(db_columns_input)
    available_db_columns = [t["db_column"] for t in targets]
    dropdown_options = ["-- NOT MATCHED --"] + available_db_columns

    try:
        raw_df = load_uploaded_df(excel_file)
    except Exception as e:
        st.error(f"Could not read source spreadsheet: {e}")
        st.stop()

    file_key = f"mapping_data_{excel_file.name}_{threshold}_{len(db_columns_input)}"
    if file_key not in st.session_state:
        base_results = map_columns(list(raw_df.columns), targets, threshold)
        st.session_state[file_key] = pd.DataFrame([
            {
                "Excel Column (Original)": r["excel_column"],
                "Target DB Column": r["db_column"] if r["db_column"] else "-- NOT MATCHED --",
                "Match Status": _badge(r["match_type"]),
                "Confidence": f"{round(r['score'])}%",
            }
            for r in base_results
        ])

    st.markdown("<div style='margin-top: 1.5rem;'></div>", unsafe_allow_html=True)
    st.markdown('<div class="step-header">🔍 3. Review & Validate Mappings</div>', unsafe_allow_html=True)

    edited_table = st.data_editor(
        st.session_state[file_key],
        use_container_width=True,
        hide_index=True,
        disabled=["Excel Column (Original)", "Match Status", "Confidence"],
        column_config={
            "Target DB Column": st.column_config.SelectboxColumn(
                "Target DB Column",
                options=dropdown_options,
                required=True,
                help="Select target column to link or override.",
            )
        },
        key=f"editor_{file_key}"
    )

    st.session_state[file_key] = edited_table

    reconciled_results = []
    assigned_targets = []
    for _, row in edited_table.iterrows():
        orig_col = row["Excel Column (Original)"]
        target = None if row["Target DB Column"] == "-- NOT MATCHED --" else row["Target DB Column"]
        is_manual = "Manual" in row["Match Status"] or row["Target DB Column"] != "-- NOT MATCHED --"

        reconciled_results.append({
            "excel_column": orig_col,
            "db_column": target,
            "match_type": "Manual" if is_manual else row["Match Status"],
            "score": 100.0 if target else 0.0,
        })
        if target:
            assigned_targets.append(target)

    # Real-time KPIs
    total_cols = len(reconciled_results)
    matched_count = sum(1 for r in reconciled_results if r["db_column"])
    unmatched_count = total_cols - matched_count
    match_rate = round((matched_count / total_cols * 100), 1) if total_cols else 0

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Source Columns</div><div class="metric-num">{total_cols}</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Mapped</div><div class="metric-num" style="color:#16a34a;">{matched_count}</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Unmapped</div><div class="metric-num" style="color:#dc2626;">{unmatched_count}</div></div>', unsafe_allow_html=True)
    with m4:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Coverage</div><div class="metric-num" style="color:#2563eb;">{match_rate}%</div></div>', unsafe_allow_html=True)

    duplicates = [col for col, count in Counter(assigned_targets).items() if count > 1]
    has_collision = len(duplicates) > 0
    if has_collision:
        st.error(f"⛔ **Mapping Collision:** Target column `{', '.join(duplicates)}` is mapped to multiple original columns.")

    rename_dict = {r["excel_column"]: r["db_column"] for r in reconciled_results if r["db_column"]}
    working_df = raw_df.rename(columns=rename_dict)

    # --- 4. Multi-Master Data Joiner ---
    st.markdown("<div style='margin-top: 2rem;'></div>", unsafe_allow_html=True)
    st.markdown('<div class="step-header">🔗 4. Master Data Joins (Up to 5 Masters)</div>', unsafe_allow_html=True)

    joined_df = working_df.copy()

    if master_files:
        st.caption(f"Loaded **{len(master_files)}** master dataset(s). Configure join parameters below:")
        tabs = st.tabs([f"Master {i+1}: {f.name}" for i, f in enumerate(master_files)])

        for i, (tab, mfile) in enumerate(zip(tabs, master_files)):
            with tab:
                try:
                    m_df = load_uploaded_df(mfile)
                    st.write(f"**Rows:** {m_df.shape[0]} | **Columns:** {m_df.shape[1]}")

                    c_enable, c_join = st.columns([1, 2])
                    with c_enable:
                        enable_this_join = st.checkbox(f"Enable Join for Master {i+1}", value=True, key=f"enable_m_{i}")
                    with c_join:
                        join_type = st.selectbox(
                            "Join Type",
                            ["left", "right", "inner", "outer"],
                            index=0,
                            key=f"join_type_{i}",
                            format_func=lambda x: f"{x.capitalize()} Join"
                        )

                    c_left, c_right = st.columns(2)
                    with c_left:
                        current_source_cols = list(joined_df.columns)
                        left_key = st.selectbox(
                            "Source Key Column",
                            options=current_source_cols,
                            key=f"left_key_{i}"
                        )
                    with c_right:
                        right_key = st.selectbox(
                            "Master Key Column",
                            options=list(m_df.columns),
                            key=f"right_key_{i}"
                        )

                    available_enrichment = [c for c in m_df.columns if c != right_key]
                    selected_enrichment = st.multiselect(
                        "Columns to bring from this Master:",
                        options=available_enrichment,
                        default=available_enrichment,
                        key=f"cols_m_{i}"
                    )

                    if enable_this_join and left_key and right_key:
                        master_subset = m_df[[right_key] + selected_enrichment].drop_duplicates(subset=[right_key])
                        suffix = f"_m{i+1}"
                        joined_df = pd.merge(
                            joined_df,
                            master_subset,
                            how=join_type,
                            left_on=left_key,
                            right_on=right_key,
                            suffixes=("", suffix)
                        )
                except Exception as e:
                    st.error(f"Error executing join for {mfile.name}: {e}")

        st.success(f"✓ Join pipeline complete. Resulting dimensions: {joined_df.shape[0]} rows × {joined_df.shape[1]} columns.")
    else:
        st.info("💡 Upload one or more master files in **Step 2B** to enrich your data.")

    # --- 5. Export ---
    st.markdown("<div style='margin-top: 2rem;'></div>", unsafe_allow_html=True)
    st.markdown('<div class="step-header">🚀 5. Preview & Export</div>', unsafe_allow_html=True)

    all_exportable_cols = list(joined_df.columns)
    with st.expander("👁️ Data Preview & Column Selection", expanded=True):
        selected_cols = st.multiselect(
            "Included Columns in Output:",
            options=all_exportable_cols,
            default=all_exportable_cols,
        )
        if selected_cols:
            st.dataframe(joined_df[selected_cols].head(10), use_container_width=True)
        else:
            st.warning("Select at least one column to export.")

    filter_download = st.checkbox("Export only the selected column subset", value=False)
    final_export_df = joined_df[selected_cols] if filter_download and selected_cols else joined_df
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
    st.info("💡 Select a report or execute custom SQL in **Step 1** to populate target columns.")
else:
    st.info("💡 Upload an incoming spreadsheet in **Step 2A** to begin matching.")
