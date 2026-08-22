# Excel Column Mapper

There are two ways to use this: a **Streamlit web app** (`app.py`, interactive,
no config file needed) or the **CLI script** (`excel_column_mapper.py`,
config-file driven, good for automation/scheduled jobs).

## Streamlit app (`app.py`)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Keep `.streamlit/config.toml` in the same folder as `app.py` — it sets the
app's color theme (indigo/violet). If you skip it, the app still works, just
with Streamlit's default theme instead of the custom one.

This opens a browser UI where you can:
1. Enter your database column names + optional known aliases directly in an
   editable table (or upload a JSON/CSV config instead)
2. Upload your Excel file
3. Adjust the fuzzy-match confidence threshold with a slider
4. See the full mapping report on screen (🟢 Exact/Alias, 🟡 Fuzzy,
   🔴 Unmatched, 🔵 Manual) and **edit the "Mapped DB Column" dropdown on any
   row** — e.g. to fix a column the automatic matcher missed, or to correct a
   wrong fuzzy guess. This updates instantly, no re-upload needed.
5. **Pick exactly which columns to view** in the preview via a multiselect —
   handy when the sheet has many columns and you only care about a few. An
   optional checkbox also restricts the downloaded file to that same subset.
6. Download the final two-sheet `.xlsx` (Mapped Data + Column Mapping Report)

Nothing is written to disk on the server — everything happens in memory and
the result comes back as a direct download.

---


Maps messy Excel column headers (`Customer Name`, `CustomerName`, `Cust_Name` ...)
to your standardized database column names (`Customer_Name`), then writes a new
Excel file with:

1. **Mapped Data** — your original rows, headers renamed to the DB names
2. **Column Mapping Report** — an audit sheet showing every original column,
   what it was mapped to, how (Exact / Alias / Fuzzy), and a confidence score
   (color-coded: green = confident, yellow = fuzzy match, red = unmatched)

## 1. Install dependencies

```bash
pip install pandas openpyxl rapidfuzz --break-system-packages
```
(`rapidfuzz` is optional but recommended — the script falls back to Python's
built-in `difflib` if it's missing, just with slightly less accurate scoring.)

## 2. Edit `target_columns.json` with YOUR database schema

```json
[
  { "db_column": "Customer_Name", "aliases": ["Customer Name", "CustomerName", "Cust_Name"] },
  { "db_column": "Order_Date",    "aliases": ["OrderDate", "Order Dt"] },
  { "db_column": "Total_Amount" }
]
```
- `db_column` (required) — the exact name you want in the output.
- `aliases` (optional) — variants you've already seen. These are matched
  exactly (100% confidence). You don't *need* aliases — the fuzzy matcher
  will usually still find `CustName` → `Customer_Name` on its own — but
  listing known variants makes matching more reliable and predictable.

## 3. Run it

```bash
python excel_column_mapper.py your_file.xlsx --config target_columns.json
```

Optional flags:
| Flag | Default | Purpose |
|---|---|---|
| `--output` | `<input>_mapped.xlsx` | Custom output path |
| `--sheet` | first sheet | Sheet name or index to read |
| `--threshold` | `75` | Minimum fuzzy score (0-100) to accept a match |

## How matching works, in order of priority

1. **Exact** — normalized header matches the `db_column` name exactly
   (case/space/underscore-insensitive)
2. **Alias** — normalized header matches one of your listed `aliases` exactly
3. **Fuzzy** — if neither hits, a similarity score (0-100) is computed against
   the db_column and all its aliases; the best score above `--threshold` wins
4. **Unmatched** — nothing cleared the threshold; the column is left as-is in
   the output and flagged red in the report, so you know to add it as a new
   DB column or as an alias for next time

If two Excel columns both end up wanting the same DB column, the
higher-confidence one wins and the other is marked `Unmatched (duplicate
claim)` rather than silently overwriting data.

## Tip: building your config over time

Every time the tool flags an `Unmatched` or low-confidence `Fuzzy` column in
the report, just add that header string into the right column's `aliases`
list in `target_columns.json`. Over a few runs the tool gets close to 100%
exact/alias matches and rarely needs to guess.
