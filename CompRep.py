import pandas as pd
import os
import sys
from itertools import combinations

# --- Config from command line ---
user_keys = None
max_lines = None
base_file = None
compare_file = None
excel_output = "comp_result.xlsx"
debug_mode = False

valid_options = ["keys=", "maxline=", "base=", "compare=", "exc=", "debug=1", "--help"]

for arg in sys.argv[1:]:
    if arg == "--help":
        print("""
AO Compare Tool

A lightweight Python script for comparing two Excel files exported from SAP Analysis for Office (AO).

USAGE:
  python CompareReports.py [options]

OPTIONS:
  keys=col1,col2         Columns to use as a unique key (comma-separated)
  maxline=N              Limit the number of rows read from Excel files
  base=filename.xlsx     Specify base Excel file
  compare=filename.xlsx  Specify compare file
  exc=filename.xlsx      Export result to Excel file (default: comp_result.xlsx)
  exc=none               Disable Excel output
  debug=1                Show column names, types, and first 3 rows
  --help                 Show this help message

EXAMPLES:
  python CompareReports.py base=a.xlsx compare=b.xlsx keys="Reference,DOC NUM"
  python CompareReports.py base=a.xlsx compare=b.xlsx maxline=500 exc=diff.xlsx
  python CompareReports.py debug=1 exc=none

OUTPUT:
- Console: Differences, only-in-base, only-in-compare
- Excel: Optional, 3 sheets (Differences, OnlyInBase, OnlyInCompare)

LICENSE:
MIT License
""")
        sys.exit(0)
    elif not any(arg.startswith(opt) for opt in valid_options):
        print(f"Unknown option: {arg}\nUse --help for usage.")
        sys.exit(1)

    if arg.startswith("keys="):
        user_keys = [k.strip() for k in arg[len("keys="):].split(",") if k.strip()]
    elif arg.startswith("maxline="):
        try:
            max_lines = int(arg[len("maxline="):])
        except ValueError:
            raise Exception("Invalid maxline value. Use an integer.")
    elif arg.startswith("base="):
        base_file = arg[len("base="):].strip()
    elif arg.startswith("compare="):
        compare_file = arg[len("compare="):].strip()
    elif arg.startswith("exc="):
        value = arg[len("exc="):].strip()
        if value.lower() == "none":
            excel_output = None
        else:
            excel_output = value
    elif arg.startswith("debug=1"):
        debug_mode = True

# --- Constants ---
min_named_cols = 3

# --- Helpers ---

def detect_header_row(df: pd.DataFrame, min_named=3):
    for i in range(len(df)):
        row = df.iloc[i]
        text_count = sum(1 for cell in row if isinstance(cell, str) and cell.strip())
        if text_count >= min_named:
            return i
    return None

def rename_unnamed_columns(columns):
    new_cols = []
    prev = "Unnamed"
    for col in columns:
        if isinstance(col, str) and col.strip().lower().startswith("unnamed:"):
            new_cols.append(f"{prev} (Text)")
        else:
            new_cols.append(col)
            prev = col
    return new_cols

def load_excel_clean(filename):
    xls = pd.ExcelFile(filename)
    sheet = next((s for s in xls.sheet_names if "hiddensheet" not in s.lower()), None)

    preview = pd.read_excel(filename, sheet_name=sheet, header=None, nrows=20)
    header_row = detect_header_row(preview, min_named_cols)

    df = pd.read_excel(filename, sheet_name=sheet, header=header_row, nrows=max_lines)
    df.columns = rename_unnamed_columns(df.columns)

    if debug_mode:
        print(f"\n--- DEBUG: Loaded columns from {filename} ---")
        print(df.columns.tolist())
        print("\n--- DEBUG: Column types ---")
        print(df.dtypes)
        print("\n--- DEBUG: First 3 rows ---")
        print(df.head(3))

    return df

def find_unique_key(df: pd.DataFrame, max_columns=3):
    cols = [
        c for c in df.columns
        if not df[c].isnull().all() and not pd.api.types.is_float_dtype(df[c])
    ]
    for col in cols:
        if df[col].is_unique:
            return [col]
    for r in range(2, max_columns + 1):
        for combo in combinations(cols, r):
            if df[list(combo)].astype(str).drop_duplicates().shape[0] == df.shape[0]:
                return list(combo)
    return None

# --- File validation ---
if base_file and not os.path.isfile(base_file):
    raise Exception(f"Base file not found: {base_file}")
if compare_file and not os.path.isfile(compare_file):
    raise Exception(f"Compare file not found: {compare_file}")

# --- File auto-detection fallback ---
if not base_file or not compare_file:
    files = sorted([f for f in os.listdir() if f.lower().endswith(".xlsx")])
    if len(files) < 2:
        raise Exception("At least two .xlsx files are required in the current directory.")
    if not base_file:
        base_file = files[0]
    if not compare_file:
        compare_file = files[1]

print("Base file:", base_file)
print("Compare file:", compare_file)

# --- Load files ---
df_base = load_excel_clean(base_file)
df_compare = load_excel_clean(compare_file)

# --- Validate user-specified keys ---
if user_keys:
    missing_base = [k for k in user_keys if k not in df_base.columns]
    missing_compare = [k for k in user_keys if k not in df_compare.columns]
    if missing_base:
        raise Exception(f"The following key columns are missing in the base file: {', '.join(missing_base)}")
    if missing_compare:
        raise Exception(f"The following key columns are missing in the compare file: {', '.join(missing_compare)}")
    key_columns = user_keys
    print("Using user-defined key column(s):", ", ".join(key_columns))
else:
    key_columns = find_unique_key(df_base)
    if not key_columns:
        raise Exception("No unique key columns found in base file.")
    print("Detected key column(s):", ", ".join(key_columns))

# --- Normalize key types ---
for df in (df_base, df_compare):
    for col in key_columns:
        df[col] = df[col].astype(str)

# --- Merge and compare ---
merged = pd.merge(df_base, df_compare, how="outer", on=key_columns, suffixes=('_base', '_compare'), indicator=True)

# --- Collect results ---
diff_rows = []
only_in_base = []
only_in_compare = []

for _, row in merged[merged['_merge'] == 'both'].iterrows():
    differences = []
    for col in df_base.columns:
        if col in key_columns or f"{col}_compare" not in merged.columns:
            continue
        val_a = row.get(f"{col}_base")
        val_b = row.get(f"{col}_compare")
        if pd.isna(val_a) and pd.isna(val_b):
            continue
        if val_a != val_b:
            differences.append(f"{col}: {val_a} / {val_b}")
    if differences:
        key_repr = " | ".join(str(row[k]) for k in key_columns)
        diff_rows.append({"Key": key_repr, "Differences": " | ".join(differences)})

for _, row in merged[merged['_merge'] == 'left_only'].iterrows():
    only_in_base.append(" | ".join(str(row[k]) for k in key_columns))

for _, row in merged[merged['_merge'] == 'right_only'].iterrows():
    only_in_compare.append(" | ".join(str(row[k]) for k in key_columns))

# --- Print to console ---
print("\n--- Differences on matching keys ---")
for r in diff_rows:
    print(f"{r['Key']} | {r['Differences']}")

print("\n--- Present in base only ---")
for r in only_in_base:
    print(r)

print("\n--- Present in compare only ---")
for r in only_in_compare:
    print(r)

# --- Export to Excel if enabled ---
if excel_output:
    with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
        pd.DataFrame(diff_rows).to_excel(writer, sheet_name="Differences", index=False)
        pd.DataFrame({"Only in base": only_in_base}).to_excel(writer, sheet_name="OnlyInBase", index=False)
        pd.DataFrame({"Only in compare": only_in_compare}).to_excel(writer, sheet_name="OnlyInCompare", index=False)
    print(f"\nResults saved to {excel_output}")
