import pandas as pd
import re
from openpyxl import load_workbook

# -----------------------------------------------------
# BASIC UTILITIES
# -----------------------------------------------------

def norm_key(x):
    """Normalize ID keys for matching (remove spaces, leading zeros, commas, .0 formatting)."""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.replace("\u00A0", " ")
    s = re.sub(r"[,\-]", "", s)
    return s.upper()


def clean_headers(headers):
    """Clean column headers and ensure no duplicates."""
    cleaned = []
    seen = {}

    for i, col in enumerate(headers):
        col_str = str(col).strip() if pd.notna(col) and str(col).strip() != "" else f"Unnamed_{i}"

        if col_str in seen:
            seen[col_str] += 1
            col_str = f"{col_str}_{seen[col_str]}"
        else:
            seen[col_str] = 0

        cleaned.append(col_str)

    return cleaned


def find_col(df, target):
    """Find column name matching target ignoring case and spaces."""
    target_clean = str(target).strip().lower().replace(" ", "")
    for col in df.columns:
        if pd.isna(col):
            continue
        col_clean = str(col).strip().lower().replace(" ", "")
        if col_clean == target_clean:
            return col
    return None


def clean_dataframe(df):
    """Remove NaN → empty string."""
    return df.replace({pd.NA: "", None: "", float("nan"): ""})


def delete_sheet_if_exists(file_path, sheet_name):
    """Delete sheet safely if present."""
    try:
        wb = load_workbook(file_path)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            wb.remove(ws)
            wb.save(file_path)
    except FileNotFoundError:
        pass


# -----------------------------------------------------
# REMOVE EMPTY & ZERO-ONLY COLUMNS
# -----------------------------------------------------

def remove_empty_or_zero_columns(df):
    """Remove columns where all data is empty/blank/zero, considering subtotal rows."""
    cleaned_df = df.copy()

    # Identify subtotal rows
    subtotal_mask = cleaned_df.iloc[:, 0].astype(str).str.strip().str.lower() == "subtotal"

    for col in cleaned_df.columns:
        if col == cleaned_df.columns[0]:
            continue

        col_vals = cleaned_df[col]

        subtotal_vals = col_vals[subtotal_mask]
        normal_vals = col_vals[~subtotal_mask]

        normal_numeric = pd.to_numeric(normal_vals.replace("", pd.NA), errors="coerce")
        subtotal_numeric = pd.to_numeric(subtotal_vals.replace("", pd.NA), errors="coerce")

        all_empty = normal_vals.replace("", pd.NA).isna().all()
        subtotal_zero = subtotal_numeric.fillna(0).sum() == 0

        if all_empty and subtotal_zero:
            cleaned_df.drop(columns=[col], inplace=True)

    return cleaned_df


# -----------------------------------------------------
# ENTRY SHEET CREATION (FINAL VERSION)
# -----------------------------------------------------

def create_entry_sheet_with_subtotals(commitment_df):
    delete_sheet_if_exists(output_file, "Entry")

    # ---- Read allocation data ----
    df_raw = pd.read_excel(
        wizard_file, sheet_name="allocation_data", engine="openpyxl", header=None
    )

    header_rows = df_raw.index[df_raw.iloc[:, 0].astype(str) == "Vehicle/Investor"].tolist()
    tables = []

    # ---- Read CDR summary ----
    cdr_summary = pd.read_excel(
        cdr_file, sheet_name="CDR Summary By Investor",
        engine="openpyxl", skiprows=2
    )

    # ---- Assign headers ----
    if header_rows:
        header_values = list(df_raw.iloc[header_rows[0]])
        header_values = clean_headers(header_values)
        df_raw.columns = header_values
        df_raw = df_raw.drop(header_rows[0]).reset_index(drop=True)

    # ---- Investor ID column ----
    id_col = find_col(df_raw, "Investor ID") or find_col(df_raw, "Investor Id")
    if not id_col:
        raise KeyError(f"No Investor ID column found. Columns={list(df_raw.columns)}")

    df_raw["_id_norm"] = df_raw[id_col].apply(norm_key)

    # ---- Prepare mappings ----
    cm = commitment_df.copy()
    cdr = cdr_summary.copy()

    cm["_inv_acct_norm"] = cm["Investran Acct ID"].apply(norm_key)

    id_to_bin = (
        cm.dropna(subset=["Bin ID"])
          .drop_duplicates(subset=["_inv_acct_norm"])
          .set_index("_inv_acct_norm")["Bin ID"]
          .to_dict()
    )

    cdr["_bin_id_form"] = cdr["Account Number"].apply(norm_key)

    id_to_amt = (
        cdr.groupby("_bin_id_form")["Investor Commitment"].sum().to_dict()
    )

    # ---- Add Bin ID & Commitment Amount ----
    df_raw["Bin ID"] = df_raw["_id_norm"].map(id_to_bin)

    df_raw["Commitment Amount"] = df_raw["Bin ID"].apply(
        lambda x: id_to_amt.get(norm_key(x), "")
        if pd.notna(x) and str(x).strip() != ""
        else ""
    )

    # ---- Clean section headers ----
    cdr_summary.columns = clean_headers(cdr_summary.columns)

    section_cols = []
    new_columns = []
    section = None

    for col in cdr_summary.columns:
        col_clean = str(col).strip().lower()

        if col_clean == "total contributions to commitment":
            section = "Contributions"
        elif col_clean == "total recallable":
            section = "Distributions"
        elif col_clean == "external expenses":
            section = "ExternalExpenses"

        if section and col not in ["Investor ID", "Account Number", "Investor Name", "Bin ID"]:
            if "total" not in col_clean:
                new_col = f"{section}_{col}"
                section_cols.append(new_col)
                new_columns.append(new_col)
            else:
                new_columns.append(f"{section}_{col}")
        else:
            new_columns.append(col)

    cdr_summary.columns = new_columns

    cdr_bin_col = find_col(cdr_summary, "Account Number")
    cdr_summary[cdr_bin_col] = cdr_summary[cdr_bin_col].astype(str).apply(norm_key)

    if not cdr_summary[cdr_bin_col].is_unique:
        cdr_summary = cdr_summary.drop_duplicates(subset=[cdr_bin_col])

    cdr_summary_indexed = cdr_summary.set_index(cdr_bin_col)

    # ---- Process blocks ----
    for i, h in enumerate(header_rows):
        start = h
        end = header_rows[i+1] if i+1 < len(header_rows) else len(df_raw)

        block = df_raw.loc[start:end].reset_index(drop=True)
        block = block.drop(0).reset_index(drop=True)
        block.columns = df_raw.columns

        entry_bin_col = find_col(block, "Bin ID")

        # ---- Matching CDR Summary values ----
        for col in section_cols:
            if entry_bin_col and col in cdr_summary_indexed.columns:
                block[col] = block[entry_bin_col].apply(
                    lambda x:
                        cdr_summary_indexed[col].get(norm_key(x), "")
                        if pd.notna(x) and str(x).strip() != "" else ""
                )
            else:
                block[col] = ""

        # ---- Add subtotal ----
        numeric_cols = block.select_dtypes(include="number").columns
        subtotal_row = {
            col: block[col].sum(skipna=True) if col in numeric_cols else ""
            for col in block.columns
        }
        subtotal_row[block.columns[0]] = "Subtotal"

        block = pd.concat(
            [block, pd.DataFrame([subtotal_row], columns=block.columns)], ignore_index=True
        )
        tables.append(block)

    # ---- Combine blocks ----
    final_df = pd.concat(tables, ignore_index=True)
    final_df = clean_dataframe(final_df)

    # ---- Remove empty or zero-value columns ----
    final_df = remove_empty_or_zero_columns(final_df)

    # ---- Write final Entry sheet ----
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        final_df.to_excel(writer, sheet_name="Entry", index=False)
