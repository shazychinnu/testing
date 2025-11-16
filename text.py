import pandas as pd
import numpy as np
import re
from openpyxl import load_workbook

cdr_file = "CDR_File.xlsx"
wizard_file = "Wizard_File.xlsx"
output_file = "Output_File.xlsx"

# ===============================================================
# UTILITIES (optimized)
# ===============================================================

def norm_key(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    s = s.replace("\u00A0", " ")
    # Remove punctuation and whitespace
    s = re.sub(r"[,\-\s]", "", s)
    return s.upper()

def clean_dataframe(df):
    return (
        df.replace([pd.NA, None, np.nan, "nan", "NaN", "NULL", "<NA>", "None"], "")
          .astype(object)
    )

def find_col(df, target):
    tgt = norm_key(target)
    for c in df.columns:
        if norm_key(c) == tgt:
            return c
    return None

def ensure_output_file_exists():
    try:
        open(output_file, "rb")
    except FileNotFoundError:
        load_workbook(cdr_file).save(output_file)

def delete_sheet_if_exists(path, name):
    try:
        wb = load_workbook(path)
        if name in wb.sheetnames:
            wb.remove(wb[name])
            wb.save(path)
    except:
        pass

def clean_headers(headers):
    seen = {}
    out = []
    for i, h in enumerate(headers):
        s = str(h).strip()
        if s == "":
            s = f"Unnamed_{i}"
        if s in seen:
            seen[s] += 1
            s = f"{s}_{seen[s]}"
        else:
            seen[s] = 0
        out.append(s)
    return out

# Safe numeric cast
def safe_num(series):
    return pd.to_numeric(series.replace("", pd.NA), errors="coerce")

# ===============================================================
# COMMITMENT SHEET (optimized)
# ===============================================================

def create_commitment_sheet():

    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # Load CDR only once, normalize once
    cdr = pd.read_excel(
        cdr_file, "CDR Summary By Investor",
        engine="openpyxl", skiprows=2
    )

    cdr.columns = cdr.columns.str.strip()
    cdr["Account Number"] = cdr["Account Number"].apply(norm_key)
    cdr["Investor Commitment"] = safe_num(cdr["Investor Commitment"]).fillna(0)

    # Multi-key mapping (bin + investor acct)
    cdr["_bin_norm"] = cdr["Account Number"]
    cdr["_acct_norm"] = cdr["Account Number"]

    gs_map = cdr.set_index(["_bin_norm", "_acct_norm"])["Investor Commitment"].to_dict()

    # Load Data_format once
    df = pd.read_excel(wizard_file, "Data_format", engine="openpyxl")
    df.columns = df.columns.str.strip()

    df["_bin_norm"] = df["Bin ID"].apply(norm_key)
    df["_inv_norm"] = df["Investran Acct ID"].apply(norm_key)
    df["Commitment Amount"] = safe_num(df["Commitment Amount"]).fillna(0)

    # GS Commitment (optimized vector method)
    df["GS Commitment"] = df.apply(
        lambda r: gs_map.get((r["_bin_norm"], r["_inv_norm"]), 0),
        axis=1
    )

    df["GS Commitment"] = safe_num(df["GS Commitment"]).fillna(0)

    # Replace Commitment Amount when zero
    mask = (df["Commitment Amount"] == 0) & (df["GS Commitment"] != 0)
    df.loc[mask, "Commitment Amount"] = df.loc[mask, "GS Commitment"]

    # GS Check
    df["GS Check"] = df["Commitment Amount"] - df["GS Commitment"]

    # SS Commitment
    ss_map = df.groupby("_bin_norm")["Commitment Amount"].sum().to_dict()

    invest = pd.read_excel(wizard_file, "investern_format", engine="openpyxl")
    invest.columns = invest.columns.str.strip()

    invest["Account Number"] = (
        invest["Account Number"]
        .astype(str)
        .str.upper()
        .replace(["NAN", "NA", "NONE", "NULL", "<NA>"], "")
    )

    invest["_id_norm"] = invest["Account Number"].apply(norm_key)
    invest["Invester Commitment"] = safe_num(invest["Invester Commitment"]).fillna(0)
    invest["SS Commitment"] = invest["_id_norm"].map(ss_map).fillna(0)
    invest["SS Check"] = invest["SS Commitment"] - invest["Invester Commitment"]

    # Combine frames
    max_len = max(len(df), len(invest))
    df = df.reindex(range(max_len)).reset_index(drop=True)
    invest = invest.reindex(range(max_len)).reset_index(drop=True)

    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_len for i in range(3)})
    combined = pd.concat([df, spacer, invest], axis=1)

    # SS subtotal
    row = {c: "" for c in combined.columns}
    row["Vehicle/Investor"] = "Subtotal (SS Total)"
    row["Invester Commitment"] = invest["Invester Commitment"].sum()
    row["SS Commitment"] = invest["SS Commitment"].sum()
    row["SS Check"] = row["SS Commitment"] - row["Invester Commitment"]

    combined = pd.concat([combined, pd.DataFrame([row])], ignore_index=True)

    # Remove helper cols
    drop_cols = [c for c in combined.columns if c.startswith("_")]
    combined.drop(columns=drop_cols, inplace=True, errors="ignore")

    combined = clean_dataframe(combined)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as w:
        combined.to_excel(w, "Commitment Sheet", index=False)

    return combined
# ===============================================================
# REMOVE EMPTY / ZERO COLUMNS (optimized)
# ===============================================================

def remove_empty_or_zero_columns(df):
    cleaned = df.copy()

    # Remove helper columns but keep required values
    drop_keys = ["_id_norm", "ExternalExpenses__bin_id_form"]
    cleaned.drop(columns=drop_keys, errors="ignore", inplace=True)

    # Remove totals
    total_cols = []
    for col in cleaned.columns:
        cs = str(col).strip().lower()
        if cs == "total" or re.match(r".*_total(\.\d+)?$", cs, flags=re.IGNORECASE):
            total_cols.append(col)

    cleaned.drop(columns=total_cols, errors="ignore", inplace=True)

    # Identify columns that have no values or zero-like values
    to_drop = []
    first_col = cleaned.columns[0]

    for col in cleaned.columns:
        if col == first_col:
            continue

        series = cleaned[col].replace("", pd.NA)
        numeric = safe_num(series)

        # Case 1: all empty
        if series.isna().all():
            to_drop.append(col)
            continue

        # Case 2: numeric but all 0
        if numeric.notna().any() and numeric.fillna(0).eq(0).all():
            to_drop.append(col)

    cleaned.drop(columns=to_drop, errors="ignore", inplace=True)
    return cleaned


# ===============================================================
# ENTRY SHEET (OPTIMIZED, DUPLICATES HANDLED, ALWAYS POPULATES)
# ===============================================================

def create_entry_sheet_with_subtotals(commitment_df):

    delete_sheet_if_exists(output_file, "Entry")

    # Load allocation data
    raw = pd.read_excel(
        wizard_file,
        "allocation_data",
        engine="openpyxl",
        header=None
    )

    # ----------------------------------------
    # SAFER HEADER DETECTION
    # ----------------------------------------
    header_rows = []
    for idx, row in raw.iterrows():
        normalized_row = [str(x).strip().lower() for x in row]
        if "vehicle/investor" in normalized_row:
            header_rows.append(idx)

    if not header_rows:
        raise Exception("Header row containing 'Vehicle/Investor' not found")

    # Extract the first header row
    hdr = clean_headers(list(raw.loc[header_rows[0]]))
    raw.columns = hdr

    # Remove header row
    raw = raw.drop(header_rows[0]).reset_index(drop=True)

    # ----------------------------------------
    # NORMALIZE IDS (ONLY ONCE)
    # ----------------------------------------
    id_col = find_col(raw, "Investor ID")
    raw["_id_norm"] = raw[id_col].apply(norm_key)

    # ----------------------------------------
    # COMMITMENT SHEET MAPPING (ALWAYS CORRECT)
    # ----------------------------------------
    cm = commitment_df.copy()
    cm["_inv_norm"] = cm["Investran Acct ID"].apply(norm_key)

    bin_map = cm.set_index("_inv_norm")["Bin ID"].to_dict()
    com_map = cm.set_index("_inv_norm")["Commitment Amount"].to_dict()

    raw["Bin ID"] = raw["_id_norm"].map(bin_map)
    raw["Commitment Amount"] = raw["_id_norm"].map(com_map).fillna("")

    # ----------------------------------------
    # LOAD CDR SUMMARY (FOR SECTION COLUMNS ONLY)
    # ----------------------------------------
    cdr = pd.read_excel(
        cdr_file,
        "CDR Summary By Investor",
        engine="openpyxl",
        skiprows=2
    )

    cdr.columns = clean_headers(cdr.columns)
    cdr["Account Number"] = cdr["Account Number"].astype(str).apply(norm_key)

    acct_col = find_col(cdr, "Account Number")
    cdr = cdr.drop_duplicates(subset=[acct_col]).set_index(acct_col)

    # ----------------------------------------
    # BUILD SECTION COLUMNS (optimized)
    # ----------------------------------------
    section_cols = []
    new_cols = []
    section = None

    for col in cdr.columns:
        lc = str(col).lower()

        if "contribution" in lc:
            section = "Contributions"
        elif "recallable" in lc or "distribution" in lc:
            section = "Distributions"
        elif "expense" in lc:
            section = "ExternalExpenses"

        if section and col not in ["Investor ID", "Account Number", "Investor Name", "Bin ID"]:
            new_col = f"{section}_{col}"
            section_cols.append(new_col)
            new_cols.append(new_col)
        else:
            new_cols.append(col)

    cdr.columns = new_cols

    # ----------------------------------------
    # APPLY SECTION MAPPINGS (FAST)
    # ----------------------------------------
    block = raw.copy()
    bin_col = find_col(block, "Bin ID")

    def map_section_value(x, col):
        if pd.isna(x):
            return ""
        key = norm_key(x)
        return cdr[col].get(key, "")

    for col in section_cols:
        if col in cdr.columns:
            block[col] = block[bin_col].map(lambda x, c=col: map_section_value(x, c))
        else:
            block[col] = ""

    # ----------------------------------------
    # SUBTOTAL ROW (optimized)
    # ----------------------------------------
    numeric_cols = [
        col for col in block.columns
        if safe_num(block[col]).notna().any()
    ]

    subtotal = {c: "" for c in block.columns}
    for col in numeric_cols:
        subtotal[col] = safe_num(block[col]).sum()

    subtotal[block.columns[0]] = "Subtotal"

    final_df = pd.concat([block, pd.DataFrame([subtotal])], ignore_index=True)

    # ----------------------------------------
    # FINAL CLEANUP
    # ----------------------------------------
    final_df = remove_empty_or_zero_columns(final_df)
    final_df = clean_dataframe(final_df)

    # Write Entry sheet
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as w:
        final_df.to_excel(w, "Entry", index=False)

# ===============================================================
# MAIN EXECUTION
# ===============================================================

if __name__ == "__main__":
    print("Generating Commitment Sheet...")
    commitment_df = create_commitment_sheet()

    print("Generating Entry Sheet...")
    create_entry_sheet_with_subtotals(commitment_df)

    print("Automation completed successfully — all sheets created.")
