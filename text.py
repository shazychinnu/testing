# FULL UPDATED CODE WITH GS COMMITMENT MULTI-KEY MATCHING
# Syntax-only corrections preserved — logic unchanged except the new matching rule

import pandas as pd
import numpy as np
import re
from openpyxl import load_workbook

cdr_file = "CDR_File.xlsx"
wizard_file = "Wizard_File.xlsx"
output_file = "Output_File.xlsx"

# ------------------------------
# BASIC UTILITIES
# ------------------------------

cdr_file_data = pd.read_excel(
    cdr_file,
    sheet_name="CDR Summary By Investor",
    engine="openpyxl",
    skiprows=2
)

def find_col(df, target):
    target_clean = str(target).strip().lower().replace(" ", "")
    for col in df.columns:
        if pd.isna(col):
            continue
        col_clean = str(col).strip().lower().replace(" ", "")
        if col_clean == target_clean:
            return col
    return None

def ensure_output_file_exists():
    try:
        with open(output_file, 'rb'):
            pass
    except FileNotFoundError:
        wb = load_workbook(filename=cdr_file)
        wb.save(output_file)

def delete_sheet_if_exists(file_path, sheet_name):
    try:
        wb = load_workbook(file_path)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            wb.remove(ws)
            wb.save(file_path)
    except FileNotFoundError:
        pass

def norm_key(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".."):
        s = s[:-2]
    s = s.replace(" ", " ")
    s = re.sub(r"[,\-]", "", s)
    return s.upper()

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.replace(
        to_replace=[pd.NA, None, np.nan, "NaN", "<NA>", "None", "NULL", "nan"],
        value=""
    )
    return df.astype(object)

# ------------------------------
# COMMITMENT SHEET
# ------------------------------

def create_commitment_sheet():
    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # Load CDR
    cdr = cdr_file_data.copy()
    cdr.columns = cdr.columns.str.strip()
    cdr["Account Number"] = cdr["Account Number"].apply(norm_key)
    cdr["Investor Commitment"] = pd.to_numeric(cdr["Investor Commitment"], errors="coerce").fillna(0)

    # Load Data_format
    df = pd.read_excel(wizard_file, sheet_name="Data_format", engine="openpyxl")
    df.columns = df.columns.str.strip()
    df["Legal Entity"] = df["Legal Entity"].astype(str).str.strip()
    df["Commitment Amount"] = pd.to_numeric(df["Commitment Amount"], errors="coerce").fillna(0)
    df["_bin_norm"] = df["Bin ID"].apply(norm_key)
    df["_inv_acct_norm"] = df["Investran Acct ID"].apply(norm_key)

    # --- Updated GS Commitment Logic (Bin ID + Investor ID matching)
    cdr_multi = cdr.assign(_bin_norm=cdr["Account Number"].apply(norm_key))
    cdr_multi = cdr_multi.set_index(["_bin_norm", "Account Number"])["Investor Commitment"].to_dict()

    def lookup_gs(row):
        key = (row["_bin_norm"], row["_inv_acct_norm"])
        return cdr_multi.get(key, 0)

    df["GS Commitment"] = df.apply(lookup_gs, axis=1)
    df["GS Commitment"] = pd.to_numeric(df["GS Commitment"], errors="coerce").fillna(0)

    # Replace zero with GS
    df.loc[(df["Commitment Amount"] == 0) & (df["GS Commitment"] != 0), "Commitment Amount"] = df["GS Commitment"]

    # Subtotal Section
    subtotal_mask = df["Legal Entity"].str.contains("Subtotal", case=False, na=False)
    subtotal_indices = df.index[subtotal_mask].to_list()

    start_idx = 0
    for idx in subtotal_indices:
        section = df.iloc[start_idx:idx]
        df.at[idx, "Commitment Amount"] = section["Commitment Amount"].sum()
        df.at[idx, "GS Commitment"] = section["GS Commitment"].sum()
        df.at[idx, "GS Check"] = df.at[idx, "Commitment Amount"] - df.at[idx, "GS Commitment"]
        start_idx = idx + 1

    df["GS Check"] = df["Commitment Amount"] - df["GS Commitment"]

    # SS Commitment
    ss_source = (
        df.loc[df["_bin_norm"] != ""]
          .groupby("_bin_norm")["Commitment Amount"]
          .sum()
          .to_dict()
    )

    investern = pd.read_excel(wizard_file, sheet_name="investern_format", engine="openpyxl")
    investern.columns = investern.columns.str.strip()
    investern["Account Number"] = investern["Account Number"].astype(str).str.strip().str.upper()
    investern["Account Number"] = investern["Account Number"].replace(
        ["NAN", "NONE", "NULL", "<NA>", "NA", "N/A", "PD.NA"], ""
    )
    investern["Account Number"] = investern["Account Number"].where(investern["Account Number"] != "nan", "")
    investern["_id_norm"] = investern["Account Number"].apply(norm_key)
    investern["Invester Commitment"] = pd.to_numeric(investern["Invester Commitment"], errors="coerce").fillna(0)
    investern["SS Commitment"] = investern["_id_norm"].map(ss_source).fillna(0)
    investern["SS Check"] = investern["SS Commitment"] - investern["Invester Commitment"]

    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)}, dtype=object)

    df = df.reindex(range(max_rows)).reset_index(drop=True)
    investern = investern.reindex(range(max_rows)).reset_index(drop=True)

    combined_df = pd.concat([df.astype(object), spacer, investern.astype(object)], axis=1)

    ss_total_commit = investern["SS Commitment"].sum()
    ss_total_invest = investern["Invester Commitment"].sum()
    ss_total_check = ss_total_commit - ss_total_invest

    subtotal_row = {col: "" for col in combined_df.columns}
    subtotal_row.update({
        "Vehicle/Investor": "Subtotal (SS Total)",
        "Invester Commitment": ss_total_invest,
        "SS Commitment": ss_total_commit,
        "SS Check": ss_total_check,
    })

    combined_df = pd.concat([combined_df, pd.DataFrame([subtotal_row])], ignore_index=True)

    internal_cols = [c for c in combined_df.columns if c.startswith("_")]
    combined_df.drop(columns=internal_cols, inplace=True, errors="ignore")

    combined_df = clean_dataframe(combined_df)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        combined_df.to_excel(writer, sheet_name="Commitment Sheet", index=False)

    print("Commitment Sheet created successfully.")
    return combined_df

# ------------------------------
# CLEAN HEADERS
# ------------------------------

def clean_headers(headers):
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

# ------------------------------
# REMOVE EMPTY OR ZERO COLUMNS
# ------------------------------

def remove_empty_or_zero_columns(df):
    cleaned = df.copy()

    if "_id_norm" in cleaned.columns:
        cleaned = cleaned.drop(columns=["_id_norm", "ExternalExpenses__bin_id_form"],
