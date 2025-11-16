import pandas as pd
import numpy as np
import re
from openpyxl import load_workbook

cdr_file = "CDR_File.xlsx"
wizard_file = "Wizard_File.xlsx"
output_file = "Output_File.xlsx"

# =====================================================================
# UTILITIES
# =====================================================================

cdr_file_data = pd.read_excel(
    cdr_file,
    sheet_name="CDR Summary By Investor",
    engine="openpyxl",
    skiprows=2
)

def find_col(df, target):
    tgt = str(target).strip().lower().replace(" ", "")
    for c in df.columns:
        if pd.isna(c):
            continue
        if str(c).strip().lower().replace(" ", "") == tgt:
            return c
    return None

def ensure_output_file_exists():
    try:
        with open(output_file, "rb"):
            pass
    except FileNotFoundError:
        wb = load_workbook(cdr_file)
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
    s = s.replace("\u00A0", " ")
    s = re.sub(r"[,\-]", "", s)
    return s.upper()

def clean_dataframe(df):
    df = df.replace(
        [pd.NA, None, np.nan, "NaN", "<NA>", "None", "NULL", "nan"],
        ""
    )
    return df.astype(object)

# =====================================================================
# COMMITMENT SHEET (corrected multi-key GS + investor logic)
# =====================================================================

def create_commitment_sheet():
    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # 1. Load CDR
    cdr = cdr_file_data.copy()
    cdr.columns = cdr.columns.str.strip()

    cdr["Account Number"] = cdr["Account Number"].apply(norm_key)
    cdr["Investor Commitment"] = pd.to_numeric(
        cdr["Investor Commitment"], errors="coerce"
    ).fillna(0)

    # MULTI-KEY: (bin_norm, acct_norm)
    cdr["_bin_norm"] = cdr["Account Number"].apply(norm_key)
    cdr["_acct_norm"] = cdr["Account Number"].apply(norm_key)

    gs_map = (
        cdr.set_index(["_bin_norm", "_acct_norm"])["Investor Commitment"]
        .to_dict()
    )

    # 2. Load Data_format
    df = pd.read_excel(
        wizard_file, sheet_name="Data_format", engine="openpyxl"
    )
    df.columns = df.columns.str.strip()

    df["Legal Entity"] = df["Legal Entity"].astype(str).str.strip()
    df["Commitment Amount"] = pd.to_numeric(
        df["Commitment Amount"], errors="coerce"
    ).fillna(0)

    df["_bin_norm"] = df["Bin ID"].apply(norm_key)
    df["_inv_acct_norm"] = df["Investran Acct ID"].apply(norm_key)

    # Apply GS multi-key
    def gs_lookup(r):
        return gs_map.get((r["_bin_norm"], r["_inv_acct_norm"]), 0)

    df["GS Commitment"] = df.apply(gs_lookup, axis=1)
    df["GS Commitment"] = pd.to_numeric(df["GS Commitment"]).fillna(0)

    # Replace Commitment Amount if 0
    df.loc[
        (df["Commitment Amount"] == 0) & (df["GS Commitment"] != 0),
        "Commitment Amount"
    ] = df["GS Commitment"]

    # Subtotals
    subtotal_mask = df["Legal Entity"].str.contains("Subtotal", case=False, na=False)
    idxs = df.index[subtotal_mask].tolist()

    start = 0
    for i in idxs:
        section = df.iloc[start:i]
        tot = section["Commitment Amount"].sum()
        gtot = section["GS Commitment"].sum()
        df.at[i, "Commitment Amount"] = tot
        df.at[i, "GS Commitment"] = gtot
        df.at[i, "GS Check"] = tot - gtot
        start = i + 1

    df["GS Check"] = df["Commitment Amount"] - df["GS Commitment"]

    # 4. SS commitment
    ss_map = (
        df[df["_bin_norm"] != ""]
        .groupby("_bin_norm")["Commitment Amount"]
        .sum()
        .to_dict()
    )

    investern = pd.read_excel(
        wizard_file, sheet_name="investern_format", engine="openpyxl"
    )
    investern.columns = investern.columns.str.strip()

    investern["Account Number"] = (
        investern["Account Number"]
        .astype(str)
        .str.strip()
        .str.upper()
        .replace(["NAN","NONE","NULL","PD.NA","<NA>","NA","N/A"], "")
    )
    investern["Account Number"] = investern["Account Number"].where(
        investern["Account Number"] != "nan", ""
    )

    investern["_id_norm"] = investern["Account Number"].apply(norm_key)
    investern["Invester Commitment"] = pd.to_numeric(
        investern["Invester Commitment"], errors="coerce"
    ).fillna(0)

    investern["SS Commitment"] = (
        investern["_id_norm"].map(ss_map)
    )
    investern["SS Commitment"] = pd.to_numeric(
        investern["SS Commitment"], errors="coerce"
    ).fillna(0)

    investern["SS Check"] = (
        investern["SS Commitment"] - investern["Invester Commitment"]
    )

    # Combine DataFrames
    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)})

    df2 = df.reindex(range(max_rows)).reset_index(drop=True)
    inv2 = investern.reindex(range(max_rows)).reset_index(drop=True)

    combined = pd.concat([df2, spacer, inv2], axis=1)

    # SS subtotal row
    ss_tot = inv2["SS Commitment"].sum()
    ss_inv = inv2["Invester Commitment"].sum()
    ss_chk = ss_tot - ss_inv

    row = {c: "" for c in combined.columns}
    row.update({
        "Vehicle/Investor": "Subtotal (SS Total)",
        "Invester Commitment": ss_inv,
        "SS Commitment": ss_tot,
        "SS Check": ss_chk
    })

    combined = pd.concat([combined, pd.DataFrame([row])], ignore_index=True)

    # Remove helper cols
    helper = [c for c in combined.columns if c.startswith("_")]
    combined.drop(columns=helper, inplace=True, errors="ignore")

    combined = clean_dataframe(combined)

    # Write
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as w:
        combined.to_excel(w, sheet_name="Commitment Sheet", index=False)

    print("Commitment Sheet created successfully.")
    return combined

# =====================================================================
# HEADER CLEANING + COLUMN CLEANUP
# =====================================================================

def clean_headers(headers):
    cleaned = []
    seen = {}
    for i, col in enumerate(headers):
        s = str(col).strip() if (pd.notna(col) and str(col).strip()) else f"Unnamed_{i}"
        if s in seen:
            seen[s] += 1
            s = f"{s}_{seen[s]}"
        else:
            seen[s] = 0
        cleaned.append(s)
    return cleaned

def remove_empty_or_zero_columns(df):
    c = df.copy()

    if "_id_norm" in c.columns:
        c.drop(columns=["_id_norm", "ExternalExpenses__bin_id_form"], errors="ignore", inplace=True)

    # Remove TOTAL cols
    drop_total = []
    for col in c.columns:
        col_s = str(col).strip()
        if col_s.lower() == "total":
            drop_total.append(col)
        elif re.match(r".*_Total(\.\d+)?$", col_s, re.IGNORECASE):
            drop_total.append(col)

    c.drop(columns=drop_total, inplace=True, errors="ignore")

    # Remove empty/zero
    drop_cols = []
    for col in c.columns:
        if col == c.columns[0]:
            continue
        s = c[col].replace("", pd.NA)
        if not pd.to_numeric(s, errors="coerce").notna().any():
            if s.isna().all():
                drop_cols.append(col)
            continue
        s_num = pd.to_numeric(s, errors="coerce")
        if s.isna().all() or s_num.fillna(0).eq(0).all():
            drop_cols.append(col)

    return c.drop(columns=drop_cols, errors="ignore")

# =====================================================================
# ENTRY SHEET (FIXED TO USE COMMITMENT SHEET FOR BIN ID + COMMITMENT)
# =====================================================================

def create_entry_sheet_with_subtotals(commitment_df):
    delete_sheet_if_exists(output_file, "Entry")

    raw = pd.read_excel(
        wizard_file, sheet_name="allocation_data", engine="openpyxl", header=None
    )

    header_rows = raw.index[raw.iloc[:,0].astype(str) == "Vehicle/Investor"].tolist()

    cdr = pd.read_excel(
        cdr_file,
        sheet_name="CDR Summary By Investor",
        engine="openpyxl",
        skiprows=2
    )

    # Header apply
    if header_rows:
        hdr = clean_headers(list(raw.loc[header_rows[0]]))
        raw.columns = hdr
        raw = raw.drop(header_rows[0]).reset_index(drop=True)

    # Investor ID column from Entry sheet
    id_col = find_col(raw, "Investor ID") or find_col(raw, "Investor Id")
    if not id_col:
        raise KeyError("Investor ID column missing in allocation_data sheet")

    raw["_id_norm"] = raw[id_col].apply(norm_key)

    # === FIXED: Use Commitment Sheet for Bin ID + Commitment Amount ===

    # Build maps from commitment_df
    cm = commitment_df.copy()
    cm["_inv_norm"] = cm["Investran Acct ID"].apply(norm_key)

    bin_map = (
        cm.dropna(subset=["Investran Acct ID"])
        .set_index("_inv_norm")["Bin ID"]
        .to_dict()
    )

    commit_map = (
        cm.dropna(subset=["Investran Acct ID"])
        .set_index("_inv_norm")["Commitment Amount"]
        .to_dict()
    )

    # Apply mapping
    raw["Bin ID"] = raw["_id_norm"].map(bin_map)
    raw["Commitment Amount"] = raw["_id_norm"].map(commit_map).fillna("")

    # ==== CDR Section columns (unchanged original logic) ====
    cdr.columns = clean_headers(cdr.columns)

    new_cols = []
    sec_cols = []
    sec = None

    for col in cdr.columns:
        lc = str(col).lower().strip()
        if lc == "total contributions to commitment":
            sec = "Contributions"
        elif lc == "total recallable":
            sec = "Distributions"
        elif lc == "external expenses":
            sec = "ExternalExpenses"

        if sec and col not in ["Investor ID","Account Number","Investor Name","Bin ID"]:
            new_col = f"{sec}_{col}"
            sec_cols.append(new_col)
            new_cols.append(new_col)
        else:
            new_cols.append(col)

    cdr.columns = new_cols

    cbin = find_col(cdr, "Account Number")
    cdr[cbin] = cdr[cbin].astype(str).apply(norm_key)
    cdr = cdr.drop_duplicates(subset=[cbin])
    cdr_index = cdr.set_index(cbin)

    tables = []

    for i, h in enumerate(header_rows):
        st = h
        en = header_rows[i+1] if i+1 < len(header_rows) else len(raw)

        block = raw.loc[st:en].reset_index(drop=True)
        block = block.drop(0).reset_index(drop=True)
        block.columns = raw.columns

        # Bin ID column
        entry_bin_col = find_col(block, "Bin ID")

        for col in sec_cols:
            if entry_bin_col and col in cdr_index.columns:
                block[col] = block[entry_bin_col].apply(
                    lambda x: cdr_index[col].get(norm_key(x), "") if pd.notna(x) and str(x).strip() else ""
                )
            else:
                block[col] = ""

        # Numeric subtotals
        nums = []
        for col in block.columns:
            s = block[col].replace("", pd.NA)
            if pd.to_numeric(s, errors="coerce").notna().any():
                nums.append(col)

        sub = {c:"" for c in block.columns}
        for col in nums:
            if col in ["Investor ID", "_id_norm"]:
                continue
            sub[col] = pd.to_numeric(block[col], errors="coerce").sum(skipna=True)

        sub[block.columns[0]] = "Subtotal"
        block = pd.concat([block, pd.DataFrame([sub])], ignore_index=True)
        tables.append(block)

    final = pd.concat(tables, ignore_index=True)

    final = remove_empty_or_zero_columns(final)
    final = clean_dataframe(final)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as w:
        final.to_excel(w, sheet_name="Entry", index=False)

    print("Entry Sheet created successfully.")

# =====================================================================
# MAIN
# =====================================================================

if __name__ == "__main__":
    cdf = create_commitment_sheet()
    create_entry_sheet_with_subtotals(cdf)
    print("Automation completed successfully.")
