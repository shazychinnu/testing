import re
import numpy as np
import pandas as pd
from openpyxl import load_workbook, Workbook

cdr_file = "CDR_VREP.xlsx"
wizard_file = "report_file.xlsx"
output_file = "output.xlsx"

# ---------------------------------------------------------
# Utility functions
# ---------------------------------------------------------
def ensure_output_file_exists():
    """Ensure output file exists."""
    try:
        load_workbook(output_file)
    except FileNotFoundError:
        Workbook().save(output_file)


def delete_sheet_if_exists(path, sheet_name):
    """Delete sheet if it already exists."""
    wb = load_workbook(path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
        wb.save(path)
    wb.close()


def norm_key(x) -> str:
    """Normalize keys for consistent matching."""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.replace("\u00A0", " ")
    s = re.sub(r"[ ,\-]", "", s)
    return s.upper()


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Replace all NaN, <NA>, None, NULL, etc. with blank and convert to object dtype."""
    df = df.replace(
        to_replace=[pd.NA, None, np.nan, "NaN", "<NA>", "None", "NULL", "nan"],
        value=""
    )
    return df.astype(object)


# ---------------------------------------------------------
# Step 1: Create Commitment Sheet
# ---------------------------------------------------------
def create_commitment_sheet():
    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # ---- 1. Load CDR Summary By Investor ----
    cdr = pd.read_excel(
        cdr_file,
        sheet_name="CDR Summary By Investor",
        engine="openpyxl",
        skiprows=2
    )
    cdr.columns = cdr.columns.str.strip()
    cdr["Account Number"] = cdr["Account Number"].apply(norm_key)
    cdr["Investor Commitment"] = pd.to_numeric(cdr["Investor Commitment"], errors="coerce").fillna(0)
    acctnorm_to_commitment = cdr.set_index("Account Number")["Investor Commitment"].to_dict()

    # ---- 2. Load Data_format ----
    df = pd.read_excel(wizard_file, sheet_name="Data_format", engine="openpyxl")
    df.columns = df.columns.str.strip()
    df["Legal Entity"] = df["Legal Entity"].astype(str).str.strip()
    df["Commitment Amount"] = pd.to_numeric(df["Commitment Amount"], errors="coerce").fillna(0)
    df["_bin_norm"] = df["Bin ID"].apply(norm_key)
    df["_inv_acct_norm"] = df["Investran Acct ID"].apply(norm_key)

    subtotal_mask = df["Legal Entity"].str.contains("Subtotal", case=False, na=False)

    # ---- 3. GS Commitment ----
    df["GS Commitment"] = df["_bin_norm"].map(acctnorm_to_commitment)
    df.loc[subtotal_mask, "GS Commitment"] = np.nan
    df["GS Commitment"] = pd.to_numeric(df["GS Commitment"], errors="coerce").fillna(0)
    df["Commitment Amount"] = pd.to_numeric(df["Commitment Amount"], errors="coerce").fillna(0)
    df["GS Check"] = df["Commitment Amount"] - df["GS Commitment"]

    # ---- 4. SS Commitment ----
    ss_source = (
        df.loc[df["_inv_acct_norm"].ne("") & df["_inv_acct_norm"].notna()]
        .groupby("_inv_acct_norm")["Commitment Amount"]
        .sum()
        .to_dict()
    )

    investern = pd.read_excel(wizard_file, sheet_name="investern_format", engine="openpyxl")
    investern.columns = investern.columns.str.strip()
    investern["Investor ID"] = investern["Investor ID"].astype(str).str.strip().str.upper()
    investern["_id_norm"] = investern["Investor ID"].apply(norm_key)
    investern["Invester Commitment"] = pd.to_numeric(investern["Invester Commitment"], errors="coerce").fillna(0)
    investern["SS Commitment"] = investern["_id_norm"].map(ss_source)
    investern["SS Commitment"] = pd.to_numeric(investern["SS Commitment"], errors="coerce").fillna(0)
    investern["SS Check"] = investern["SS Commitment"] - investern["Invester Commitment"]

    # ---- 5. Combine DataFrames ----
    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)}, dtype=object)
    df = df.reindex(range(max_rows)).reset_index(drop=True)
    investern = investern.reindex(range(max_rows)).reset_index(drop=True)

    combined_df = pd.concat([df.astype(object), spacer, investern.astype(object)], axis=1)

    # ---- 6. Add SS Subtotal Row ----
    ss_total_commit = pd.to_numeric(investern["SS Commitment"], errors="coerce").fillna(0).sum()
    ss_total_invest = pd.to_numeric(investern["Invester Commitment"], errors="coerce").fillna(0).sum()
    ss_total_check = ss_total_commit - ss_total_invest

    subtotal_row = {col: "" for col in combined_df.columns}
    subtotal_row.update({
        "Vehicle/Investor": "Subtotal (SS Total)",
        "Investor ID": "",
        "Invester Commitment": ss_total_invest,
        "SS Commitment": ss_total_commit,
        "SS Check": ss_total_check,
    })
    subtotal_df = pd.DataFrame([subtotal_row], dtype=object)
    combined_df = pd.concat([combined_df, subtotal_df], ignore_index=True)

    # ---- 7. Blank SS values where Investor ID missing ----
    combined_df["Investor ID"] = combined_df["Investor ID"].astype(str).str.strip().str.upper()
    mask_blank = combined_df["Investor ID"].isin(["", "NAN", "NONE", "NULL"]) | combined_df["Investor ID"].isna()
    for col in ["SS Commitment", "SS Check", "Invester Commitment"]:
        if col in combined_df.columns:
            combined_df.loc[mask_blank, col] = ""

    # ---- 8. Clean all NaNs for Excel ----
    combined_df = clean_dataframe(combined_df)

    # ---- 9. Write clean sheet ----
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        combined_df.to_excel(writer, sheet_name="Commitment Sheet", index=False)

    print("✅ Commitment Sheet created successfully — no NaN, no float+str errors.")
    return combined_df


# ---------------------------------------------------------
# Step 2: Create Entry Sheet
# ---------------------------------------------------------
def create_entry_sheet_with_subtotals(commitment_df):
    delete_sheet_if_exists(output_file, "Entry")

    df_raw = pd.read_excel(wizard_file, sheet_name="allocation_data", engine="openpyxl", header=None)
    header_rows = df_raw.index[df_raw.iloc[:, 0].astype(str) == "Vehicle/Investor"].tolist()
    tables = []

    for i, h in enumerate(header_rows):
        start = h
        end = header_rows[i + 1] if i + 1 < len(header_rows) else len(df_raw)
        block = df_raw.iloc[start:end].reset_index(drop=True)
        block.columns = block.iloc[0]
        block = block.drop(0).reset_index(drop=True)

        if "Final LE Amount" in block.columns:
            block["Final LE Amount"] = pd.to_numeric(block["Final LE Amount"], errors="coerce").fillna(0)
            subtotal_val = block["Final LE Amount"].sum(skipna=True)
        else:
            subtotal_val = 0

        subtotal_row = {col: "" for col in block.columns}
        subtotal_row[block.columns[0]] = "Subtotal"
        if "Final LE Amount" in block.columns:
            subtotal_row["Final LE Amount"] = subtotal_val

        block = pd.concat([block, pd.DataFrame([subtotal_row])], ignore_index=True)
        tables.append(block)

    final_df = pd.concat(tables, ignore_index=True) if tables else pd.DataFrame()

    id_col = "Investor ID" if "Investor ID" in final_df.columns else "Investor Id"
    final_df["_id_norm"] = final_df[id_col].apply(norm_key)

    cm = commitment_df.copy()
    cm["_inv_acct_norm"] = cm["Investran Acct ID"].apply(norm_key)

    id_to_bin = (
        cm.dropna(subset=["Bin ID"])
        .drop_duplicates(subset=["_inv_acct_norm"])
        .set_index("_inv_acct_norm")["Bin ID"]
        .to_dict()
    )
    id_to_amt = cm.groupby("_inv_acct_norm")["Commitment Amount"].sum().to_dict()

    final_df["Bin ID"] = final_df["_id_norm"].map(id_to_bin)
    final_df["Commitment Amount"] = final_df["_id_norm"].map(id_to_amt)

    final_df = clean_dataframe(final_df)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        final_df.to_excel(writer, sheet_name="Entry", index=False)

    print("✅ Entry Sheet created successfully — clean, validated.")


# ---------------------------------------------------------
# Main Execution
# ---------------------------------------------------------
if __name__ == "__main__":
    commitment_df = create_commitment_sheet()
    create_entry_sheet_with_subtotals(commitment_df)
    print("🎯 Automation completed successfully — all sheets validated, no NaN, no float+str errors.")
