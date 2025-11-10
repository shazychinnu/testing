import re
import pandas as pd
from openpyxl import load_workbook, Workbook

cdr_file = "CDR_VREP.xlsx"
wizard_file = "report_file.xlsx"
output_file = "output.xlsx"

# -----------------------------
# Utilities
# -----------------------------
def ensure_output_file_exists():
    try:
        load_workbook(output_file)
    except FileNotFoundError:
        Workbook().save(output_file)

def delete_sheet_if_exists(path, sheet_name):
    wb = load_workbook(path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
        wb.save(path)
    wb.close()

def norm_key(x) -> str:
    """Normalize keys so joins are reliable (strip, drop .0, remove spaces/commas/hyphens, uppercase)."""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.replace("\u00A0", " ")  # NBSP
    s = re.sub(r"[ ,\-]", "", s)
    return s.upper()

# -----------------------------
# Step 1: Commitment Sheet
# -----------------------------
def create_commitment_sheet():
    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # ---- CDR Summary (GS source) ----
    cdr = pd.read_excel(
        cdr_file,
        sheet_name="CDR Summary By Investor",
        engine="openpyxl",
        skiprows=2,
    )
    cdr.columns = cdr.columns.str.strip()
    cdr["Account Number"] = cdr["Account Number"].apply(norm_key)
    # value source
    acctnorm_to_commitment = cdr.set_index("Account Number")["Investor Commitment"].to_dict()

    # ---- Data_format (left block) ----
    df = pd.read_excel(wizard_file, sheet_name="Data_format", engine="openpyxl")
    df.columns = df.columns.str.strip()

    required = ["Legal Entity", "Bin ID", "Investran Acct ID", "Commitment Amount"]
    for col in required:
        if col not in df.columns:
            raise KeyError(f"Missing required column in Data_format: '{col}'")

    df["Legal Entity"] = df["Legal Entity"].astype(str).str.strip()
    df["Commitment Amount"] = pd.to_numeric(df["Commitment Amount"], errors="coerce").fillna(0)

    # normalized join keys
    df["_bin_norm"] = df["Bin ID"].apply(norm_key)
    df["_inv_acct_norm"] = df["Investran Acct ID"].apply(norm_key)

    subtotal_mask = df["Legal Entity"].str.contains("Subtotal", case=False, na=False)

    # ---- GS Commitment (Bin ID ↔ Account Number); only where Bin ID available & not subtotal ----
    df["GS Commitment"] = pd.NA
    valid_gs = (~subtotal_mask) & df["_bin_norm"].ne("") & df["_bin_norm"].notna()
    df.loc[valid_gs, "GS Commitment"] = df.loc[valid_gs, "_bin_norm"].map(acctnorm_to_commitment)

    # Compute subtotal rows (sum above blocks)
    sub_idxs = df.index[subtotal_mask].tolist()
    if sub_idxs:
        sub_idxs.append(len(df))
        start = 0
        for idx in sub_idxs[:-1]:
            block = df.iloc[start:idx]
            df.at[idx, "Commitment Amount"] = block["Commitment Amount"].sum()
            df.at[idx, "GS Commitment"] = pd.to_numeric(block["GS Commitment"], errors="coerce").sum(skipna=True)
            start = idx + 1

    # Per rule: subtotal rows should display blank GS
    df.loc[subtotal_mask, "GS Commitment"] = pd.NA
    df["GS Check"] = df["Commitment Amount"] - pd.to_numeric(df["GS Commitment"], errors="coerce")

    # ---- SS Commitment source: sum(Commitment Amount) by Investran Acct ID (non-blank) ----
    ss_source = (
        df.loc[df["_inv_acct_norm"].ne("") & df["_inv_acct_norm"].notna()]
        .groupby("_inv_acct_norm", as_index=True)["Commitment Amount"]
        .sum()
        .to_dict()
    )

    # ---- investern_format (right block) ----
    investern = pd.read_excel(wizard_file, sheet_name="investern_format", engine="openpyxl")
    investern.columns = investern.columns.str.strip()

    for col in ["Investor ID", "Invester Commitment"]:
        if col not in investern.columns:
            raise KeyError(f"Missing required column in investern_format: '{col}'")

    investern["_id_norm"] = investern["Investor ID"].apply(norm_key)
    investern["Invester Commitment"] = pd.to_numeric(investern["Invester Commitment"], errors="coerce")

    # Map SS only for rows with a valid Investor ID key; leave others blank
    investern["SS Commitment"] = investern["_id_norm"].map(ss_source)
    investern.loc[investern["_id_norm"].eq("") | investern["_id_norm"].isna(), "SS Commitment"] = pd.NA
    investern["SS Commitment"] = pd.to_numeric(investern["SS Commitment"], errors="coerce")

    investern["SS Check"] = investern["SS Commitment"] - investern["Invester Commitment"]

    # Append one grand-total row for SS block
    ss_total_row = {
        "Vehicle/Investor": "Subtotal (SS Total)",
        "Investor ID": "",
        "Invester Commitment": investern["Invester Commitment"].sum(skipna=True),
        "SS Commitment": investern["SS Commitment"].sum(skipna=True),
        "SS Check": investern["SS Commitment"].sum(skipna=True) - investern["Invester Commitment"].sum(skipna=True),
        "_id_norm": "",
    }
    investern = pd.concat([investern, pd.DataFrame([ss_total_row])], ignore_index=True)

    # ---- Combine left + spacer + right ----
    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)})

    left = df.reindex(range(max_rows)).reset_index(drop=True)
    right = investern.reindex(range(max_rows)).reset_index(drop=True)
    combined_df = pd.concat([left, spacer, right], axis=1)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as w:
        combined_df.to_excel(w, sheet_name="Commitment Sheet", index=False)

    print("✅ Commitment Sheet created.")
    return combined_df


# -----------------------------
# Step 2: Entry Sheet
# -----------------------------
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
            block["Final LE Amount"] = pd.to_numeric(block["Final LE Amount"], errors="coerce")
            subtotal_val = block["Final LE Amount"].sum()
        else:
            subtotal_val = 0

        subtotal_row = {col: "" for col in block.columns}
        subtotal_row[block.columns[0]] = "Subtotal"
        if "Final LE Amount" in block.columns:
            subtotal_row["Final LE Amount"] = subtotal_val

        block = pd.concat([block, pd.DataFrame([subtotal_row])], ignore_index=True)
        tables.append(block)

    final_df = pd.concat(tables, ignore_index=True) if tables else pd.DataFrame()

    # Normalize Investor ID in allocation sheet
    id_col = "Investor ID" if "Investor ID" in final_df.columns else "Investor Id"
    final_df["_id_norm"] = final_df[id_col].apply(norm_key)

    # Build maps from commitment_df (left block columns are present there)
    cm = commitment_df.copy()
    cm["_inv_acct_norm"] = cm["Investran Acct ID"].apply(norm_key)

    # Bin ID: pick first non-null per Investran key
    id_to_bin = (
        cm.dropna(subset=["Bin ID"])
          .drop_duplicates(subset=["_inv_acct_norm"])
          .set_index("_inv_acct_norm")["Bin ID"]
          .to_dict()
    )
    # Commitment Amount: sum per Investran key
    id_to_amt = cm.groupby("_inv_acct_norm", as_index=True)["Commitment Amount"].sum().to_dict()

    final_df["Bin ID"] = final_df["_id_norm"].map(id_to_bin)
    final_df["Commitment Amount"] = final_df["_id_norm"].map(id_to_amt)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as w:
        final_df.to_excel(w, sheet_name="Entry", index=False)

    print("✅ Entry Sheet created.")

# -----------------------------
# Main
# -----------------------------
if __name__ == "__main__":
    commitment_df = create_commitment_sheet()
    create_entry_sheet_with_subtotals(commitment_df)
    print("🎯 Automation completed successfully!")
