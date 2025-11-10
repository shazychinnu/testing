import pandas as pd
from openpyxl import load_workbook, Workbook

cdr_file = "CDR_VREP.xlsx"
wizard_file = "report_file.xlsx"
output_file = "output.xlsx"


# ---------------------------------------------------------
# Utilities
# ---------------------------------------------------------
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


# ---------------------------------------------------------
# Step 1: Commitment Sheet
# ---------------------------------------------------------
def create_commitment_sheet():
    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # ---- 1. Read CDR Summary (for GS Commitment) ----
    cdr = pd.read_excel(
        cdr_file,
        sheet_name="CDR Summary By Investor",
        engine="openpyxl",
        skiprows=2,
    )
    cdr.columns = cdr.columns.str.strip()
    cdr["Account Number"] = cdr["Account Number"].astype(str).str.strip()
    cdr["Investor ID"] = cdr["Investor ID"].astype(str).str.strip().str.upper()

    # Map Account Number → Investor Commitment
    acct_to_commitment = cdr.set_index("Account Number")["Investor Commitment"].to_dict()

    # ---- 2. Read Data_format ----
    df = pd.read_excel(wizard_file, sheet_name="Data_format", engine="openpyxl")
    df.columns = df.columns.str.strip()

    # Validate required columns
    required = ["Legal Entity", "Bin ID", "Investran Acct ID", "Commitment Amount"]
    for col in required:
        if col not in df.columns:
            raise KeyError(f"❌ Missing required column '{col}' in Data_format sheet.")

    df["Bin ID"] = df["Bin ID"].astype(str).str.strip()
    df["Investran Acct ID"] = df["Investran Acct ID"].astype(str).str.strip().str.upper()
    df["Commitment Amount"] = pd.to_numeric(df["Commitment Amount"], errors="coerce").fillna(0)
    df["Legal Entity"] = df["Legal Entity"].astype(str).str.strip()

    # Detect subtotal rows
    subtotal_mask = df["Legal Entity"].str.contains("Subtotal", case=False, na=False)

    # ---- 3. GS Commitment (Bin ID ↔ Account Number) ----
    df["GS Commitment"] = pd.NA
    valid_gs = (~subtotal_mask) & df["Bin ID"].ne("")
    df.loc[valid_gs, "GS Commitment"] = df.loc[valid_gs, "Bin ID"].map(acct_to_commitment)

    # Compute subtotals
    sub_idxs = df.index[subtotal_mask].tolist()
    if sub_idxs:
        sub_idxs.append(len(df))
        start = 0
        for idx in sub_idxs[:-1]:
            block = df.iloc[start:idx]
            df.at[idx, "Commitment Amount"] = block["Commitment Amount"].sum()
            df.at[idx, "GS Commitment"] = pd.to_numeric(block["GS Commitment"], errors="coerce").sum(skipna=True)
            start = idx + 1

    # Blank GS for subtotal rows
    df.loc[subtotal_mask, "GS Commitment"] = pd.NA
    df["GS Check"] = pd.to_numeric(df["Commitment Amount"], errors="coerce") - pd.to_numeric(
        df["GS Commitment"], errors="coerce"
    )

    # ---- 4. SS Commitment (Investran Acct ID ↔ Investor ID) ----
    ss_map = df.set_index("Investran Acct ID")["Commitment Amount"].to_dict()

    investern = pd.read_excel(wizard_file, sheet_name="investern_format", engine="openpyxl")
    investern.columns = investern.columns.str.strip()
    investern["Investor ID"] = investern["Investor ID"].astype(str).str.strip().str.upper()
    investern["Invester Commitment"] = pd.to_numeric(investern["Invester Commitment"], errors="coerce")

    investern["SS Commitment"] = investern["Investor ID"].map(ss_map)
    investern["SS Commitment"] = pd.to_numeric(investern["SS Commitment"], errors="coerce")
    investern["SS Check"] = investern["SS Commitment"] - investern["Invester Commitment"]

    # ---- 5. Add final subtotal row (SS total) ----
    subtotal_row = {
        "Vehicle/Investor": "Subtotal (SS Total)",
        "Investor ID": "",
        "Invester Commitment": investern["Invester Commitment"].sum(skipna=True),
        "SS Commitment": investern["SS Commitment"].sum(skipna=True),
        "SS Check": investern["SS Commitment"].sum(skipna=True)
        - investern["Invester Commitment"].sum(skipna=True),
    }
    investern = pd.concat([investern, pd.DataFrame([subtotal_row])], ignore_index=True)

    # ---- 6. Combine left & right ----
    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)})

    left = df.reindex(range(max_rows)).reset_index(drop=True)
    right = investern.reindex(range(max_rows)).reset_index(drop=True)
    combined_df = pd.concat([left, spacer, right], axis=1)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as w:
        combined_df.to_excel(w, sheet_name="Commitment Sheet", index=False)

    print("✅ Commitment Sheet created successfully.")
    return combined_df


# ---------------------------------------------------------
# Step 2: Entry Sheet
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
            block["Final LE Amount"] = pd.to_numeric(block["Final LE Amount"], errors="coerce")
            subtotal_val = block["Final LE Amount"].sum()
        else:
            subtotal_val = 0

        subtotal_row = {col: "" for col in block.columns}
        first_col = block.columns[0]
        subtotal_row[first_col] = "Subtotal"
        if "Final LE Amount" in block.columns:
            subtotal_row["Final LE Amount"] = subtotal_val
        block = pd.concat([block, pd.DataFrame([subtotal_row])], ignore_index=True)
        tables.append(block)

    final_df = pd.concat(tables, ignore_index=True) if tables else pd.DataFrame()

    id_col = "Investor ID" if "Investor ID" in final_df.columns else "Investor Id"
    final_df[id_col] = final_df[id_col].astype(str).str.strip().str.upper()

    cm = commitment_df.copy()
    cm["Investran Acct ID"] = cm["Investran Acct ID"].astype(str).str.strip().str.upper()

    id_to_bin = (
        cm.dropna(subset=["Bin ID"])
        .drop_duplicates(["Investran Acct ID"])
        .set_index("Investran Acct ID")["Bin ID"]
        .to_dict()
    )
    id_to_amt = cm.groupby("Investran Acct ID", as_index=True)["Commitment Amount"].sum().to_dict()

    final_df["Bin ID"] = final_df[id_col].map(id_to_bin)
    final_df["Commitment Amount"] = final_df[id_col].map(id_to_amt)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as w:
        final_df.to_excel(w, sheet_name="Entry", index=False)

    print("✅ Entry Sheet created successfully.")


# ---------------------------------------------------------
# Main
# ---------------------------------------------------------
if __name__ == "__main__":
    commitment_df = create_commitment_sheet()
    create_entry_sheet_with_subtotals(commitment_df)
    print("🎯 Automation completed successfully!")
