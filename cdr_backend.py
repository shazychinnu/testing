import pandas as pd
from openpyxl import load_workbook, Workbook

cdr_file = "CDR_VREP.xlsx"
wizard_file = "report_file.xlsx"
output_file = "output.xlsx"


# ---------------------------------------------------------
# Utility functions
# ---------------------------------------------------------
def ensure_output_file_exists():
    """Make sure the output file exists before writing."""
    try:
        wb = load_workbook(output_file)
    except FileNotFoundError:
        wb = Workbook()
    wb.save(output_file)


def delete_sheet_if_exists(file_path, sheet_name):
    """Remove an Excel sheet if it already exists."""
    wb = load_workbook(file_path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
        wb.save(file_path)
    wb.close()


# ---------------------------------------------------------
# Step 1: Create Commitment Sheet
# ---------------------------------------------------------
def create_commitment_sheet():
    """Create the Commitment Sheet and return it as a DataFrame."""
    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # ---- Load CDR Summary ----
    cdr = pd.read_excel(
        cdr_file,
        sheet_name="CDR Summary By Investor",
        engine="openpyxl",
        skiprows=2
    )
    cdr.columns = cdr.columns.str.strip()
    cdr["Account Number"] = cdr["Account Number"].astype(str).str.strip()
    cdr["Investor ID"] = cdr["Investor ID"].astype(str).str.strip().str.upper()

    # Create mapping dict for GS Commitment
    account_to_commitment = cdr.set_index("Account Number")["Investor Commitment"].to_dict()

    # ---- Load Data_format ----
    df = pd.read_excel(wizard_file, sheet_name="Data_format", engine="openpyxl")
    df.columns = df.columns.str.strip()

    df["Investran Acct ID"] = df["Investran Acct ID"].astype(str).str.strip()
    df["Investor ID"] = df["Investor ID"].astype(str).str.strip().str.upper()
    df["Commitment Amount"] = pd.to_numeric(df["Commitment Amount"], errors="coerce")
    df["Legal Entity"] = df["Legal Entity"].astype(str).str.strip()

    # Identify subtotal rows
    subtotal_mask = df["Legal Entity"].str.contains("Subtotal", case=False, na=False)

    # ---- Apply GS Commitment ----
    df["GS Commitment"] = None
    valid_mask = (~subtotal_mask) & df["Investran Acct ID"].notna() & (df["Investran Acct ID"] != "")
    df.loc[valid_mask, "GS Commitment"] = df.loc[valid_mask, "Investran Acct ID"].map(account_to_commitment)

    # ---- Calculate subtotals ----
    subtotal_indices = df.index[subtotal_mask].tolist() + [len(df)]
    start_idx = 0
    for idx in subtotal_indices[:-1]:
        block = df.iloc[start_idx:idx]
        df.at[idx, "Commitment Amount"] = block["Commitment Amount"].sum()
        df.at[idx, "GS Commitment"] = block["GS Commitment"].astype(float).sum(skipna=True)
        start_idx = idx + 1

    # Blank out GS Commitment for subtotal rows
    df.loc[subtotal_mask, "GS Commitment"] = None
    df["GS Check"] = df["Commitment Amount"] - df["GS Commitment"]

    # ---- Prepare SS Commitment Mapping ----
    # Build mapping from current data itself (Investor ID → Commitment Amount)
    investorid_to_commitment = df.set_index("Investor ID")["Commitment Amount"].to_dict()

    # ---- Load investern_format ----
    investern = pd.read_excel(wizard_file, sheet_name="investern_format", engine="openpyxl")
    investern.columns = investern.columns.str.strip()
    investern["Investor ID"] = investern["Investor ID"].astype(str).str.strip().str.upper()
    investern["Invester Commitment"] = pd.to_numeric(investern["Invester Commitment"], errors="coerce")

    # Map SS Commitment using data from the current DataFormat set
    investern["SS Commitment"] = investern["Investor ID"].map(investorid_to_commitment)
    investern["SS Commitment"] = pd.to_numeric(investern["SS Commitment"], errors="coerce")
    investern["SS Check"] = investern["SS Commitment"] - investern["Invester Commitment"]

    # Debug missing IDs if any
    missing_ss = investern.loc[investern["SS Commitment"].isna(), "Investor ID"].unique()
    if len(missing_ss) > 0:
        print("⚠️ Warning: SS Commitment missing for Investor IDs:", list(missing_ss))

    # ---- Combine both sides ----
    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)})

    df = df.reindex(range(max_rows)).reset_index(drop=True)
    investern = investern.reindex(range(max_rows)).reset_index(drop=True)
    combined_df = pd.concat([df, spacer, investern], axis=1)

    # ---- Write to Excel ----
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        combined_df.to_excel(writer, sheet_name="Commitment Sheet", index=False)

    print("✅ Commitment Sheet created successfully.")
    return combined_df


# ---------------------------------------------------------
# Step 2: Create Entry Sheet (reuse in-memory commitment data)
# ---------------------------------------------------------
def create_entry_sheet_with_subtotals(commitment_df):
    delete_sheet_if_exists(output_file, "Entry")

    df_raw = pd.read_excel(wizard_file, sheet_name="allocation_data", engine="openpyxl", header=None)
    header_rows = df_raw.index[df_raw.iloc[:, 0] == "Vehicle/Investor"].tolist()
    tables = []

    for i, header_idx in enumerate(header_rows):
        start_idx = header_idx
        end_idx = header_rows[i + 1] if i + 1 < len(header_rows) else len(df_raw)
        block = df_raw.iloc[start_idx:end_idx].reset_index(drop=True)
        block.columns = block.iloc[0]
        block = block.drop(0).reset_index(drop=True)

        numeric_cols = ["Final LE Amount"]
        subtotal = block[numeric_cols].sum()
        subtotal_row = {
            col: subtotal[col] if col in numeric_cols else ("Subtotal" if col == block.columns[0] else "")
            for col in block.columns
        }
        block = pd.concat([block, pd.DataFrame([subtotal_row])], ignore_index=True)
        tables.append(block)

    processed = []
    for df in tables:
        if set(df.iloc[0]) == set(df.columns):
            df = df.iloc[1:]
        processed.append(df)

    final_df = pd.concat(processed, ignore_index=True)

    # Map using Commitment DF
    mapping = commitment_df[
        ["Investran Acct ID", "Bin ID", "Commitment Amount"]
    ].drop_duplicates(subset=["Investran Acct ID"])
    id_to_bin = mapping.set_index("Investran Acct ID")["Bin ID"].to_dict()
    id_to_commit = mapping.set_index("Investran Acct ID")["Commitment Amount"].to_dict()

    final_df["Bin ID"] = final_df["Investor ID"].map(id_to_bin)
    final_df["Commitment Amount"] = final_df["Investor ID"].map(id_to_commit)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        final_df.to_excel(writer, sheet_name="Entry", index=False)

    print("✅ Entry Sheet created successfully.")


# ---------------------------------------------------------
# Main
# ---------------------------------------------------------
if __name__ == "__main__":
    commitment_df = create_commitment_sheet()
    create_entry_sheet_with_subtotals(commitment_df)
    print("🎯 Automation completed successfully!")
