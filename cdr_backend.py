import pandas as pd
from openpyxl import load_workbook, Workbook

cdr_file = "CDR_VREP.xlsx"
wizard_file = "report_file.xlsx"
output_file = "output.xlsx"


def ensure_output_file_exists():
    """Ensure that the output Excel file exists."""
    try:
        wb = load_workbook(output_file)
    except FileNotFoundError:
        wb = Workbook()
    wb.save(output_file)


def delete_sheet_if_exists(workbook_path, sheet_name):
    """Delete sheet from workbook if it already exists."""
    wb = load_workbook(workbook_path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
        wb.save(workbook_path)
    wb.close()


def create_commitment_sheet():
    """Create or replace the Commitment Sheet and return it as a DataFrame."""
    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # ------------------ CDR Summary ------------------
    cdr_summary = pd.read_excel(
        cdr_file,
        sheet_name="CDR Summary By Investor",
        engine="openpyxl",
        skiprows=2
    )
    cdr_summary.columns = cdr_summary.columns.str.strip()
    cdr_summary["Account Number"] = cdr_summary["Account Number"].astype(str).str.strip()
    cdr_summary["Investor ID"] = cdr_summary["Investor ID"].astype(str).str.strip().str.upper()

    account_to_commitment = cdr_summary.set_index("Account Number")["Investor Commitment"].to_dict()
    investorid_to_commitment = cdr_summary.set_index("Investor ID")["Investor Commitment"].to_dict()

    # ------------------ Data_format ------------------
    df = pd.read_excel(wizard_file, sheet_name="Data_format", engine="openpyxl")
    df.columns = df.columns.str.strip()
    df["Bin ID"] = df["Bin ID"].astype(str).str.strip()
    df["Legal Entity"] = df["Legal Entity"].astype(str).str.strip()
    df["Commitment Amount"] = pd.to_numeric(df["Commitment Amount"], errors="coerce")

    # Identify subtotal rows
    subtotal_mask = df["Legal Entity"].str.contains("Subtotal", case=False, na=False)

    # Map GS Commitment only for non-subtotal rows
    df["GS Commitment"] = pd.NA
    df.loc[~subtotal_mask, "GS Commitment"] = df.loc[~subtotal_mask, "Bin ID"].map(account_to_commitment)

    # Calculate subtotals
    subtotal_indices = df.index[subtotal_mask].tolist() + [len(df)]
    start_idx = 0
    for idx in subtotal_indices[:-1]:
        block = df.iloc[start_idx:idx]
        df.at[idx, "Commitment Amount"] = block["Commitment Amount"].sum()
        df.at[idx, "GS Commitment"] = block["GS Commitment"].astype(float).sum(skipna=True)
        start_idx = idx + 1

    # Re-blank out subtotal rows for GS Commitment per your rule
    df.loc[subtotal_mask, "GS Commitment"] = pd.NA
    df["GS Check"] = df["Commitment Amount"] - df["GS Commitment"]

    # ------------------ investern_format ------------------
    investern = pd.read_excel(wizard_file, sheet_name="investern_format", engine="openpyxl")
    investern.columns = investern.columns.str.strip()

    # Normalize IDs
    investern["Investor ID"] = investern["Investor ID"].astype(str).str.strip().str.upper()
    investern["Invester Commitment"] = pd.to_numeric(investern["Invester Commitment"], errors="coerce")

    # Map SS Commitment
    investern["SS Commitment"] = investern["Investor ID"].map(investorid_to_commitment)
    investern["SS Commitment"] = pd.to_numeric(investern["SS Commitment"], errors="coerce")
    investern["SS Check"] = investern["SS Commitment"] - investern["Invester Commitment"]

    # Debug info for missing mappings
    missing_ids = investern.loc[investern["SS Commitment"].isna(), "Investor ID"].unique()
    if len(missing_ids) > 0:
        print(f"⚠️ Warning: SS Commitment not found for Investor IDs: {list(missing_ids)}")

    # ------------------ Combine Both ------------------
    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)})

    df = df.reindex(range(max_rows)).reset_index(drop=True)
    investern = investern.reindex(range(max_rows)).reset_index(drop=True)
    spacer = spacer.reindex(range(max_rows)).reset_index(drop=True)

    combined_df = pd.concat([df, spacer, investern], axis=1)

    # Write Commitment Sheet
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        combined_df.to_excel(writer, sheet_name="Commitment Sheet", index=False)

    print("✅ Commitment Sheet created (replaced if existed).")
    return combined_df


def create_entry_sheet_with_subtotals(commitment_df):
    """Create the Entry sheet using in-memory Commitment data."""
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
        subtotal_row = {col: subtotal[col] if col in numeric_cols else ("Subtotal" if col == block.columns[0] else "")
                        for col in block.columns}
        block = pd.concat([block, pd.DataFrame([subtotal_row])], ignore_index=True)
        tables.append(block)

    processed = []
    for df in tables:
        if set(df.iloc[0]) == set(df.columns):
            df = df.iloc[1:]
        processed.append(df)

    final_df = pd.concat(processed, ignore_index=True)

    # Map from in-memory Commitment DF
    mapping = commitment_df[["Investran Acct ID", "Bin ID", "Commitment Amount"]].drop_duplicates(subset=["Investran Acct ID"])
    id_to_bin = mapping.set_index("Investran Acct ID")["Bin ID"].to_dict()
    id_to_commit = mapping.set_index("Investran Acct ID")["Commitment Amount"].to_dict()

    final_df["Bin ID"] = final_df["Investor ID"].map(id_to_bin)
    final_df["Commitment Amount"] = final_df["Investor ID"].map(id_to_commit)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        final_df.to_excel(writer, sheet_name="Entry", index=False)

    print("✅ Entry Sheet created (replaced if existed).")


if __name__ == "__main__":
    commitment_df = create_commitment_sheet()
    create_entry_sheet_with_subtotals(commitment_df)
    print("🎯 Automation completed successfully!")
