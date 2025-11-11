import pandas as pd
import re
from openpyxl import load_workbook


def create_entry_sheet_with_subtotals(commitment_df):
    """
    Create an 'Entry' sheet in the output Excel file, combining data from wizard and commitment files,
    calculating subtotals per block, and joining commitment and BIN mapping data.
    Dynamic header handling is implemented for multi-level Excel headers.
    """

    # --- Helper function to normalize keys ---
    def norm_key(x):
        if pd.isna(x):
            return ""
        s = str(x).strip()
        if s.endswith(".0"):
            s = s[:-2]
        s = s.replace("\u00A0", " ")
        s = re.sub(r"[,\-]", "", s)
        return s.upper()

    # --- Delete existing sheet if it exists ---
    delete_sheet_if_exists(output_file, "Entry")

    # --- Load raw data from wizard file ---
    df_raw = pd.read_excel(wizard_file, sheet_name="allocation_data", engine="openpyxl", header=None)
    cdr = cdr_file_data.copy()

    # --- Identify table headers ---
    header_rows = df_raw.index[df_raw.iloc[:, 0].astype(str) == "Vehicle/Investor"].tolist()
    tables = []

    for i, h in enumerate(header_rows):
        start = h
        end = header_rows[i + 1] if i + 1 < len(header_rows) else len(df_raw)
        block = df_raw.iloc[start:end].reset_index(drop=True)

        # First row = column headers
        block.columns = block.iloc[0]
        block = block.drop(0).reset_index(drop=True)

        # --- Calculate subtotal if "Final LE Amount" column exists ---
        if "Final LE Amount" in block.columns:
            block["Final LE Amount"] = pd.to_numeric(block["Final LE Amount"], errors="coerce").fillna(0)
            subtotal_val = block["Final LE Amount"].sum(skipna=True)
        else:
            subtotal_val = 0

        # --- Add subtotal row ---
        subtotal_row = {col: "" for col in block.columns}
        subtotal_row[block.columns[0]] = "Subtotal"
        if "Final LE Amount" in block.columns:
            subtotal_row["Final LE Amount"] = subtotal_val

        block = pd.concat([block, pd.DataFrame([subtotal_row])], ignore_index=True)
        tables.append(block)

    # --- Combine all tables ---
    final_df = pd.concat(tables, ignore_index=True) if tables else pd.DataFrame()

    # --- Normalize ID columns ---
    id_col = "Investor ID" if "Investor ID" in final_df.columns else "Investor Id"
    final_df["_id_norm"] = final_df[id_col].apply(norm_key)

    cm = commitment_df.copy()
    cm["_inv_acct_norm"] = cm["Investran Acct ID"].apply(norm_key)

    # -------------------------------------------------------------------
    # 🧩 NEW LOGIC: Dynamic column handling for CDR Summary By Investor
    # -------------------------------------------------------------------
    cdr = pd.read_excel(
        cdr_file_data,
        sheet_name="CDR Summary By Investor",
        header=[0, 1],  # Two header rows
        engine="openpyxl"
    )

    # Flatten dynamic multi-level headers
    flattened_columns = []
    for lvl1, lvl2 in cdr.columns:
        if pd.isna(lvl2) or str(lvl2).startswith("Unnamed"):
            flat_name = str(lvl1).strip().replace(" ", "_")
        else:
            flat_name = f"{str(lvl1).strip()}_{str(lvl2).strip()}".replace(" ", "_")
        flattened_columns.append(flat_name)
    cdr.columns = flattened_columns

    # Normalize BIN ID column dynamically
    bin_cols = [c for c in cdr.columns if "BIN" in c.upper()]
    if not bin_cols:
        raise ValueError("No BIN ID column found in the CDR Summary sheet.")
    bin_col = bin_cols[0]
    cdr[bin_col] = cdr[bin_col].apply(norm_key)

    # --- Map normalized account ID to BIN ID ---
    cm["_inv_acct_norm"] = cm["_inv_acct_norm"].astype(str).fillna("").str.strip().str.upper()
    id_to_bin = (
        cm.dropna(subset=["Bin ID"])
        .drop_duplicates(subset=["_inv_acct_norm"])
        .set_index("_inv_acct_norm")["Bin ID"]
        .to_dict()
    )
    final_df["Bin ID"] = final_df["_id_norm"].map(id_to_bin)

    # --- Merge Commitment and CDR Data by BIN ID ---
    merged_df = pd.merge(final_df, cdr, left_on="Bin ID", right_on=bin_col, how="left")

    # --- Dynamic numeric column detection ---
    numeric_cols = merged_df.select_dtypes(include=["number"]).columns.tolist()
    merged_df[numeric_cols] = merged_df[numeric_cols].fillna(0)

    # --- Add a Grand Total row dynamically ---
    totals = merged_df[numeric_cols].sum(numeric_only=True)
    total_row = {col: "" for col in merged_df.columns}
    total_row[list(totals.index)] = totals
    total_row["Vehicle/Investor" if "Vehicle/Investor" in merged_df.columns else merged_df.columns[0]] = "TOTAL"
    merged_df = pd.concat([merged_df, pd.DataFrame([total_row])], ignore_index=True)

    # --- Final clean-up ---
    merged_df.drop(columns=[c for c in merged_df.columns if c.startswith("_")], inplace=True, errors="ignore")
    merged_df = clean_dataframe(merged_df)

    # --- Write to Excel ---
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        merged_df.to_excel(writer, sheet_name="Entry", index=False)

    print("✅ Entry Sheet created successfully with dynamic column handling.")
