import pandas as pd

def create_entry_sheet_with_subtotals(commitment_df):
    """
    Create an 'Entry' sheet in the output Excel file, combining data from wizard and commitment files,
    calculating subtotals per block, and joining commitment and BIN mapping data.
    """
    
    # --- Helper function to normalize keys ---
    def norm_key(x):
        s = str(x).strip()
        if s.endswith(".0):
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
    cdr["_bin_id_form"] = cdr["Account Number"].apply(norm_key)

    # --- Map normalized account ID to BIN ID ---
    id_to_bin = (
        cm.dropna(subset=["Bin ID"])
        .drop_duplicates(subset=["_inv_acct_norm"])
        .set_index("_inv_acct_norm")["Bin ID"]
        .to_dict()
    )

    cm["_inv_acct_norm"] = cm["_inv_acct_norm"].astype(str).fillna("").str.strip().str.upper()
    cm["Commitment Amount"] = pd.to_numeric(cm["Commitment Amount"], errors="coerce").fillna(0)

    final_df["Bin ID"] = final_df["_id_norm"].map(id_to_bin)

    cdr["_bin_id_form"] = cdr["_bin_id_form"].astype(str).fillna("").str.strip().str.upper()
    id_to_amt = cdr.groupby("_bin_id_form")["Investor Commitment"].sum().to_dict()
    final_df["Commitment Amount"] = final_df["Bin ID"].map(id_to_amt)

    # --- Remove helper columns and clean ---
    final_df.drop(columns=[c for c in final_df.columns if c.startswith("_")], inplace=True, errors="ignore")
    final_df = clean_dataframe(final_df)

    # --- Write to Excel ---
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        final_df.to_excel(writer, sheet_name="Entry", index=False)

    print("Entry Sheet created successfully - clean and validated.")
