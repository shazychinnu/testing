def create_entry_sheet_with_subtotals(commitment_df):
    delete_sheet_if_exists(output_file, "Entry")

    # ---- Read allocation data without headers ----
    df_raw = pd.read_excel(
        wizard_file,
        sheet_name="allocation_data",
        engine="openpyxl",
        header=None
    )

    # Identify header rows
    header_rows = df_raw.index[
        df_raw.iloc[:, 0].astype(str) == "Vehicle/Investor"
    ].tolist()
    tables = []

    # Read CDR Summary
    cdr_summary = pd.read_excel(
        cdr_file,
        sheet_name="CDR Summary By Investor",
        engine="openpyxl",
        skiprows=2
    )

    # Assign and clean headers globally
    if header_rows:
        header_values = list(df_raw.iloc[header_rows[0]])
        header_values = clean_headers(header_values)
        df_raw.columns = header_values
        df_raw = df_raw.drop(header_rows[0]).reset_index(drop=True)

    # Determine Investor ID column
    id_col = find_col(df_raw, "Investor ID") or find_col(df_raw, "Investor Id")
    if not id_col:
        raise KeyError(f"No Investor ID column in allocation_data")

    # Normalize IDs
    df_raw["_id_norm"] = df_raw[id_col].apply(norm_key)

    # Prepare mapping tables
    cm = commitment_df.copy()
    cdr = cdr_summary.copy()

    cm["_inv_acct_norm"] = cm["Investran Acct ID"].apply(norm_key)

    id_to_bin = (
        cm.dropna(subset=["Bin ID"])
          .drop_duplicates(subset=["_inv_acct_norm"])
          .set_index("_inv_acct_norm")["Bin ID"]
          .to_dict()
    )

    cdr["Account Number"] = cdr["Account Number"].astype(str).str.upper()
    cdr["_bin_id_form"] = cdr["Account Number"].apply(norm_key)
    id_to_amt = cdr.groupby("_bin_id_form")["Investor Commitment"].sum().to_dict()

    # Assign BIN and Commitment Amount
    df_raw["Bin ID"] = df_raw["_id_norm"].map(id_to_bin)
    df_raw["Commitment Amount"] = df_raw["Bin ID"].apply(
        lambda x: id_to_amt.get(norm_key(x), "") if pd.notna(x) and str(x).strip() else ""
    )

    # Clean CDR Summary headers
    cdr_summary.columns = clean_headers(cdr_summary.columns)

    # ---- Build section columns ----
    new_columns = []
    section_cols = []
    section = None

    for col in cdr_summary.columns:
        col_clean = str(col).strip().lower()
        if col_clean == "total contributions to commitment":
            section = "Contributions"
        elif col_clean == "total recallable":
            section = "Distributions"
        elif col_clean == "external expenses":
            section = "ExternalExpenses"

        if section and col not in ["Investor ID", "Account Number", "Investor Name", "Bin ID"]:
            new_col = f"{section}_{col}"
            section_cols.append(new_col)
            new_columns.append(new_col)
        else:
            new_columns.append(col)

    cdr_summary.columns = new_columns

    # Normalize CDR index
    cdr_bin_col = find_col(cdr_summary, "Account Number")
    cdr_summary[cdr_bin_col] = cdr_summary[cdr_bin_col].astype(str).apply(norm_key)
    cdr_summary = cdr_summary.drop_duplicates(subset=[cdr_bin_col])
    cdr_summary_indexed = cdr_summary.set_index(cdr_bin_col)

    # ---- Process each block ----
    for i, h in enumerate(header_rows):
        start = h
        end = header_rows[i + 1] if i + 1 < len(header_rows) else len(df_raw)

        block = df_raw.loc[start:end].reset_index(drop=True)
        block = block.drop(0).reset_index(drop=True)
        block.columns = df_raw.columns

        entry_bin_col = find_col(block, "Bin ID")

        # Map section columns
        for col in section_cols:
            if entry_bin_col and col in cdr_summary_indexed.columns:
                block[col] = block[entry_bin_col].apply(
                    lambda x: cdr_summary_indexed[col].get(norm_key(x), "")
                    if pd.notna(x) and str(x).strip() else ""
                )
            else:
                block[col] = ""

        # ---- Determine numeric-like columns ----
        numeric_like = []
        for col in block.columns:
            sample = block[col].replace("", pd.NA)
            try_num = pd.to_numeric(sample, errors="coerce")
            if try_num.notna().any():
                numeric_like.append(col)

        # ---- Build subtotal row ----
        subtotal_row = {col: "" for col in block.columns}

        for col in numeric_like:
            if col in ["Investor ID", "_id_norm"]:
                continue  # Do NOT subtotal these columns
            subtotal_row[col] = pd.to_numeric(block[col], errors="coerce").sum(skipna=True)

        subtotal_row[block.columns[0]] = "Subtotal"

        # Append subtotal
        block = pd.concat([block, pd.DataFrame([subtotal_row], columns=block.columns)], ignore_index=True)

        tables.append(block)

    # ---- Combine blocks ----
    final_df = pd.concat(tables, ignore_index=True) if tables else pd.DataFrame()

    # Clean and remove empty/zero columns
    final_df = clean_dataframe(final_df)
    final_df = remove_empty_or_zero_columns(final_df)

    # ---- Write Entry sheet ----
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        final_df.to_excel(writer, sheet_name="Entry", index=False)
