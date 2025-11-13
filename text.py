def remove_empty_or_zero_columns(df):
    """
    Remove columns where ALL values are:
      - empty/blank/NA
      - OR zero/zero-like ("0","0.0", 0.0)
    Always remove _id_norm.
    """
    cleaned = df.copy()

    # Always remove helper column
    if "_id_norm" in cleaned.columns:
        cleaned = cleaned.drop(columns=["_id_norm"])

    cols_to_drop = []

    for col in cleaned.columns:
        if col == cleaned.columns[0]:  # never drop first column
            continue

        series = cleaned[col]

        # Treat blank as NA
        s = series.replace("", pd.NA)

        # Convert everything to numeric (invalid -> NaN)
        s_num = pd.to_numeric(s, errors="coerce")

        # All blank?
        all_empty = s.isna().all()

        # All zero? (0, 0.0, "0", "0.0")
        all_zero = s_num.fillna(0).eq(0).all()

        if all_empty or all_zero:
            cols_to_drop.append(col)

    return cleaned.drop(columns=cols_to_drop)


def create_entry_sheet_with_subtotals(commitment_df):
    delete_sheet_if_exists(output_file, "Entry")

    # ---- Read allocation data (no headers) ----
    df_raw = pd.read_excel(
        wizard_file, sheet_name="allocation_data",
        engine="openpyxl", header=None
    )

    # Identify header rows
    header_rows = df_raw.index[df_raw.iloc[:, 0].astype(str) == "Vehicle/Investor"].tolist()
    tables = []

    # Read CDR Summary By Investor
    cdr_summary = pd.read_excel(
        cdr_file,
        sheet_name="CDR Summary By Investor",
        engine="openpyxl",
        skiprows=2
    )

    # Assign global headers
    if header_rows:
        header_values = clean_headers(list(df_raw.iloc[header_rows[0]]))
        df_raw.columns = header_values
        df_raw = df_raw.drop(header_rows[0]).reset_index(drop=True)

    # Identify Investor ID column
    id_col = find_col(df_raw, "Investor ID") or find_col(df_raw, "Investor Id")
    if not id_col:
        raise KeyError("Investor ID column missing in allocation_data")

    # Normalize IDs
    df_raw["_id_norm"] = df_raw[id_col].apply(norm_key)

    # Prepare mapping sources
    cm = commitment_df.copy()
    cm["_inv_acct_norm"] = cm["Investran Acct ID"].apply(norm_key)

    # InvestorID → BIN ID
    id_to_bin = (
        cm.dropna(subset=["Bin ID"])
          .drop_duplicates(subset=["_inv_acct_norm"])
          .set_index("_inv_acct_norm")["Bin ID"]
          .to_dict()
    )

    # Prepare CDR Summary mapping
    cdr_summary["Account Number"] = cdr_summary["Account Number"].astype(str).str.upper()
    cdr_summary["_bin_id_form"] = cdr_summary["Account Number"].apply(norm_key)

    id_to_amt = cdr_summary.groupby("_bin_id_form")["Investor Commitment"].sum().to_dict()

    # Apply mapping to allocation_data
    df_raw["Bin ID"] = df_raw["_id_norm"].map(id_to_bin)
    df_raw["Commitment Amount"] = df_raw["Bin ID"].apply(
        lambda x: id_to_amt.get(norm_key(x), "") if pd.notna(x) and str(x).strip() else ""
    )

    # Clean CDR
    cdr_summary.columns = clean_headers(cdr_summary.columns)

    # Build section columns
    new_columns, section_cols = [], []
    section = None

    for col in cdr_summary.columns:
        col_clean = str(col).lower().strip()

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

    # Index CDR on normalized Account Number
    cdr_bin_col = find_col(cdr_summary, "Account Number")
    cdr_summary[cdr_bin_col] = cdr_summary[cdr_bin_col].astype(str).apply(norm_key)
    cdr_summary = cdr_summary.drop_duplicates(subset=[cdr_bin_col])
    cdr_indexed = cdr_summary.set_index(cdr_bin_col)

    # ---- PROCESS EACH BLOCK ----
    for i, h in enumerate(header_rows):
        start = h
        end = header_rows[i+1] if i+1 < len(header_rows) else len(df_raw)

        block = df_raw.loc[start:end].reset_index(drop=True)
        block = block.drop(0).reset_index(drop=True)
        block.columns = df_raw.columns

        entry_bin_col = find_col(block, "Bin ID")

        # Map section columns
        for col in section_cols:
            if entry_bin_col and col in cdr_indexed.columns:
                block[col] = block[entry_bin_col].apply(
                    lambda x: cdr_indexed[col].get(norm_key(x), "")
                    if pd.notna(x) and str(x).strip() else ""
                )
            else:
                block[col] = ""

        # Identify numeric-like columns
        numeric_like = []
        for col in block.columns:
            s = block[col].replace("", pd.NA)
            if pd.to_numeric(s, errors="coerce").notna().any():
                numeric_like.append(col)

        # Subtotal row
        subtotal_row = {col: "" for col in block.columns}

        for col in numeric_like:
            if col in ["Investor ID", "_id_norm"]:
                continue
            subtotal_row[col] = pd.to_numeric(block[col], errors="coerce").sum(skipna=True)

        subtotal_row[block.columns[0]] = "Subtotal"

        block = pd.concat([block, pd.DataFrame([subtotal_row])], ignore_index=True)
        tables.append(block)

    # Combine blocks
    final_df = pd.concat(tables, ignore_index=True)

    # Final cleanup
    final_df = clean_dataframe(final_df)
    final_df = remove_empty_or_zero_columns(final_df)

    # Write sheet
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        final_df.to_excel(writer, sheet_name="Entry", index=False)
