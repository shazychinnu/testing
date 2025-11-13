# ---- improved cleaning helper ----
def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Replace NA/None/nan with empty string for Excel output.
    Return dataframe (does not coerce dtypes).
    """
    df = df.copy()
    # Use pandas' fillna with inplace copy and also replace common markers
    df = df.replace([pd.NA, None, float("nan"), "nan", "<NA>", "None", "NULL"], "")
    df = df.fillna("")
    return df

# ---- improved removal of empty / zero-only columns ----
def remove_empty_or_zero_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Drop columns where:
      - all non-subtotal rows are empty/blank/NA (after treating empty string as NA)
      - AND subtotal rows (if any) sum to 0 (treating numeric-like strings as numeric)
    Keeps the first (description) column always.
    """
    cleaned_df = df.copy()

    # treat empty strings as NA for checks
    temp = cleaned_df.replace("", pd.NA)

    # identify subtotal rows (first column contains "Subtotal", case-insensitive)
    first_col = cleaned_df.columns[0]
    subtotal_mask = cleaned_df[first_col].astype(str).str.strip().str.lower() == "subtotal"

    cols_to_drop = []
    for col in cleaned_df.columns:
        if col == first_col:
            continue

        col_series = temp[col]  # with empty strings -> pd.NA

        # non-subtotal rows
        normal = col_series[~subtotal_mask]
        subtotal_rows = col_series[subtotal_mask]

        # If all normal rows are NA/empty -> consider as empty
        all_normal_empty = normal.isna().all()

        # Convert subtotal rows to numeric where possible, treat NA as 0 for the sum test
        subtotal_numeric = pd.to_numeric(subtotal_rows, errors="coerce").fillna(0)

        # If subtotal sum is zero and all normal rows empty, drop this column
        if all_normal_empty and subtotal_numeric.sum() == 0:
            cols_to_drop.append(col)
            continue

        # Also, if there are no subtotal rows, and all normal rows are numeric NA/0/empty,
        # drop if all normal are NA
        if not subtotal_mask.any() and normal.isna().all():
            cols_to_drop.append(col)

    if cols_to_drop:
        cleaned_df = cleaned_df.drop(columns=cols_to_drop)

    return cleaned_df

# ---- final corrected Entry sheet function ----
def create_entry_sheet_with_subtotals(commitment_df):
    """
    Build Entry sheet blocks, map Bin ID and multiple CDR-derived columns,
    compute subtotals per block robustly, and drop empty/zero-only columns.
    """
    delete_sheet_if_exists(output_file, "Entry")

    # Read allocation data without headers
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
        raise KeyError(f"Neither 'Investor ID' nor 'Investor Id' found in columns: {list(df_raw.columns)}")

    # Normalize IDs
    df_raw["_id_norm"] = df_raw[id_col].apply(norm_key)

    # Prepare mapping sources
    cm = commitment_df.copy()
    cdr = cdr_summary.copy()

    # Normalize Investran Acct ID in commitment df
    cm["_inv_acct_norm"] = cm["Investran Acct ID"].apply(norm_key)

    # Build InvestorID -> Bin ID mapping
    id_to_bin = (
        cm.dropna(subset=["Bin ID"])
          .drop_duplicates(subset=["_inv_acct_norm"])
          .set_index("_inv_acct_norm")["Bin ID"]
          .to_dict()
    )

    # Validate CDR
    if "Investor Commitment" not in cdr.columns:
        raise KeyError("Column 'Investor Commitment' not found in CDR Summary By Investor")

    # Normalize Account Number and prepare account->commitment map
    cdr["Account Number"] = cdr["Account Number"].astype(str).str.strip().str.upper()
    cdr["_bin_id_form"] = cdr["Account Number"].apply(norm_key)
    id_to_amt = cdr.groupby("_bin_id_form")["Investor Commitment"].sum().to_dict()

    # Map Bin ID and (only if exists) Commitment Amount
    df_raw["Bin ID"] = df_raw["_id_norm"].map(id_to_bin)
    df_raw["Commitment Amount"] = df_raw["Bin ID"].apply(
        lambda x: id_to_amt.get(norm_key(x), "")
        if pd.notna(x) and str(x).strip() != ""
        else ""
    )

    # Clean CDR summary headers then prefix section columns
    cdr_summary.columns = clean_headers(cdr_summary.columns)

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
            if "total" not in col_clean:
                new_col = f"{section}_{col}"
                section_cols.append(new_col)
                new_columns.append(new_col)
            else:
                new_columns.append(f"{section}_{col}")
        else:
            new_columns.append(col)
    cdr_summary.columns = new_columns

    # Prepare indexed CDR for mapping
    cdr_bin_col = find_col(cdr_summary, "Account Number")
    if not cdr_bin_col:
        raise KeyError("Could not find 'Account Number' in CDR Summary By Investor")
    # Normalize account numbers in index
    cdr_summary[cdr_bin_col] = cdr_summary[cdr_bin_col].astype(str).apply(norm_key)
    if not cdr_summary[cdr_bin_col].is_unique:
        cdr_summary = cdr_summary.drop_duplicates(subset=[cdr_bin_col])
    cdr_summary_indexed = cdr_summary.set_index(cdr_bin_col)

    # Process each block independently
    for i, h in enumerate(header_rows):
        start = h
        end = header_rows[i + 1] if i + 1 < len(header_rows) else len(df_raw)
        block = df_raw.loc[start:end].reset_index(drop=True)

        # drop block header inside block and keep global headers
        block = block.drop(0).reset_index(drop=True)
        block.columns = df_raw.columns

        entry_bin_col = find_col(block, "Bin ID")

        # Map each section column only when Bin ID present
        for col in section_cols:
            if entry_bin_col and col in cdr_summary_indexed.columns:
                # map on normalized bin values; leave empty if bin missing
                block[col] = block[entry_bin_col].apply(
                    lambda x:
                        cdr_summary_indexed[col].get(norm_key(x), "")
                        if pd.notna(x) and str(x).strip() != ""
                        else ""
                )
            else:
                block[col] = ""

        # Determine numeric-like columns by coercion (so "0" strings count)
        numeric_like = []
        for col in block.columns:
            # skip description column (first) if non-numeric
            sample = block[col].replace("", pd.NA)
            # if any non-subtotal normal row can be numeric, mark as numeric-like
            try_numeric = pd.to_numeric(sample[~(sample.astype(str).str.strip().str.lower() == "subtotal")], errors="coerce")
            if try_numeric.notna().any():
                numeric_like.append(col)

        # Build subtotal row using numeric coercion
        subtotal_row = {}
        for col in block.columns:
            if col in numeric_like:
                # coerce all values to numeric and sum (ignore non-numeric)
                val_sum = pd.to_numeric(block[col], errors="coerce").sum(skipna=True)
                subtotal_row[col] = val_sum if not pd.isna(val_sum) else ""
            else:
                subtotal_row[col] = ""

        # Label first column
        subtotal_row[block.columns[0]] = "Subtotal"

        # Append subtotal row
        block = pd.concat([block, pd.DataFrame([subtotal_row], columns=block.columns)], ignore_index=True)

        tables.append(block)

    # Combine blocks
    final_df = pd.concat(tables, ignore_index=True) if tables else pd.DataFrame()

    # Clean and remove empty/zero columns
    final_df = clean_dataframe(final_df)
    final_df = remove_empty_or_zero_columns(final_df)

    # Write final Entry sheet
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        final_df.to_excel(writer, sheet_name="Entry", index=False)
