def create_commitment_sheet():
    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # ---- 1. Load CDR Summary (force proper dtypes)
    cdr = pd.read_excel(
        cdr_file,
        sheet_name="CDR Summary By Investor",
        engine="openpyxl",
        skiprows=2,
        dtype={"Account Number": str, "Investor ID": str}
    )

    cdr.columns = cdr.columns.str.strip()

    # Normalize for matching
    cdr["_bin_norm"] = cdr["Account Number"].apply(norm_key)      # BIN ID equivalent
    cdr["_inv_id_norm"] = cdr["Investor ID"].astype(str).apply(norm_key)

    cdr["Investor Commitment"] = pd.to_numeric(
        cdr["Investor Commitment"], errors="coerce"
    ).fillna(0)

    # ---------------------------------------------------------
    # TRUE FINAL FIX: BUILD MULTI-KEY LOOKUP WITHOUT SUM
    # (Bin ID + Investor ID) → Commitment
    # ---------------------------------------------------------
    acct_to_commit_multi = {}
    for _, row in cdr.iterrows():
        key = (row["_bin_norm"], row["_inv_id_norm"])
        # keep FIRST occurrence only (correct business logic)
        if key not in acct_to_commit_multi:
            acct_to_commit_multi[key] = row["Investor Commitment"]

    # ---- 2. Load Data_format with correct string dtype
    df = pd.read_excel(
        wizard_file,
        sheet_name="Data_format",
        engine="openpyxl",
        dtype={"Investran Acct ID": str, "Bin ID": str, "Legal Entity": str}
    )

    df.columns = df.columns.str.strip()
    df["Legal Entity"] = df["Legal Entity"].astype(str).str.strip()

    df["Commitment Amount"] = pd.to_numeric(
        df["Commitment Amount"], errors="coerce"
    ).fillna(0)

    # Normalize BIN + Investran Acct ID
    df["_bin_norm"] = df["Bin ID"].apply(norm_key)
    df["_inv_acct_norm"] = df["Investran Acct ID"].apply(norm_key)

    # ---------------------------------------------------------
    # CORRECT GS COMMITMENT LOOKUP (NO SUM, NO DUPLICATE ERROR)
    # ---------------------------------------------------------
    def lookup_gs(row):
        key = (row["_bin_norm"], row["_inv_acct_norm"])

        # FIRST: exact match on (BIN ID + Investor ID)
        if key in acct_to_commit_multi:
            return acct_to_commit_multi[key]

        return 0  # no fallback needed unless you tell me

    df["GS Commitment"] = df.apply(lookup_gs, axis=1)

    # Replace Commitment Amount where missing
    df.loc[
        (df["Commitment Amount"] == 0) & (df["GS Commitment"] != 0),
        "Commitment Amount"
    ] = df["GS Commitment"]

    # ------------------------------
    # SUBTOTAL LOGIC (unchanged)
    # ------------------------------
    subtotal_mask = df["Legal Entity"].str.contains("Subtotal", case=False, na=False)
    subtotal_indices = df.index[subtotal_mask].to_list()

    start_idx = 0
    for idx in subtotal_indices:
        section = df.iloc[start_idx:idx]
        df.at[idx, "Commitment Amount"] = section["Commitment Amount"].sum()
        df.at[idx, "GS Commitment"] = section["GS Commitment"].sum()
        df.at[idx, "GS Check"] = df.at[idx, "Commitment Amount"] - df.at[idx, "GS Commitment"]
        start_idx = idx + 1

    df["GS Check"] = df["Commitment Amount"] - df["GS Commitment"]

    # ---- 4. SS Commitment
    ss_source = (
        df.loc[df["_bin_norm"].ne("") & df["_bin_norm"].notna()]
          .groupby("_bin_norm")["Commitment Amount"]
          .sum()
          .to_dict()
    )

    investern = pd.read_excel(
        wizard_file,
        sheet_name="investern_format",
        engine="openpyxl",
        dtype={"Account Number": str}
    )

    investern.columns = investern.columns.str.strip()
    investern["Account Number"] = investern["Account Number"].astype(str).str.strip().str.upper()
    investern["_id_norm"] = investern["Account Number"].apply(norm_key)

    investern["Invester Commitment"] = pd.to_numeric(
        investern["Invester Commitment"], errors="coerce"
    ).fillna(0)

    investern["SS Commitment"] = investern["_id_norm"].map(ss_source).fillna(0)
    investern["SS Check"] = investern["SS Commitment"] - investern["Invester Commitment"]

    # ---- 5. Combine DataFrames
    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)}, dtype=object)

    df = df.reindex(range(max_rows)).reset_index(drop=True)
    investern = investern.reindex(range(max_rows)).reset_index(drop=True)

    combined_df = pd.concat([df.astype(object), spacer, investern.astype(object)], axis=1)

    # ---- 6. SS Subtotal Row
    subtotal_row = {col: "" for col in combined_df.columns}
    subtotal_row.update({
        "Vehicle/Investor": "Subtotal (SS Total)",
        "Invester Commitment": investern["Invester Commitment"].sum(),
        "SS Commitment": investern["SS Commitment"].sum(),
        "SS Check": investern["SS Commitment"].sum() - investern["Invester Commitment"].sum()
    })

    combined_df = pd.concat([combined_df, pd.DataFrame([subtotal_row], dtype=object)], ignore_index=True)

    # ---- 7. Remove helper columns
    internal_cols = [c for c in combined_df.columns if c.startswith("_")]
    combined_df.drop(columns=internal_cols, inplace=True, errors="ignore")

    # ---- 8. Final cleanup
    combined_df = clean_dataframe(combined_df)

    # ---- 9. Write final output
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        combined_df.to_excel(writer, sheet_name="Commitment Sheet", index=False)

    print("Commitment Sheet created successfully.")
    return combined_df
