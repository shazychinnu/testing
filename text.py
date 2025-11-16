def create_commitment_sheet():
    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # ---- 1. Load CDR Summary By Investor ----
    cdr = pd.read_excel(
        cdr_file,
        sheet_name="CDR Summary By Investor",
        engine="openpyxl",
        skiprows=2,
        dtype={"Account Number": str, "Investor ID": str}
    )

    cdr.columns = cdr.columns.str.strip()

    # Clean & normalize CDR keys
    cdr["Account Number"] = cdr["Account Number"].astype(str).str.strip()
    cdr["Investor ID"] = cdr["Investor ID"].astype(str).str.strip()

    cdr["Account Number"] = cdr["Account Number"].str.replace(r"[^A-Za-z0-9]", "", regex=True)
    cdr["_bin_norm"] = cdr["Account Number"].apply(norm_key)

    cdr["_investor_norm"] = cdr["Investor ID"].astype(str).apply(norm_key)

    cdr["Investor Commitment"] = pd.to_numeric(
        cdr["Investor Commitment"], errors="coerce"
    ).fillna(0)

    # ---------------------------------------------------------
    # PRIMARY MATCH: BUILD MULTI-KEY LOOKUP
    # (BIN ID + INVESTOR ID) → Commitment
    # ---------------------------------------------------------
    multi_key_commit = {
        (row["_bin_norm"], row["_investor_norm"]): row["Investor Commitment"]
        for _, row in cdr.iterrows()
    }

    # ---------------------------------------------------------
    # SECONDARY MATCH: FALLBACK ON BIN ID ONLY
    # ---------------------------------------------------------
    bin_only_commit = {}
    for _, row in cdr.iterrows():
        key = row["_bin_norm"]
        if key not in bin_only_commit:
            bin_only_commit[key] = row["Investor Commitment"]

    # ---- 2. Load Data_format ----
    df = pd.read_excel(
        wizard_file,
        sheet_name="Data_format",
        engine="openpyxl",
        dtype={"Investran Acct ID": str, "Bin ID": str, "Legal Entity": str}
    )

    df.columns = df.columns.str.strip()

    df["Legal Entity"] = df["Legal Entity"].astype(str).str.strip()
    df["Commitment Amount"] = pd.to_numeric(df["Commitment Amount"], errors="coerce").fillna(0)

    # Normalize Data_format keys
    df["Bin ID"] = df["Bin ID"].astype(str).str.strip().str.replace(r"[^A-Za-z0-9]", "", regex=True)
    df["Investran Acct ID"] = df["Investran Acct ID"].astype(str).str.strip()

    df["_bin_norm"] = df["Bin ID"].apply(norm_key)
    df["_inv_acct_norm"] = df["Investran Acct ID"].apply(norm_key)

    # ---------------------------------------------------------
    # FINAL GS COMMITMENT LOOKUP LOGIC
    # ---------------------------------------------------------
    def lookup_gs(row):
        key1 = (row["_bin_norm"], row["_inv_acct_norm"])  # best match
        if key1 in multi_key_commit:
            return multi_key_commit[key1]

        key2 = row["_bin_norm"]  # fallback match
        return bin_only_commit.get(key2, 0)

    df["GS Commitment"] = df.apply(lookup_gs, axis=1)

    # Replace Commitment Amount if needed
    df.loc[
        (df["Commitment Amount"] == 0) & (df["GS Commitment"] != 0),
        "Commitment Amount"
    ] = df["GS Commitment"]

    # ---- Subtotals (unchanged) ----
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

    # ---- 4. SS Commitment (unchanged) ----
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

    # ---- 5. Combine DataFrames ----
    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)}, dtype=object)

    df = df.reindex(range(max_rows)).reset_index(drop=True)
    investern = investern.reindex(range(max_rows)).reset_index(drop=True)

    combined_df = pd.concat([df.astype(object), spacer, investern.astype(object)], axis=1)

    # ---- 6. SS Subtotal ----
    subtotal_row = {col: "" for col in combined_df.columns}
    subtotal_row.update({
        "Vehicle/Investor": "Subtotal (SS Total)",
        "Invester Commitment": investern["Invester Commitment"].sum(),
        "SS Commitment": investern["SS Commitment"].sum(),
        "SS Check": investern["SS Commitment"].sum() - investern["Invester Commitment"].sum()
    })

    combined_df = pd.concat([combined_df, pd.DataFrame([subtotal_row], dtype=object)], ignore_index=True)

    # ---- 7. Remove internal columns ----
    internal_cols = [c for c in combined_df.columns if c.startswith("_")]
    combined_df.drop(columns=internal_cols, inplace=True, errors="ignore")

    # ---- 8. Clean output ----
    combined_df = clean_dataframe(combined_df)

    # ---- 9. Write to Excel ----
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        combined_df.to_excel(writer, sheet_name="Commitment Sheet", index=False)

    print("Commitment Sheet created successfully.")
    return combined_df
