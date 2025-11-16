def create_commitment_sheet():
    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # ---- 1. Load CDR Summary By Investor
    cdr = cdr_file_data.copy()
    cdr.columns = cdr.columns.str.strip()

    # Normalize account, investor name, and ID
    cdr["Account Number Norm"] = cdr["Account Number"].apply(norm_key)
    cdr["Investor Name Norm"] = cdr["Investor Name"].astype(str).str.strip().str.upper()

    # Investor ID column may vary, normalize safely
    if "Investor ID" in cdr.columns:
        cdr["Investor ID Norm"] = cdr["Investor ID"].astype(str).apply(norm_key)
    else:
        cdr["Investor ID Norm"] = ""

    cdr["Investor Commitment"] = pd.to_numeric(cdr["Investor Commitment"], errors="coerce").fillna(0)

    # Create multi-key mapping for duplicate BIN IDs:
    # (Bin ID, Investor Name, Investor ID)
    acct_to_commit_multi = (
        cdr.groupby(["Account Number Norm", "Investor Name Norm", "Investor ID Norm"])["Investor Commitment"]
           .sum()
           .to_dict()
    )

    # Old single-key mapping (fallback)
    acct_to_commit = cdr.set_index("Account Number Norm")["Investor Commitment"].to_dict()

    # ---- 2. Load Data_format
    df = pd.read_excel(wizard_file, sheet_name="Data_format", engine="openpyxl")
    df.columns = df.columns.str.strip()

    df["Legal Entity"] = df["Legal Entity"].astype(str).str.strip()
    df["Legal Entity Norm"] = df["Legal Entity"].str.upper()

    df["Investran Acct ID"] = df["Investran Acct ID"].astype(str)
    df["Inv Acct Norm"] = df["Investran Acct ID"].apply(norm_key)

    df["Commitment Amount"] = pd.to_numeric(df["Commitment Amount"], errors="coerce").fillna(0)
    df["_bin_norm"] = df["Bin ID"].apply(norm_key)

    # ---- 3. Improved GS Commitment Logic (New)
    def get_gs_commit(row):
        key_multi = (row["_bin_norm"], row["Legal Entity Norm"], row["Inv Acct Norm"])

        # First try multi-key match
        if key_multi in acct_to_commit_multi:
            return acct_to_commit_multi[key_multi]

        # Fallback: try mapping by Bin ID only
        return acct_to_commit.get(row["_bin_norm"], 0)

    df["GS Commitment"] = df.apply(get_gs_commit, axis=1)
    df["GS Commitment"] = pd.to_numeric(df["GS Commitment"], errors="coerce").fillna(0)

    # Replace commitment amount where missing
    df.loc[(df["Commitment Amount"] == 0) & (df["GS Commitment"] != 0), 
           "Commitment Amount"] = df["GS Commitment"]

    # ---- Subtotal logic
    subtotal_mask = df["Legal Entity"].str.contains("Subtotal", case=False, na=False)
    subtotal_indices = df.index[subtotal_mask].to_list()

    start_idx = 0
    for idx in subtotal_indices:
        section = df.iloc[start_idx:idx]
        total_commit = section["Commitment Amount"].sum()
        total_gs_commit = section["GS Commitment"].sum()

        df.at[idx, "Commitment Amount"] = total_commit
        df.at[idx, "GS Commitment"] = total_gs_commit
        df.at[idx, "GS Check"] = total_commit - total_gs_commit

        start_idx = idx + 1

    df["GS Check"] = df["Commitment Amount"] - df["GS Commitment"]

    # ---- 4. SS Commitment (unchanged)
    ss_source = (
        df.loc[df["_bin_norm"].ne("") & df["_bin_norm"].notna()]
          .groupby("_bin_norm")["Commitment Amount"]
          .sum()
          .to_dict()
    )

    investern = pd.read_excel(wizard_file, sheet_name="investern_format", engine="openpyxl")
    investern.columns = investern.columns.str.strip()
    investern["Account Number"] = investern["Account Number"].astype(str).str.strip().str.upper()
    investern["Account Number"] = investern["Account Number"].replace(
        ["NAN", "NONE", "NULL", "<NA>", "NA", "N/A", "PD.NA"], value=""
    )
    investern["Account Number"] = investern["Account Number"].where(investern["Account Number"] != "nan", "")
    investern["_id_norm"] = investern["Account Number"].apply(lambda x: norm_key(x) if x != "" else "")

    investern["Invester Commitment"] = pd.to_numeric(investern["Invester Commitment"], errors="coerce").fillna(0)
    investern["SS Commitment"] = investern["_id_norm"].map(ss_source)
    investern["SS Commitment"] = pd.to_numeric(investern["SS Commitment"], errors="coerce").fillna(0)
    investern["SS Check"] = investern["SS Commitment"] - investern["Invester Commitment"]

    # ---- 5. Combine DataFrames
    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)}, dtype=object)

    df = df.reindex(range(max_rows)).reset_index(drop=True)
    investern = investern.reindex(range(max_rows)).reset_index(drop=True)

    combined_df = pd.concat([df.astype(object), spacer, investern.astype(object)], axis=1)

    # ---- 6. Add SS Subtotal Row
    ss_total_commit = pd.to_numeric(investern["SS Commitment"], errors="coerce").fillna(0).sum()
    ss_total_invest = pd.to_numeric(investern["Invester Commitment"], errors="coerce").fillna(0).sum()
    ss_total_check = ss_total_commit - ss_total_invest

    subtotal_row = {col: "" for col in combined_df.columns}
    subtotal_row.update({
        "Vehicle/Investor": "Subtotal (SS Total)",
        "Investor ID": "",
        "Invester Commitment": ss_total_invest,
        "SS Commitment": ss_total_commit,
        "SS Check": ss_total_check,
    })

    combined_df = pd.concat([combined_df, pd.DataFrame([subtotal_row], dtype=object)], ignore_index=True)

    # ---- 7. Remove helper columns
    internal_cols = [c for c in combined_df.columns if c.startswith("_")]
    combined_df.drop(columns=internal_cols, inplace=True, errors="ignore")

    # ---- 8. Final cleanup
    combined_df = clean_dataframe(combined_df)

    # ---- 9. Write clean sheet
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        combined_df.to_excel(writer, sheet_name="Commitment Sheet", index=False)

    print("Commitment Sheet created successfully.")
    return combined_df
