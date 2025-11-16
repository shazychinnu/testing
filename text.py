def create_commitment_sheet():
    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # ---- 1. Load CDR Summary ----
    cdr = cdr_file_data.copy()
    cdr.columns = cdr.columns.str.strip()
    cdr["Account Number"] = cdr["Account Number"].astype(str).str.strip()
    cdr["Investor ID"] = cdr["Investor ID"].astype(str).str.strip()

    cdr["_bin_norm"] = cdr["Account Number"].apply(norm_key)
    cdr["_investor_norm"] = cdr["Investor ID"].apply(norm_key)

    cdr["Investor Commitment"] = pd.to_numeric(
        cdr["Investor Commitment"], errors="coerce"
    ).fillna(0)

    multi_key_commit = {
        (row["_bin_norm"], row["_investor_norm"]): row["Investor Commitment"]
        for _, row in cdr.iterrows()
    }

    bin_only_commit = {}
    for _, row in cdr.iterrows():
        if row["_bin_norm"] not in bin_only_commit:
            bin_only_commit[row["_bin_norm"]] = row["Investor Commitment"]

    # ---- 2. Load Data_format ----
    df = pd.read_excel(
        wizard_file,
        sheet_name="Data_format",
        engine="openpyxl",
        dtype={"Bin ID": str, "Investran Acct ID": str, "Legal Entity": str}
    )
    df.columns = df.columns.str.strip()

    df["Legal Entity"] = df["Legal Entity"].astype(str).fillna("").str.strip()
    df["Bin ID"] = df["Bin ID"].astype(str).fillna("").str.strip()
    df["Investran Acct ID"] = df["Investran Acct ID"].astype(str).fillna("").str.strip()

    df["Commitment Amount"] = pd.to_numeric(df["Commitment Amount"], errors="coerce").fillna(0)

    df["_bin_norm"] = df["Bin ID"].apply(norm_key)
    df["_inv_acct_norm"] = df["Investran Acct ID"].apply(norm_key)

    # ---- 3. GS COMMITMENT from CDR ----
    def lookup_gs(row):
        key = (row["_bin_norm"], row["_inv_acct_norm"])
        if key in multi_key_commit:
            return multi_key_commit[key]
        return bin_only_commit.get(row["_bin_norm"], 0)

    df["GS Commitment"] = df.apply(lookup_gs, axis=1)
    df["GS Check"] = df["Commitment Amount"] - df["GS Commitment"]

    # --------------------------------------------------------------------
    # ⭐ STEP A — SMART SECTION DETECTION (WORKS EVEN IF FEEDER ANYWHERE)
    # --------------------------------------------------------------------
    subtotal_mask = df["Legal Entity"].str.upper().str.contains("SUBTOTAL", na=False)
    subtotal_indices = list(df.index[subtotal_mask])

    section_id = []
    current_section = -1

    for i in range(len(df)):
        if i in subtotal_indices:
            section_id.append(current_section)
        else:
            if i == 0 or (i - 1) in subtotal_indices:
                current_section += 1
            section_id.append(current_section)

    df["SectionID"] = section_id

    # --------------------------------------------------------------------
    # ⭐ STEP B — GET GS SUBTOTAL FOR EACH SECTION
    # --------------------------------------------------------------------
    section_totals = {}
    for s in df["SectionID"].unique():
        sec_df = df[df["SectionID"] == s]
        subtotal_row = sec_df[sec_df["Legal Entity"].str.upper().str.contains("SUBTOTAL")]

        if not subtotal_row.empty:
            gs_total = float(subtotal_row["GS Commitment"].iloc[0] or 0)
            section_totals[s] = gs_total
        else:
            section_totals[s] = 0

    # --------------------------------------------------------------------
    # ⭐ STEP C — APPLY FEEDER FIX (FINAL DF DIRECT UPDATE)
    # --------------------------------------------------------------------
    feeder_mask = df["Bin ID"].str.upper().str.contains("FEEDER", na=False)
    feeder_indices = df[feeder_mask].index

    for idx in feeder_indices:
        sec = df.loc[idx, "SectionID"]
        subtotal_gs = section_totals.get(sec, 0)

        df.loc[idx, "GS Commitment"] = subtotal_gs
        df.loc[idx, "GS Check"] = df.loc[idx, "Commitment Amount"] - subtotal_gs

    # --------------------------------------------------------------------
    # ⭐ NOW FEEDER GS IS 100% CORRECT  
    # NOTHING CAN OVERWRITE IT ANYMORE
    # --------------------------------------------------------------------

    # ---- 4. SS Commitment ----
    ss_source = df.groupby("_bin_norm")["Commitment Amount"].sum().to_dict()

    investern = pd.read_excel(
        wizard_file,
        sheet_name="investern_format",
        engine="openpyxl",
        dtype={"Account Number": str}
    )
    investern.columns = investern.columns.str.strip()

    investern["Account Number"] = investern["Account Number"].astype(str).str.upper().str.strip()
    investern["_id_norm"] = investern["Account Number"].apply(norm_key)

    investern["Invester Commitment"] = pd.to_numeric(
        investern["Invester Commitment"], errors="coerce"
    ).fillna(0)

    investern["SS Commitment"] = investern["_id_norm"].map(ss_source).fillna(0)
    investern["SS Check"] = investern["SS Commitment"] - investern["Invester Commitment"]

    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)}, dtype=object)

    df = df.reindex(range(max_rows)).reset_index(drop=True)
    investern = investern.reindex(range(max_rows)).reset_index(drop=True)

    combined_df = pd.concat([df.astype(object), spacer, investern.astype(object)], axis=1)

    subtotal_row = {col: "" for col in combined_df.columns}
    subtotal_row.update({
        "Vehicle/Investor": "Subtotal (SS Total)",
        "Invester Commitment": investern["Invester Commitment"].sum(),
        "SS Commitment": investern["SS Commitment"].sum(),
        "SS Check": investern["SS Commitment"].sum() - investern["Invester Commitment"].sum(),
    })
    combined_df = pd.concat([combined_df, pd.DataFrame([subtotal_row])], ignore_index=True)

    internal_cols = [c for c in combined_df.columns if c.startswith("_")]
    combined_df.drop(columns=internal_cols, inplace=True, errors="ignore")
    combined_df = clean_dataframe(combined_df)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        combined_df.to_excel(writer, sheet_name="Commitment Sheet", index=False)

    print("Commitment Sheet created successfully.")
    return combined_df
