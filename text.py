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

    cdr["Investor Commitment"] = pd.to_numeric(cdr["Investor Commitment"], errors="coerce").fillna(0)

    # Multi-key mapping (BIN + Investor ID)
    multi_key_commit = {
        (row["_bin_norm"], row["_investor_norm"]): row["Investor Commitment"]
        for _, row in cdr.iterrows()
    }

    # Fallback BIN-only mapping
    bin_only_commit = {}
    for _, row in cdr.iterrows():
        if row["_bin_norm"] not in bin_only_commit:
            bin_only_commit[row["_bin_norm"]] = row["Investor Commitment"]

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

    df["Bin ID"] = df["Bin ID"].astype(str).str.strip()
    df["Investran Acct ID"] = df["Investran Acct ID"].astype(str).str.strip()

    df["_bin_norm"] = df["Bin ID"].apply(norm_key)
    df["_inv_acct_norm"] = df["Investran Acct ID"].apply(norm_key)

    # ---- GS COMMITMENT LOOKUP ----
    def lookup_gs(row):
        key1 = (row["_bin_norm"], row["_inv_acct_norm"])
        if key1 in multi_key_commit:
            return multi_key_commit[key1]
        return bin_only_commit.get(row["_bin_norm"], 0)

    df["GS Commitment"] = df.apply(lookup_gs, axis=1)
    df["GS Check"] = df["Commitment Amount"] - df["GS Commitment"]

    # ---- 3. SECTION SPLITTING ----
    sections = []
    start = 0

    subtotal_mask = df["Legal Entity"].str.upper().str.contains("SUBTOTAL")

    for idx in df.index[subtotal_mask]:
        section = df.iloc[start:idx + 1].copy()   # include subtotal row
        sections.append(section)
        start = idx + 1

    # ---- 4. PROCESS EACH SECTION ----
    processed_sections = []

    for section in sections:

        # Identify subtotal row (last row)
        subtotal_row_idx = section.index[-1]

        # Extract initial subtotal GS Commitment
        initial_subtotal_gs = float(section.at[subtotal_row_idx, "GS Commitment"])

        # Identify FEEDER rows using substring match
        feeder_mask = section["Bin ID"].str.upper().str.contains("FEEDER")

        # Update FEEDER GS using initial subtotal GS
        section.loc[feeder_mask, "GS Commitment"] = initial_subtotal_gs

        # Recompute GS Check for FEEDER rows
        section.loc[feeder_mask, "GS Check"] = (
            section.loc[feeder_mask, "Commitment Amount"] - initial_subtotal_gs
        )

        # ---- Recalculate subtotal again (final subtotal) ----
        data_rows = section.iloc[:-1]  # exclude subtotal row

        final_commit = data_rows["Commitment Amount"].sum()
        final_gs = data_rows["GS Commitment"].sum()

        # Write final corrected subtotal
        section.at[subtotal_row_idx, "Commitment Amount"] = final_commit
        section.at[subtotal_row_idx, "GS Commitment"] = final_gs
        section.at[subtotal_row_idx, "GS Check"] = final_commit - final_gs

        processed_sections.append(section)

    # Combine all processed sections back
    df = pd.concat(processed_sections, ignore_index=True)

    # ---- 5. SS Commitment (unchanged logic) ----
    ss_source = (
        df[df["_bin_norm"] != ""]
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

    investern["Account Number"] = investern["Account Number"].astype(str).str.upper().str.strip()
    investern["_id_norm"] = investern["Account Number"].apply(norm_key)

    investern["Invester Commitment"] = pd.to_numeric(investern["Invester Commitment"], errors="coerce").fillna(0)
    investern["SS Commitment"] = investern["_id_norm"].map(ss_source).fillna(0)
    investern["SS Check"] = investern["SS Commitment"] - investern["Invester Commitment"]

    # ---- 6. COMBINE ----
    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)})

    df = df.reindex(range(max_rows)).reset_index(drop=True)
    investern = investern.reindex(range(max_rows)).reset_index(drop=True)

    combined_df = pd.concat([df.astype(object), spacer, investern.astype(object)], axis=1)

    # ---- 7. SS TOTAL ----
    subtotal_row = {col: "" for col in combined_df.columns}
    subtotal_row.update({
        "Vehicle/Investor": "Subtotal (SS Total)",
        "Invester Commitment": investern["Invester Commitment"].sum(),
        "SS Commitment": investern["SS Commitment"].sum(),
        "SS Check": investern["SS Commitment"].sum() - investern["Invester Commitment"].sum(),
    })

    combined_df = pd.concat([combined_df, pd.DataFrame([subtotal_row])], ignore_index=True)

    # ---- 8. CLEANUP ----
    drop_cols = [c for c in combined_df.columns if c.startswith("_")]
    combined_df.drop(columns=drop_cols, inplace=True, errors="ignore")
    combined_df = clean_dataframe(combined_df)

    # ---- 9. WRITE ----
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        combined_df.to_excel(writer, sheet_name="Commitment Sheet", index=False)

    print("Commitment Sheet created successfully.")
    return combined_df
