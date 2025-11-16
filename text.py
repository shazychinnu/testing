def create_commitment_sheet():
    ensure_output_file_exists()
    delete_sheet_if_exists(output_file, "Commitment Sheet")

    # ---- 1. Load CDR Summary (use existing cdr_file_data if available) ----
    # If you already have cdr_file_data preloaded, use a copy. Otherwise load.
    try:
        cdr = cdr_file_data.copy()
    except NameError:
        cdr = pd.read_excel(
            cdr_file,
            sheet_name="CDR Summary By Investor",
            engine="openpyxl",
            skiprows=2,
            dtype={"Account Number": str, "Investor ID": str}
        )

    cdr.columns = cdr.columns.str.strip()
    cdr["Account Number"] = cdr["Account Number"].astype(str).str.strip()
    cdr["Investor ID"] = cdr["Investor ID"].astype(str).str.strip()

    cdr["_bin_norm"] = cdr["Account Number"].apply(norm_key)
    cdr["_investor_norm"] = cdr["Investor ID"].apply(norm_key)
    cdr["Investor Commitment"] = pd.to_numeric(cdr["Investor Commitment"], errors="coerce").fillna(0)

    # Build lookups
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
        dtype={"Investran Acct ID": str, "Bin ID": str, "Legal Entity": str}
    )
    df.columns = df.columns.str.strip()

    df["Legal Entity"] = df["Legal Entity"].astype(str).fillna("").str.strip()
    df["Commitment Amount"] = pd.to_numeric(df["Commitment Amount"], errors="coerce").fillna(0)

    df["Bin ID"] = df["Bin ID"].astype(str).fillna("").str.strip()
    df["Investran Acct ID"] = df["Investran Acct ID"].astype(str).fillna("").str.strip()

    df["_bin_norm"] = df["Bin ID"].apply(norm_key)
    df["_inv_acct_norm"] = df["Investran Acct ID"].apply(norm_key)

    # ---- 3. GS COMMITMENT LOOKUP (per-row) ----
    def lookup_gs(row):
        key1 = (row["_bin_norm"], row["_inv_acct_norm"])
        if key1 in multi_key_commit:
            return multi_key_commit[key1]
        return bin_only_commit.get(row["_bin_norm"], 0)

    df["GS Commitment"] = df.apply(lookup_gs, axis=1)
    df["GS Commitment"] = pd.to_numeric(df["GS Commitment"], errors="coerce").fillna(0)
    df["GS Check"] = df["Commitment Amount"] - df["GS Commitment"]

    # ---- 4. Split into sections (each section includes its subtotal row at the end) ----
    subtotal_mask = df["Legal Entity"].str.upper().str.contains("SUBTOTAL", na=False)
    subtotal_indices = list(df.index[subtotal_mask])

    sections = []
    start = 0
    for idx in subtotal_indices:
        # include subtotal row at idx
        section = df.iloc[start: idx + 1].copy()
        sections.append(section)
        start = idx + 1

    # If there are trailing rows without a subtotal (unlikely given your rules), add them
    if start < len(df):
        tail = df.iloc[start:].copy()
        sections.append(tail)

    # ---- 5. Process each section independently ----
    processed_sections = []
    for section in sections:
        # reset index for safe positional indexing inside section
        section = section.reset_index(drop=True)

        # must have at least 1 row; subtotal row is expected to be the last row in the section
        if len(section) == 0:
            processed_sections.append(section)
            continue

        subtotal_pos = len(section) - 1  # position (integer) of subtotal row inside section

        # data_rows are all rows before last row (exclude subtotal row)
        data_rows = section.iloc[:subtotal_pos].copy()

        # --- compute initial subtotal GS from data_rows BEFORE FEEDER changes ---
        # (this is the initial subtotal the FEEDER rows should use)
        initial_subtotal_gs = pd.to_numeric(data_rows["GS Commitment"], errors="coerce").fillna(0).sum()

        # Detect FEEDER rows inside data_rows: BIN contains "FEEDER" (case-insensitive)
        if not data_rows.empty:
            feeder_mask = data_rows["Bin ID"].str.upper().str.contains("FEEDER", na=False)

            # Update FEEDER GS Commitment to initial_subtotal_gs (always overwrite GS Commitment for feeders)
            # Use .loc on the section (positional mapping)
            if feeder_mask.any():
                # Map mask from data_rows back to section positions (same order, so positions match 0..subtotal_pos-1)
                feeder_positions = data_rows.index[feeder_mask].tolist()  # these are positions inside data_rows (and section)
                for pos in feeder_positions:
                    section.at[pos, "GS Commitment"] = initial_subtotal_gs
                    # update GS Check for feeder row
                    # Commitment Amount is left unchanged
                    section.at[pos, "GS Check"] = section.at[pos, "Commitment Amount"] - initial_subtotal_gs

        # --- Recalculate final subtotal after FEEDER updates ---
        # recompute sums from updated data_rows (section rows 0..subtotal_pos-1)
        updated_data = section.iloc[:subtotal_pos]
        final_commit_sum = pd.to_numeric(updated_data["Commitment Amount"], errors="coerce").fillna(0).sum()
        final_gs_sum = pd.to_numeric(updated_data["GS Commitment"], errors="coerce").fillna(0).sum()

        # Write final subtotal values into the subtotal row (last row)
        section.at[subtotal_pos, "Commitment Amount"] = final_commit_sum
        section.at[subtotal_pos, "GS Commitment"] = final_gs_sum
        section.at[subtotal_pos, "GS Check"] = final_commit_sum - final_gs_sum

        # ensure types consistent
        section["Commitment Amount"] = pd.to_numeric(section["Commitment Amount"], errors="coerce").fillna(0)
        section["GS Commitment"] = pd.to_numeric(section["GS Commitment"], errors="coerce").fillna(0)
        section["GS Check"] = pd.to_numeric(section["GS Check"], errors="coerce").fillna(0)

        processed_sections.append(section)

    # ---- 6. Re-combine processed sections into one dataframe ----
    if processed_sections:
        df = pd.concat(processed_sections, ignore_index=True)
    else:
        df = pd.DataFrame(columns=df.columns)

    # Recompute GS Check for entire df as a safety
    df["GS Check"] = pd.to_numeric(df["Commitment Amount"], errors="coerce").fillna(0) - pd.to_numeric(df["GS Commitment"], errors="coerce").fillna(0)

    # ---- 7. SS Commitment (unchanged) ----
    ss_source = (
        df[df["_bin_norm"].notna() & (df["_bin_norm"] != "")]
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

    # ---- 8. Combine left/right into final combined_df ----
    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)}, dtype=object)

    df = df.reindex(range(max_rows)).reset_index(drop=True)
    investern = investern.reindex(range(max_rows)).reset_index(drop=True)

    combined_df = pd.concat([df.astype(object), spacer, investern.astype(object)], axis=1)

    # ---- 9. SS Subtotal ----
    subtotal_row = {col: "" for col in combined_df.columns}
    subtotal_row.update({
        "Vehicle/Investor": "Subtotal (SS Total)",
        "Invester Commitment": investern["Invester Commitment"].sum(),
        "SS Commitment": investern["SS Commitment"].sum(),
        "SS Check": investern["SS Commitment"].sum() - investern["Invester Commitment"].sum()
    })
    combined_df = pd.concat([combined_df, pd.DataFrame([subtotal_row], dtype=object)], ignore_index=True)

    # ---- 10. Cleanup & write ----
    internal_cols = [c for c in combined_df.columns if c.startswith("_")]
    combined_df.drop(columns=internal_cols, inplace=True, errors="ignore")
    combined_df = clean_dataframe(combined_df)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        combined_df.to_excel(writer, sheet_name="Commitment Sheet", index=False)

    print("Commitment Sheet created successfully.")
    return combined_df
