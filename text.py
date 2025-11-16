def clean_legal_entity(name):
    """Normalize Subtotal Legal Entity names ONLY (do NOT remove HOLDING here)."""
    if not isinstance(name, str):
        return ""
    s = name.upper().strip()

    # Remove SUBTOTAL: prefix only
    s = s.replace("SUBTOTAL:", "")

    # Normalize multiple spaces
    s = " ".join(s.split())

    return s.strip()


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

    # GS source: Multi-key & fallback
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

    # ---- 3. GS COMMITMENT from CDR ----
    def lookup_gs(row):
        key = (row["_bin_norm"], row["_inv_acct_norm"])
        if key in multi_key_commit:
            return multi_key_commit[key]
        return bin_only_commit.get(row["_bin_norm"], 0)

    df["GS Commitment"] = df.apply(lookup_gs, axis=1)
    df["GS Check"] = df["Commitment Amount"] - df["GS Commitment"]

    # ---- 4. Split sections based on Subtotal rows ----
    subtotal_mask = df["Legal Entity"].str.upper().str.contains("SUBTOTAL", na=False)
    subtotal_indices = list(df.index[subtotal_mask])

    sections = []
    start = 0
    for idx in subtotal_indices:
        section = df.iloc[start:idx + 1].copy()
        sections.append(section)
        start = idx + 1

    # ---- 5. Build section_totals using normalized names ----
    section_totals = {}
    for section in sections:
        sec = section.reset_index(drop=True)
        subtotal_row = sec.iloc[-1]

        legal_norm = clean_legal_entity(subtotal_row["Legal Entity"])
        gs_value = float(subtotal_row["GS Commitment"])

        section_totals[legal_norm] = gs_value

    # ---- 6. Process sections with FEEDER fix ----
    processed_sections = []

    for section in sections:
        section = section.reset_index(drop=True)
        subtotal_pos = len(section) - 1
        data_rows = section.iloc[:subtotal_pos]

        # Normalized Subtotal Section Legal Entity
        subtotal_legal_norm = clean_legal_entity(section.iloc[-1]["Legal Entity"])

        # Correct GS subtotal for this Legal Entity
        correct_section_gs = section_totals.get(subtotal_legal_norm, 0)

        # FEEDER identification
        feeder_mask = data_rows["Bin ID"].str.upper().str.contains("FEEDER", na=False)
        feeder_positions = data_rows.index[feeder_mask].tolist()

        # ---- FEEDER-only HOLDING removal ----
        for pos in feeder_positions:
            feeder_legal = section.at[pos, "Legal Entity"].upper()
            feeder_legal = feeder_legal.replace("HOLDING", "")  # ONLY remove HOLDING here
            feeder_legal = " ".join(feeder_legal.split())      # normalize spaces

            # If FEEDER legal matches this section → apply GS subtotal
            if feeder_legal == subtotal_legal_norm:
                section.at[pos, "GS Commitment"] = correct_section_gs
                section.at[pos, "GS Check"] = section.at[pos, "Commitment Amount"] - correct_section_gs

        # ---- Recalculate FINAL subtotal ----
        updated_data = section.iloc[:subtotal_pos]
        final_commit_sum = pd.to_numeric(updated_data["Commitment Amount"], errors="coerce").fillna(0).sum()
        final_gs_sum = pd.to_numeric(updated_data["GS Commitment"], errors="coerce").fillna(0).sum()

        # Write final subtotal back to the subtotal row
        section.at[subtotal_pos, "Commitment Amount"] = final_commit_sum
        section.at[subtotal_pos, "GS Commitment"] = final_gs_sum
        section.at[subtotal_pos, "GS Check"] = final_commit_sum - final_gs_sum

        processed_sections.append(section)

    # Combine processed sections
    df = pd.concat(processed_sections, ignore_index=True)

    # ---- 7. SS Commitment (unchanged) ----
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

    investern["Invester Commitment"] = pd.to_numeric(investern["Invester Commitment"], errors="coerce").fillna(0)
    investern["SS Commitment"] = investern["_id_norm"].map(ss_source).fillna(0)
    investern["SS Check"] = investern["SS Commitment"] - investern["Invester Commitment"]

    # ---- 8. Final combine ----
    max_rows = max(len(df), len(investern))
    spacer = pd.DataFrame({f"Empty_{i}": [""] * max_rows for i in range(3)}, dtype=object)

    df = df.reindex(range(max_rows)).reset_index(drop=True)
    investern = investern.reindex(range(max_rows)).reset_index(drop=True)

    combined_df = pd.concat([df.astype(object), spacer, investern.astype(object)], axis=1)

    # ---- 9. SS Total ----
    subtotal_row = {col: "" for col in combined_df.columns}
    subtotal_row.update({
        "Vehicle/Investor": "Subtotal (SS Total)",
        "Invester Commitment": investern["Invester Commitment"].sum(),
        "SS Commitment": investern["SS Commitment"].sum(),
        "SS Check": investern["SS Commitment"].sum() - investern["Invester Commitment"].sum()
    })

    combined_df = pd.concat([combined_df, pd.DataFrame([subtotal_row])], ignore_index=True)

    # ---- 10. Cleanup and Write ----
    internal_cols = [c for c in combined_df.columns if c.startswith("_")]
    combined_df.drop(columns=internal_cols, inplace=True, errors="ignore")
    combined_df = clean_dataframe(combined_df)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        combined_df.to_excel(writer, sheet_name="Commitment Sheet", index=False)

    print("Commitment Sheet created successfully.")
    return combined_df
