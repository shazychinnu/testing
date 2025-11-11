import pandas as pd
import re
from openpyxl import load_workbook


def create_entry_sheet_with_subtotals(commitment_df):
    """
    Create an 'Entry' sheet in the output Excel file, combining data from wizard and commitment files,
    calculating subtotals per block, and joining commitment and BIN mapping data.
    Dynamic handling for CDR Summary sheet (1-row or 2-row headers supported).
    """

    # ---------------------------------------------------------------------
    # Helper: normalize keys for matching
    # ---------------------------------------------------------------------
    def norm_key(x):
        if pd.isna(x):
            return ""
        s = str(x).strip()
        if s.endswith(".0"):
            s = s[:-2]
        s = s.replace("\u00A0", " ")
        s = re.sub(r"[,\-]", "", s)
        return s.upper()

    # ---------------------------------------------------------------------
    # Helper: delete existing Entry sheet if present
    # ---------------------------------------------------------------------
    def delete_sheet_if_exists(file, sheet_name):
        try:
            wb = load_workbook(file)
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
                wb.save(file)
        except FileNotFoundError:
            pass

    # ---------------------------------------------------------------------
    # Helper: basic clean up
    # ---------------------------------------------------------------------
    def clean_dataframe(df):
        df = df.copy()
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        df.reset_index(drop=True, inplace=True)
        return df

    # ---------------------------------------------------------------------
    # Delete old sheet
    # ---------------------------------------------------------------------
    delete_sheet_if_exists(output_file, "Entry")

    # ---------------------------------------------------------------------
    # Load wizard allocation data
    # ---------------------------------------------------------------------
    df_raw = pd.read_excel(wizard_file, sheet_name="allocation_data",
                           engine="openpyxl", header=None)
    cdr = cdr_file_data  # keep reference

    # ---------------------------------------------------------------------
    # Split wizard data into tables & add subtotals
    # ---------------------------------------------------------------------
    header_rows = df_raw.index[df_raw.iloc[:, 0].astype(str) == "Vehicle/Investor"].tolist()
    tables = []

    for i, h in enumerate(header_rows):
        start = h
        end = header_rows[i + 1] if i + 1 < len(header_rows) else len(df_raw)
        block = df_raw.iloc[start:end].reset_index(drop=True)
        block.columns = block.iloc[0]
        block = block.drop(0).reset_index(drop=True)

        if "Final LE Amount" in block.columns:
            block["Final LE Amount"] = pd.to_numeric(block["Final LE Amount"],
                                                     errors="coerce").fillna(0)
            subtotal_val = block["Final LE Amount"].sum(skipna=True)
        else:
            subtotal_val = 0

        subtotal_row = {col: "" for col in block.columns}
        subtotal_row[block.columns[0]] = "Subtotal"
        if "Final LE Amount" in block.columns:
            subtotal_row["Final LE Amount"] = subtotal_val

        block = pd.concat([block, pd.DataFrame([subtotal_row])], ignore_index=True)
        tables.append(block)

    final_df = pd.concat(tables, ignore_index=True) if tables else pd.DataFrame()

    # ---------------------------------------------------------------------
    # Normalize IDs
    # ---------------------------------------------------------------------
    id_col = "Investor ID" if "Investor ID" in final_df.columns else "Investor Id"
    final_df["_id_norm"] = final_df[id_col].apply(norm_key)

    cm = commitment_df.copy()
    cm["_inv_acct_norm"] = cm["Investran Acct ID"].apply(norm_key)

    # ---------------------------------------------------------------------
    # Load CDR Summary By Investor (supports 1 or 2 header rows)
    # ---------------------------------------------------------------------
    if isinstance(cdr_file_data, pd.DataFrame):
        cdr = cdr_file_data.copy()
    else:
        try:
            xl = pd.ExcelFile(cdr_file_data, engine="openpyxl")
            possible_sheets = [s for s in xl.sheet_names if "CDR" in s and "Investor" in s]
            sheet_name = possible_sheets[0] if possible_sheets else xl.sheet_names[0]

            # Try two-row header first
            try:
                cdr = pd.read_excel(xl, sheet_name=sheet_name, header=[0, 1])
                cdr.columns = [
                    f"{lvl1}_{lvl2}".strip().replace(" ", "_")
                    if not str(lvl2).startswith("Unnamed")
                    else lvl1.strip().replace(" ", "_")
                    for lvl1, lvl2 in cdr.columns
                ]
            except ValueError:
                # fallback single header
                cdr = pd.read_excel(xl, sheet_name=sheet_name, header=0)
                cdr.columns = [str(c).strip().replace(" ", "_") for c in cdr.columns]

        except Exception as e:
            raise ValueError(f"❌ Error loading CDR Summary: {e}")

    # ---------------------------------------------------------------------
    # Normalize BIN IDs & join
    # ---------------------------------------------------------------------
    bin_cols = [c for c in cdr.columns if "BIN" in c.upper()]
    if not bin_cols:
        raise ValueError("No BIN column found in CDR Summary sheet.")
    bin_col = bin_cols[0]

    cdr[bin_col] = cdr[bin_col].apply(norm_key)
    cm["_inv_acct_norm"] = cm["_inv_acct_norm"].astype(str).fillna("").str.strip().str.upper()

    id_to_bin = (
        cm.dropna(subset=["Bin ID"])
          .drop_duplicates(subset=["_inv_acct_norm"])
          .set_index("_inv_acct_norm")["Bin ID"]
          .to_dict()
    )
    final_df["Bin ID"] = final_df["_id_norm"].map(id_to_bin)

    merged_df = pd.merge(final_df, cdr, left_on="Bin ID", right_on=bin_col, how="left")

    # ---------------------------------------------------------------------
    # Compute totals dynamically
    # ---------------------------------------------------------------------
    numeric_cols = merged_df.select_dtypes(include=["number"]).columns.tolist()
    merged_df[numeric_cols] = merged_df[numeric_cols].fillna(0)

    totals = merged_df[numeric_cols].sum(numeric_only=True)
    total_row = {col: "" for col in merged_df.columns}
    for col, val in totals.items():
        total_row[col] = val
    total_row[merged_df.columns[0]] = "TOTAL"

    merged_df = pd.concat([merged_df, pd.DataFrame([total_row])], ignore_index=True)

    # ---------------------------------------------------------------------
    # Final clean and save
    # ---------------------------------------------------------------------
    merged_df.drop(columns=[c for c in merged_df.columns if c.startswith("_")],
                   inplace=True, errors="ignore")
    merged_df = clean_dataframe(merged_df)

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        merged_df.to_excel(writer, sheet_name="Entry", index=False)

    print("✅ Entry Sheet created successfully with dynamic header handling.")
