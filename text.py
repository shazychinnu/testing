import pandas as pd
import re
from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------------------------

def norm_key(x):
    """Normalize keys for matching (remove .0, spaces, commas, hyphens, etc.)."""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.replace("\u00A0", " ")
    s = re.sub(r"[,\-]", "", s)
    return s.upper()


def delete_sheet_if_exists(output_file, sheet_name):
    """Delete a sheet from the Excel file if it already exists."""
    try:
        wb = load_workbook(output_file)
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
            wb.save(output_file)
    except FileNotFoundError:
        # File doesn't exist yet — fine for first run
        pass


def clean_dataframe(df):
    """Basic dataframe cleanup — strip strings and reset index."""
    df = df.copy()
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df.reset_index(drop=True, inplace=True)
    return df


# ---------------------------------------------------------------------------
# Main Function
# ---------------------------------------------------------------------------

def create_entry_sheet_dynamic(commitment_df, cdr_file_path, output_file):
    """
    Dynamically create an 'Entry' sheet using data from the CDR Summary By Investor sheet
    and commitment mapping. Handles variable headers and duplicate subcolumn names.
    """

    # -----------------------------------------------------------------------
    # STEP 1: Read and flatten multi-level CDR Summary By Investor headers
    # -----------------------------------------------------------------------
    cdr_df = pd.read_excel(
        cdr_file_path,
        sheet_name="CDR Summary By Investor",
        header=[0, 1],  # Two header rows
        engine="openpyxl"
    )

    # --- Flatten column names dynamically ---
    flattened_columns = []
    for lvl1, lvl2 in cdr_df.columns:
        # Handle Unnamed or empty secondary headers
        if pd.isna(lvl2) or str(lvl2).startswith("Unnamed"):
            flat_name = str(lvl1).strip().replace(" ", "_")
        else:
            flat_name = f"{str(lvl1).strip()}_{str(lvl2).strip()}".replace(" ", "_")
        flattened_columns.append(flat_name)

    cdr_df.columns = flattened_columns

    # -----------------------------------------------------------------------
    # STEP 2: Normalize keys for joining
    # -----------------------------------------------------------------------
    # Ensure Bin ID exists (or similar)
    possible_bin_cols = [col for col in cdr_df.columns if "Bin" in col or "BIN" in col]
    if not possible_bin_cols:
        raise ValueError("CDR Summary sheet must contain a 'Bin ID' or similar column.")
    bin_col = possible_bin_cols[0]  # Use the first detected Bin ID column

    cdr_df[bin_col] = cdr_df[bin_col].apply(norm_key)
    commitment_df["Bin_ID"] = commitment_df["Bin ID"].apply(norm_key)

    # -----------------------------------------------------------------------
    # STEP 3: Merge commitment and CDR data on Bin ID
    # -----------------------------------------------------------------------
    merged_df = pd.merge(commitment_df, cdr_df, left_on="Bin_ID", right_on=bin_col, how="left")
    merged_df.fillna(0, inplace=True)

    # -----------------------------------------------------------------------
    # STEP 4: Identify numeric columns dynamically
    # -----------------------------------------------------------------------
    numeric_cols = merged_df.select_dtypes(include=["number"]).columns.tolist()
    main_cols = [c for c in merged_df.columns if c not in numeric_cols]

    # -----------------------------------------------------------------------
    # STEP 5: Add a grand total row
    # -----------------------------------------------------------------------
    totals = merged_df[numeric_cols].sum(numeric_only=True)
    total_row = pd.DataFrame([[ "TOTAL" ] + [0]*(len(main_cols)-1) + list(totals)], 
                             columns=main_cols + numeric_cols)
    entry_df = pd.concat([merged_df, total_row], ignore_index=True)

    # -----------------------------------------------------------------------
    # STEP 6: Clean and save to Excel
    # -----------------------------------------------------------------------
    entry_df = clean_dataframe(entry_df)
    delete_sheet_if_exists(output_file, "Entry")

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a") as writer:
        entry_df.to_excel(writer, sheet_name="Entry", index=False)

    print(f"✅ Entry Sheet created successfully with {len(entry_df)} rows and dynamic columns.")
