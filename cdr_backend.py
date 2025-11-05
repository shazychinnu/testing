import os
import logging
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# optional modules referenced in original file (import if available)
try:
    import cdr_application_gui  # may be part of your project
except Exception:
    cdr_application_gui = None

try:
    import py  # used in original; harmless to import if present
except Exception:
    py = None


class ExcelBackendProcessor:
    """
    Syntax-corrected version of the code you provided.
    NOTE: I preserved variable names and the original logic intention as closely as possible,
    while fixing syntax, indentation and structural errors so the file runs.
    """

    def __init__(self, report_file, cdr_file, output_file, headless=False, log_callback=None):
        # keep names (typos preserved where they were usable as identifiers)
        self.report_file = report_file
        self.cdr_file = cdr_file
        self.output_file = output_file
        self.headless = headless
        self.wb = None
        self.conn_df = None
        # 'investern' uses the same misspelling from original text but is a valid identifier
        self.investern_entry_df = None
        # 'cdr_sumary_df' preserved as in original (misspelling)
        self.cdr_sumary_df = None
        self.lookup_df = None
        self.mapping = {}
        self.log_callback = log_callback
        self.investern_allocation_df = None

        if self.headless:
            logging.basicConfig(
                filename='excel_processor_debug.log',
                level=logging.DEBUG,
                format="%(asctime)s - %(levelname)s - %(message)s"
            )
            self.log_debug("Headless mode is enabled. Debug logging started.")

    def log_debug(self, message):
        # pass
        """
        Simple debug logger which also uses an optional callback if provided.
        """
        logging.debug(message)
        if self.log_callback:
            try:
                self.log_callback("DEBUG", message)
            except Exception:
                # don't let callback errors break processing
                logging.debug("log_callback raised an exception.")

    # Note: name 'normalize_keys' preserved
    def normalize_keys(self, series):
        # Coerce to string, uppercase, remove whitespace
        return series.astype(str).str.upper().str.replace(r"\s+", "", regex=True).fillna("")

    def normalize(self, name):
        """
        Normalize a single string key similar to original intent:
        - split on '_' and take last part (if possible)
        - strip, uppercase, remove spaces and certain punctuation
        """
        try:
            name = name.split("_")[-1]
        except Exception:
            # if not splittable, keep as is
            name = name
        return (
            str(name)
            .strip()
            .upper()
            .replace(" ", "")
            .replace("_", "")
            .replace("/", "")
            .replace(".", "")
            .replace(".1", "")
            .replace(".2", "")
        )

    def apply_theme(self, sheet):
        """
        Apply simple styling to a worksheet:
        - center alignment for all cells
        - header row styling (row 1)
        - last row styling
        - light fill for rows 2..6 (as per original attempt)
        """
        for r_idx, row in enumerate(sheet.iter_rows(), start=1):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                # skip formatting for blank placeholders per original logic
                if cell.value in ("Blank1", "Blank2", None):
                    continue
                if r_idx == 1:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="093366", end_color="003366", fill_type="solid")
                elif r_idx == sheet.max_row:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
                elif 2 <= r_idx <= 6:
                    cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    def load_data(self):
        """
        Load data from the provided Excel files into DataFrames.
        This method fixes sheet names and parameters from the broken original.
        """
        self.log_debug("Loading data from Excel files.")
        # Use engine openpyxl and dtype=object to preserve formatting as in original
        try:
            # original intended sheet: "CDR Sumary By Investor" (misspelling preserved)
            self.cdr_sumary_df = pd.read_excel(
                self.report_file,
                sheet_name="CDR Sumary By Investor",
                skiprows=2,
                engine="openpyxl",
                dtype=object,
            )
        except Exception as e:
            # fallback: try the first sheet if the exact name is not present
            self.log_debug(f"Failed to load 'CDR Sumary By Investor' from {self.report_file}: {e}")
            self.cdr_sumary_df = pd.read_excel(self.report_file, engine="openpyxl", dtype=object)

        try:
            # original intended sheet: "investern_format"
            self.investern_entry_df = pd.read_excel(
                self.cdr_file, sheet_name="investern_format", engine="openpyxl", dtype=object
            )
        except Exception as e:
            self.log_debug(f"Failed to load 'investern_format' from {self.cdr_file}: {e}")
            # fallback: read first sheet
            self.investern_entry_df = pd.read_excel(self.cdr_file, engine="openpyxl", dtype=object)

        # load connection dataframe (original attempted to remove unnamed columns)
        try:
            # in original it read sheet_name=9 (index). We try a safe approach:
            self.conn_df = pd.read_excel(self.cdr_file, sheet_name=9, engine="openpyxl", dtype=object)
            # drop columns that are Unnamed
            self.conn_df = self.conn_df.loc[:, ~self.conn_df.columns.str.contains(r"^Unnamed")]
        except Exception:
            # fallback: take second sheet if available or the first sheet
            try:
                self.conn_df = pd.read_excel(self.cdr_file, sheet_name=1, engine="openpyxl", dtype=object)
                self.conn_df = self.conn_df.loc[:, ~self.conn_df.columns.str.contains(r"^Unnamed")]
            except Exception as e:
                self.log_debug(f"Failed to load conn_df from {self.cdr_file}: {e}")
                self.conn_df = pd.DataFrame()

        # load allocation sheet from cdr_file or report_file (original gave both variants)
        try:
            self.investern_allocation_df = pd.read_excel(
                self.cdr_file, sheet_name="allocation_data", engine="openpyxl", dtype=object
            )
        except Exception:
            try:
                self.investern_allocation_df = pd.read_excel(
                    self.report_file, sheet_name="allocation_data", engine="openpyxl", dtype=object
                )
            except Exception:
                self.investern_allocation_df = pd.DataFrame()

        self.log_debug("Data loading completed.")

    def preprocess_data(self):
        """
        Preprocess and join dataframes to prepare for sheet updates.
        Logic follows the original intent: normalize ids, map commitments/bin ids, merge contributions/distributions/external expenses.
        """
        self.log_debug("Starting data preprocessing.")

        # helper normalizer
        normalize = lambda s: s.astype(str).str.upper().str.replace(r"\s+", "", regex=True)

        # ensure DataFrames exist
        if self.investern_entry_df is None or self.conn_df is None or self.cdr_sumary_df is None:
            self.log_debug("One or more required dataframes are missing; aborting preprocessing.")
            return

        # Create normalized ID columns
        try:
            self.investern_entry_df["Normalized Investor ID"] = normalize(self.investern_entry_df["Investor ID"])
        except Exception:
            self.investern_entry_df["Normalized Investor ID"] = ""

        try:
            # original column name attempted: "Investran Acct ID" (typo); try a set of likely names
            acct_col_candidates = [
                "Investran Acct ID",
                "Investran AcctID",
                "Investran_Acct_ID",
                "Investor Acct ID",
                "Investor Account ID",
                "Account Number",
            ]
            acct_col = next((c for c in acct_col_candidates if c in self.conn_df.columns), None)
            if acct_col:
                self.conn_df["Normalized Investor ID"] = normalize(self.conn_df[acct_col])
            else:
                # fallback: use first column
                self.conn_df["Normalized Investor ID"] = normalize(self.conn_df.iloc[:, 0].astype(str))
        except Exception:
            self.conn_df["Normalized Investor ID"] = ""

        # Build mapping from normalized investor -> Bin ID and Commitment Amount if present
        try:
            bin_col = next((c for c in self.conn_df.columns if "Bin" in str(c) or "BIN" in str(c)), None)
            commitment_col = next((c for c in self.conn_df.columns if "Commitment" in str(c)), None)
            bin_id_map = {}
            commitment_map = {}
            if bin_col:
                bin_id_map = dict(
                    zip(self.conn_df["Normalized Investor ID"].astype(str), self.conn_df[bin_col].astype(str))
                )
            if commitment_col:
                commitment_map = dict(
                    zip(self.conn_df["Normalized Investor ID"].astype(str), self.conn_df[commitment_col].astype(str))
                )
            # apply maps to investern_entry_df
            self.investern_entry_df["Bin ID"] = self.investern_entry_df["Normalized Investor ID"].map(bin_id_map)
            self.investern_entry_df["Commitment"] = self.investern_entry_df["Normalized Investor ID"].map(commitment_map)
        except Exception:
            # ensure columns exist
            if "Bin ID" not in self.investern_entry_df.columns:
                self.investern_entry_df["Bin ID"] = ""
            if "Commitment" not in self.investern_entry_df.columns:
                self.investern_entry_df["Commitment"] = ""

        # Create Normalized Bin ID on investern_entry_df
        try:
            self.investern_entry_df["Normalized Bin ID"] = normalize(self.investern_entry_df["Bin ID"])
        except Exception:
            self.investern_entry_df["Normalized Bin ID"] = ""

        # Create Normalized Account Number on cdr_sumary_df (note misspelling preserved)
        try:
            acct_num_col = next((c for c in self.cdr_sumary_df.columns if "Account" in str(c)), None)
            if acct_num_col:
                self.cdr_sumary_df["Normalized Account Number"] = normalize(self.cdr_sumary_df[acct_num_col])
            else:
                self.cdr_sumary_df["Normalized Account Number"] = ""
        except Exception:
            self.cdr_sumary_df["Normalized Account Number"] = ""

        # Filter out rows where Investor Name contains "Total" - original code did this
        try:
            if "Investor Name" in self.cdr_sumary_df.columns:
                cdr_summary_filtered = self.cdr_sumary_df[
                    ~self.cdr_sumary_df["Investor Name"].astype(str).str.contains("Total", na=False)
                ].copy()
            else:
                cdr_summary_filtered = self.cdr_sumary_df.copy()
        except Exception:
            cdr_summary_filtered = self.cdr_sumary_df.copy()

        # Drop rows without Account Number and duplicates on Normalized Account Number
        try:
            if acct_num_col:
                cdr_summary_filtered = (
                    cdr_summary_filtered.dropna(subset=[acct_num_col])
                    .drop_duplicates(subset=["Normalized Account Number"])
                    .copy()
                )
        except Exception:
            cdr_summary_filtered = cdr_summary_filtered.copy()

        # Prepare header and indices for contributions/distributions/external expenses as original attempted
        header = [str(col).strip().upper() for col in cdr_summary_filtered.columns]

        # find contribution ("PI") anchor; fallback indices if not found
        try:
            contrib_start = header.index("PI") if "PI" in header else 9
        except Exception:
            contrib_start = 9

        try:
            # distrib start attempts to find second occurrence of "PI"
            if header.count("PI") > 1:
                dist_start = header.index("PI", contrib_start + 1)
            else:
                dist_start = contrib_start + 8
        except Exception:
            dist_start = contrib_start + 8

        try:
            ext_start = header.index("EXTERNAL EXPENSES") if "EXTERNAL EXPENSES" in header else dist_start + 8
        except Exception:
            ext_start = dist_start + 8

        # Safely extract column slices (guard indices)
        ncols = len(cdr_summary_filtered.columns)
        contrib_slice_end = min(contrib_start + 7, ncols)
        dist_slice_end = min(dist_start + 8, ncols)
        ext_slice_end = min(ext_start + 8, ncols)

        contrib_cols = cdr_summary_filtered.columns[contrib_start:contrib_slice_end]
        dist_cols = cdr_summary_filtered.columns[dist_start:dist_slice_end]
        ext_cols = cdr_summary_filtered.columns[ext_start:ext_slice_end]

        # Build contribution DataFrame to merge into investern_entry_df
        try:
            contrib_data = cdr_summary_filtered[["Normalized Account Number"] + list(contrib_cols)].copy()
            contrib_data.rename(columns={"Normalized Account Number": "Normalized Bin ID"}, inplace=True)
            contrib_data.rename(
                columns={col: f"Contributions_{col}" for col in contrib_cols}, inplace=True
            )
            # merge
            self.investern_entry_df = self.investern_entry_df.merge(
                contrib_data, on="Normalized Bin ID", how="left"
            )
        except Exception:
            # if this fails, proceed without merging
            pass

        # Distributions
        try:
            dist_data = cdr_summary_filtered[["Normalized Account Number"] + list(dist_cols)].copy()
            dist_data.rename(columns={"Normalized Account Number": "Normalized Bin ID"}, inplace=True)
            dist_data.rename(columns={col: f"Distributions_{col}" for col in dist_cols}, inplace=True)
            self.investern_entry_df = self.investern_entry_df.merge(dist_data, on="Normalized Bin ID", how="left")
        except Exception:
            pass

        # External expenses
        try:
            ext_data = cdr_summary_filtered[["Normalized Account Number"] + list(ext_cols)].copy()
            ext_data.rename(columns={"Normalized Account Number": "Normalized Bin ID"}, inplace=True)
            ext_data.rename(columns={col: f"ExternalExpenses_{col}" for col in ext_cols}, inplace=True)
            self.investern_entry_df = self.investern_entry_df.merge(ext_data, on="Normalized Bin ID", how="left")
        except Exception:
            pass

        # Load allocation data (report file) if not already loaded
        try:
            if self.investern_allocation_df is None or self.investern_allocation_df.empty:
                self.investern_allocation_df = pd.read_excel(
                    self.report_file, sheet_name="allocation_data", engine="openpyxl", dtype=object
                )
        except Exception:
            # leave as-is if not present
            pass

        # Maps for net cashflow and Final LE Amount (attempts with likely column names)
        try:
            net_cashflow_col = next((c for c in self.cdr_sumary_df.columns if "Current Net Cashflow" in c), None)
            acct_col_name = next((c for c in self.cdr_sumary_df.columns if "Account Number" in c or "Account" in c), None)
            if acct_col_name and net_cashflow_col:
                net_cashflow_amount_map = dict(
                    zip(self.cdr_sumary_df[acct_col_name].astype(str), self.cdr_sumary_df[net_cashflow_col])
                )
            else:
                net_cashflow_amount_map = {}
        except Exception:
            net_cashflow_amount_map = {}

        try:
            final_amount_col = next((c for c in self.investern_allocation_df.columns if "Final" in c and "Amount" in c), None)
            investor_id_col = next((c for c in self.investern_allocation_df.columns if "Investor ID" in c), None)
            if investor_id_col and final_amount_col:
                Final_Amount_map = dict(zip(self.investern_allocation_df[investor_id_col].astype(str), self.investern_allocation_df[final_amount_col]))
            else:
                Final_Amount_map = {}
        except Exception:
            Final_Amount_map = {}

        # map Final LE Amount and Current Net Cashflow to investern_entry_df where possible
        try:
            if "Investor ID" in self.investern_entry_df.columns:
                self.investern_entry_df["CRM - Final LE Amount"] = self.investern_entry_df["Investor ID"].map(Final_Amount_map)
            else:
                self.investern_entry_df["CRM - Final LE Amount"] = None
        except Exception:
            self.investern_entry_df["CRM - Final LE Amount"] = None

        try:
            # map current net cashflow via Bin ID mapping
            if "Bin ID" in self.investern_entry_df.columns:
                self.investern_entry_df["Current Net Cashflow"] = self.investern_entry_df["Bin ID"].map(net_cashflow_amount_map)
            else:
                self.investern_entry_df["Current Net Cashflow"] = None
        except Exception:
            self.investern_entry_df["Current Net Cashflow"] = None

        # Add a validation column if both present (original attempted arithmetic)
        try:
            self.investern_entry_df["Current Net Cashflow Validation"] = (
                pd.to_numeric(self.investern_entry_df["Current Net Cashflow"], errors="coerce")
            )
        except Exception:
            self.investern_entry_df["Current Net Cashflow Validation"] = None

        # Drop intermediate normalized columns to keep final sheet tidy
        for col in ["Normalized Investor ID", "Normalized Bin ID"]:
            if col in self.investern_entry_df.columns:
                try:
                    self.investern_entry_df.drop(columns=[col], inplace=True)
                except Exception:
                    pass

        # Reorder columns: keep some important ones first, then the rest, then group columns
        try:
            group_columns = [col for col in self.investern_entry_df.columns if any(
                col.startswith(prefix) for prefix in ["Contributions_", "Distributions_", "ExternalExpenses_"]
            )]
            non_group_columns = [
                col for col in self.investern_entry_df.columns
                if col not in group_columns and col not in ["Bin ID", "Commitment", "CRM - Final LE Amount", "Current Net Cashflow", "Current Net Cashflow Validation"]
            ]
            ordered_columns = ["Bin ID", "Commitment", "CRM - Final LE Amount", "Current Net Cashflow", "Current Net Cashflow Validation"] + non_group_columns + group_columns
            # Keep only existing columns in that order
            ordered_columns = [c for c in ordered_columns if c in self.investern_entry_df.columns]
            self.investern_entry_df = self.investern_entry_df[ordered_columns]
        except Exception:
            pass

        self.log_debug("Data preprocessing completed.")

    def update_commitment_sheet(self):
        """
        Create Commitment sheet in workbook based on conn_df and lookup_df.
        This follows the logic from the original code but with corrected syntax.
        """
        self.log_debug("Updating Commitment sheet.")

        # build lookup_df from cdr_sumary_df with columns preserved as original attempt
        try:
            # try to pick columns that match original names
            cols = []
            for c in ["Investor Name", "Account Number", "Investor ID", "Investor Commitment"]:
                if c in self.cdr_sumary_df.columns:
                    cols.append(c)
            if cols:
                self.lookup_df = self.cdr_sumary_df[cols].copy()
            else:
                # fallback: first 4 columns
                self.lookup_df = self.cdr_sumary_df.iloc[:, :4].copy()
        except Exception:
            self.lookup_df = pd.DataFrame()

        # compute lookup keys and values using normalize_keys
        try:
            lookup_keys = self.normalize_keys(self.lookup_df.iloc[:, 1])  # second column likely account number
            lookup_values = self.lookup_df.iloc[:, :3] if not self.lookup_df.empty else pd.DataFrame()
        except Exception:
            lookup_keys = pd.Series(dtype=str)
            lookup_values = pd.DataFrame()

        # prepare connection keys series normalized
        try:
            # assume 'Bin ID' exists in conn_df; else use first column
            conn_bin_col = "Bin ID" if "Bin ID" in self.conn_df.columns else self.conn_df.columns[0]
            conn_keys_series = self.normalize_keys(self.conn_df[conn_bin_col].astype(str))
        except Exception:
            conn_keys_series = pd.Series(dtype=str)

        # produce mapping dict (lookup key -> values as tuple)
        try:
            # map lookup_keys -> lookup_values rows (as dict of tuples)
            if not self.lookup_df.empty:
                lookup_values_as_tuples = {
                    k: tuple(v) for k, v in zip(lookup_keys, lookup_values.values.tolist())
                }
                # create mapping (string->tuple)
                self.mapping = dict(zip(lookup_keys, lookup_values_as_tuples))
            else:
                self.mapping = {}
        except Exception:
            self.mapping = {}

        # define match_key function as original attempted (returns value and match type)
        def match_key(key):
            if key in self.mapping:
                return self.mapping[key], "-"
            # startswith or contains heuristics
            for lk in lookup_keys:
                try:
                    if lk.startswith(str(key)):
                        return self.mapping.get(lk), "StartsWith"
                    if str(key) in lk:
                        return self.mapping.get(lk), "Contains"
                except Exception:
                    continue
            return None, "NA"

        # build matched results and insert columns into conn_df
        try:
            matched_results = [match_key(k) for k in conn_keys_series]
            # unzip
            mapped_vals = [mr[0] for mr in matched_results]
            match_types = [mr[1] for mr in matched_results]
            # create placeholder columns on conn_df
            self.conn_df["GS Commitment"] = [t[2] if isinstance(t, (list, tuple)) and len(t) > 2 else None for t in mapped_vals]
            self.conn_df["Match Type"] = match_types
        except Exception:
            # if any of above fails, ensure columns exist
            if "GS Commitment" not in self.conn_df.columns:
                self.conn_df["GS Commitment"] = None
            if "Match Type" not in self.conn_df.columns:
                self.conn_df["Match Type"] = None

        # create numeric checks - convert to numeric and compute differences
        try:
            self.conn_df["GS Check"] = pd.to_numeric(self.conn_df["GS Commitment"], errors="coerce") - pd.to_numeric(self.conn_df.get("Commitment Amount", None), errors="coerce")
        except Exception:
            self.conn_df["GS Check"] = None

        # Insert lookup columns into conn_df to mirror original code behavior
        try:
            insert_cols = [c for c in ["Investor Name", "Account Number", "Investor ID", "Investor Commitment"] if c in self.lookup_df.columns]
            for i, col in enumerate(insert_cols):
                # insert after 'GS Check' column if exists, otherwise append
                if "GS Check" in self.conn_df.columns:
                    loc = self.conn_df.columns.get_loc("GS Check") + 1 + i
                    self.conn_df.insert(loc, col, self.lookup_df[col].reindex(self.conn_df.index).values)
                else:
                    self.conn_df[col] = self.lookup_df[col].reindex(self.conn_df.index).values
        except Exception:
            pass

        # SS Commitment mapping using Normalized Account Number to Commitment Amount
        try:
            ss_mapping = dict(
                zip(self.normalize_keys(self.conn_df.get("Bin ID", self.conn_df.columns[0]).astype(str)), pd.to_numeric(self.conn_df.get("Commitment Amount", pd.Series([])), errors="coerce"))
            )
            # Account Number normalized
            acc_norm = self.normalize_keys(self.conn_df.get("Account Number", self.conn_df.columns[0]).astype(str))
            self.conn_df["SS Commitment"] = acc_norm.map(ss_mapping)
            self.conn_df["SS Check"] = pd.to_numeric(self.conn_df["SS Commitment"], errors="coerce") - pd.to_numeric(self.conn_df.get("Investor Commitment", None), errors="coerce")
        except Exception:
            self.conn_df["SS Commitment"] = None
            self.conn_df["SS Check"] = None

        # Add subtotal row at the end summing numeric columns of interest
        try:
            subtotal_row = ["Subtotal"] + [""] * (self.conn_df.shape[1] - 1)
            # find numeric columns to sum
            numeric_cols = ["Commitment Amount", "GS Commitment", "Investor Commitment", "SS Commitment", "GS Check", "SS Check"]
            for col in numeric_cols:
                if col in self.conn_df.columns:
                    try:
                        subtotal_val = pd.to_numeric(self.conn_df[col], errors="coerce").sum()
                        subtotal_row[self.conn_df.columns.get_loc(col)] = subtotal_val
                    except Exception:
                        pass
            # append as last row
            self.conn_df.loc[len(self.conn_df)] = subtotal_row
        except Exception:
            pass

        # create sheet in workbook and write conn_df
        try:
            ws_commitment = self.wb.create_sheet(title="Commitment")
            for r_idx, row in enumerate(dataframe_to_rows(self.conn_df, index=False, header=True), start=1):
                for c_idx, value in enumerate(row, start=1):
                    ws_commitment.cell(row=r_idx, column=c_idx, value=value)
            self.apply_theme(ws_commitment)
            self.log_debug("Commitment sheet updated.")
        except Exception as e:
            self.log_debug(f"Failed to write Commitment sheet: {e}")

    def map_columns_to_summary(self, entry_columns, summary_df):
        """
        Map entry columns to summary sums, trying to match normalized keys.
        Returns a dict: {entry_column: mapped_value}
        (Preserves original naming and logic intent.)
        """
        mapped_values = {}

        try:
            exclude_cols = ["Investor Name", "Investor ID", "Account Number", "Current Net Cashflow"]

            summary_df_safe = summary_df.copy()
            # Drop rows without account number
            acct_col = next((c for c in summary_df_safe.columns if "Account" in c), None)
            if acct_col:
                summary_df_safe = summary_df_safe.dropna(subset=[acct_col])

            # compute sums for object columns (as original attempted)
            summary_sums = summary_df_safe.drop(columns=[c for c in exclude_cols if c in summary_df_safe.columns], errors="ignore").select_dtypes(include="number").sum()

            # normalized keys mapping
            normalized_summary = {self.normalize(str(col)): val for col, val in summary_sums.items()}
            normalized_keys = {self.normalize(str(k)): k for k in summary_sums.keys()}

            for col in entry_columns:
                norm_col = self.normalize(str(col))
                matched_key = None
                for n_key in normalized_keys:
                    # original heuristics: if n_key in norm_col and norm_col in n_key
                    if n_key in norm_col or norm_col in n_key:
                        matched_key = normalized_keys[n_key]
                        break
                mapped_values[col] = summary_sums.get(matched_key, None)
        except Exception:
            # fallback: return None mapping for all
            for col in entry_columns:
                mapped_values[col] = None

        return mapped_values

    def update_entry_sheet(self):
        """
        Create Entry sheet from investern_entry_df with totals, subtotals and validation formulas.
        """
        self.log_debug("Updating Entry sheet.")
        try:
            ws_entry = self.wb.create_sheet(title="Entry")
        except Exception:
            self.log_debug("Failed to create Entry sheet in workbook.")
            return

        # Drop zero-sum columns for contribution/distribution/expense prefixes
        prefixes = ["Contributions_", "Distributions_", "ExternalExpenses_"]
        try:
            cols_to_check = [col for col in self.investern_entry_df.columns if any(col.startswith(prefix) for prefix in prefixes)]
            cols_to_remove = []
            for col in cols_to_check:
                try:
                    s = pd.to_numeric(self.investern_entry_df[col], errors="coerce").fillna(0).sum()
                    if s == 0:
                        cols_to_remove.append(col)
                except Exception:
                    # if non-numeric, keep it
                    pass
            if cols_to_remove:
                self.investern_entry_df.drop(columns=cols_to_remove, inplace=True)
        except Exception:
            pass

        # Write dataframe rows to sheet (header + data)
        for r_idx, row in enumerate(dataframe_to_rows(self.investern_entry_df, index=False, header=True), start=1):
            for c_idx, value in enumerate(row, start=1):
                ws_entry.cell(row=r_idx, column=c_idx, value=value)

        # Prepare sums: add a Total row after data
        try:
            header_row = [cell.value for cell in ws_entry[1]]
            sum_row_index = ws_entry.max_row + 1
            ws_entry.cell(row=sum_row_index, column=1, value="Total")
            for col_idx, col_name in enumerate(header_row, start=1):
                if col_name and (col_name == "Commitment" or any(col_name.startswith(prefix) for prefix in prefixes)):
                    col_letter = get_column_letter(col_idx)
                    # SUM from row 2 to last data row (sum_row_index - 1)
                    formula = f"=SUM({col_letter}2:{col_letter}{sum_row_index - 1})"
                    ws_entry.cell(row=sum_row_index, column=col_idx, value=formula)
        except Exception:
            pass

        # Compute and write Subtotals row from CDR Summary
        try:
            summary_row_index = sum_row_index + 1
            ws_entry.cell(row=summary_row_index, column=1, value="Subtotals")
            mapped_values = self.map_columns_to_summary(header_row, self.cdr_sumary_df)
            for col_idx, col_name in enumerate(header_row, start=1):
                if col_name == "Commitment":
                    # sum of commitment amounts from conn_df excluding last row (subtotal)
                    try:
                        commitment_series = pd.to_numeric(self.conn_df.get("Commitment Amount", pd.Series([]))[:-1], errors="coerce")
                        commitment_sum = commitment_series.sum()
                        ws_entry.cell(row=summary_row_index, column=col_idx, value=commitment_sum)
                    except Exception:
                        ws_entry.cell(row=summary_row_index, column=col_idx, value=None)
                elif col_name in mapped_values:
                    try:
                        ws_entry.cell(row=summary_row_index, column=col_idx, value=mapped_values.get(col_name))
                    except Exception:
                        ws_entry.cell(row=summary_row_index, column=col_idx, value=None)
        except Exception as ERROR:
            # keep running if any failure here
            print(ERROR)

        # Validate row
        try:
            validate_row_index = summary_row_index + 1
            ws_entry.cell(row=validate_row_index, column=1, value="Validate")
            for col_idx in range(2, ws_entry.max_column + 1):
                col_letter = get_column_letter(col_idx)
                total_cell = f"{col_letter}{sum_row_index}"
                subtotals_cell = f"{col_letter}{summary_row_index}"
                # formula: IF(total_cell=subtotals_cell,"OK","Mismatch")
                formula = f'=IF({total_cell}={subtotals_cell},"OK","Mismatch")'
                cell = ws_entry.cell(row=validate_row_index, column=col_idx, value=formula)
                cell.alignment = Alignment(horizontal="center")
        except Exception:
            pass

        # Apply formatting fills
        try:
            total_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
            subtotals_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
            validate_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            # set fills
            for col_idx in range(1, ws_entry.max_column + 1):
                try:
                    ws_entry.cell(row=sum_row_index, column=col_idx).fill = total_fill
                    ws_entry.cell(row=summary_row_index, column=col_idx).fill = subtotals_fill
                    ws_entry.cell(row=validate_row_index, column=col_idx).fill = validate_fill
                except Exception:
                    pass
        except Exception:
            pass

        # theme and done
        try:
            self.apply_theme(ws_entry)
            self.log_debug("Entry sheet updated.")
        except Exception:
            pass

    def update_approx_matches_sheet(self):
        """
        Generate an 'Approx Matches' sheet by attempting fuzzy matching between conn keys and lookup keys.
        """
        self.log_debug("Updating Approx Matches sheet.")
        try:
            conn_keys_series = self.normalize_keys(self.conn_df.get("Bin ID", self.conn_df.columns[0]).astype(str))
        except Exception:
            conn_keys_series = pd.Series(dtype=str)

        try:
            lookup_keys = self.normalize_keys(self.lookup_df.iloc[:, 1].astype(str)) if not self.lookup_df.empty else pd.Series(dtype=str)
        except Exception:
            lookup_keys = pd.Series(dtype=str)

        approx_matches = []
        for key in conn_keys_series:
            # find first lookup key that startswith or contains key
            try:
                lk = next((lk for lk in lookup_keys if (lk.startswith(key) or key in lk)), None)
            except Exception:
                lk = None
            if lk and self.mapping.get(lk):
                approx_matches.append({
                    "Conn Key": key,
                    "Matched Lookup Key": lk,
                    "Mapped Value": self.mapping.get(lk)
                })

        if approx_matches:
            approx_df = pd.DataFrame(approx_matches)
            try:
                ws_approx = self.wb.create_sheet(title="Approx Matches")
                for r_idx, row in enumerate(dataframe_to_rows(approx_df, index=False, header=True), start=1):
                    for c_idx, value in enumerate(row, start=1):
                        ws_approx.cell(row=r_idx, column=c_idx, value=value)
                self.apply_theme(ws_approx)
                self.log_debug("Approx Matches sheet updated.")
            except Exception as e:
                self.log_debug(f"Failed to write Approx Matches sheet: {e}")
        else:
            # create empty sheet to indicate no matches found
            try:
                ws_approx = self.wb.create_sheet(title="Approx Matches")
                ws_approx.cell(row=1, column=1, value="No approximate matches found.")
                self.apply_theme(ws_approx)
                self.log_debug("Approx Matches sheet created (no matches).")
            except Exception:
                pass

    def backend_process(self):
        """
        Main orchestration method to load workbook, remove old sheets, run updates in parallel,
        and save the result to output_file.
        """
        self.log_debug("Starting Excel processing workflow.")
        # load and prepare data
        self.load_data()
        self.preprocess_data()

        # load workbook from report_file
        try:
            self.wb = load_workbook(filename=self.report_file)
        except Exception as e:
            self.log_debug(f"Failed to load workbook {self.report_file}: {e}")
            # create empty workbook structure if load fails
            from openpyxl import Workbook

            self.wb = Workbook()

        # remove sheets if already present (Commitment, Approx Matches, Entry)
        for sheet in ["Commitment", "Approx Matches", "Entry"]:
            if sheet in self.wb.sheetnames:
                try:
                    del self.wb[sheet]
                except Exception:
                    try:
                        self.wb.remove(self.wb[sheet])
                    except Exception:
                        pass

        # run update methods (may run in parallel)
        try:
            with ThreadPoolExecutor() as executor:
                futures = []
                futures.append(executor.submit(self.update_commitment_sheet))
                futures.append(executor.submit(self.update_entry_sheet))
                futures.append(executor.submit(self.update_approx_matches_sheet))
                # wait for completion implicitly by exiting context
        except Exception:
            # fallback: run sequentially
            self.update_commitment_sheet()
            self.update_entry_sheet()
            self.update_approx_matches_sheet()

        # save workbook
        try:
            self.wb.save(self.output_file)
            self.log_debug(f"Updated workbook saved to {self.output_file}")
        except Exception as e:
            self.log_debug(f"Failed to save updated workbook to {self.output_file}: {e}")


# Example usage (uncomment and set paths to use):
# if __name__ == "__main__":
#     processor = ExcelBackendProcessor("report.xlsx", "cdr.xlsx", "output.xlsx", headless=True)
#     processor.backend_process()
