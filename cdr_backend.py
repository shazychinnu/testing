import logging
import os
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows


class ExcelBackendProcessor:
    """
    Excel Backend Processor
    -----------------------
    Loads and merges Excel data from report and CDR files,
    normalizes investor and BIN IDs, and generates:
      • Commitment Sheet
      • Entry Sheet
      • Approx Matches Sheet
    """

    def __init__(
        self,
        report_file: str,
        cdr_file: str,
        output_file: str,
        headless: bool = False,
        log_callback=None,
    ):
        self.report_file = report_file
        self.cdr_file = cdr_file
        self.output_file = output_file
        self.headless = headless
        self.log_callback = log_callback

        self.wb = None
        self.conn_df = None
        self.investor_entry_df = None
        self.cdr_summary_df = None
        self.investor_allocation_df = None
        self.lookup_df = None
        self.mapping = {}

        if self.headless:
            logging.basicConfig(
                filename="excel_processor_debug.log",
                level=logging.DEBUG,
                format="%(asctime)s - %(levelname)s - %(message)s",
            )
            self.log_debug("Headless mode is enabled. Debug logging started.")

    # -----------------------------
    # Logging & normalization utils
    # -----------------------------
    def log_debug(self, message: str):
        if self.log_callback:
            try:
                self.log_callback("DEBUG", message)
            except Exception:
                pass
        logging.debug(message)

    def normalize_keys(self, series: pd.Series) -> pd.Series:
        return (
            series.astype(str)
            .str.upper()
            .str.replace(r"\s+", "", regex=True)
            .fillna("")
        )

    def normalize(self, name: str) -> str:
        return (
            str(name)
            .strip()
            .upper()
            .replace(" ", "")
            .replace("_", "")
            .replace("/", "")
            .replace(".", "")
        )

    def apply_theme(self, sheet):
        for r_idx, row in enumerate(sheet.iter_rows(), start=1):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if cell.value in ("Blank1", "Blank2", None):
                    continue
                if r_idx == 1:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(
                        start_color="003366", end_color="003366", fill_type="solid"
                    )
                elif r_idx == sheet.max_row:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(
                        start_color="FFFF99", end_color="FFFF99", fill_type="solid"
                    )
                elif r_idx % 2 == 0:
                    cell.fill = PatternFill(
                        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
                    )

    # ------------------
    # Step 1: Load data
    # ------------------
    def load_data(self):
        self.log_debug("Loading data from Excel files.")

        self.cdr_summary_df = pd.read_excel(
            self.report_file,
            sheet_name="CDR Summary By Investor",
            skiprows=2,
            engine="openpyxl",
            dtype=object,
        )

        self.investor_entry_df = pd.read_excel(
            self.cdr_file,
            sheet_name="investor_format",
            engine="openpyxl",
            dtype=object,
        )

        self.conn_df = (
            pd.read_excel(self.cdr_file, sheet_name=9, engine="openpyxl", dtype=object)
            .loc[:, lambda df: ~df.columns.str.contains("^Unnamed")]
            .copy()
        )

        self.investor_allocation_df = pd.read_excel(
            self.cdr_file,
            sheet_name="allocation_data",
            engine="openpyxl",
            dtype=object,
        )

        self.log_debug("Data loading completed.")

    # ------------------------
    # Step 2: Preprocess data
    # ------------------------
    def preprocess_data(self):
        self.log_debug("Starting data preprocessing.")

        normalize = lambda s: s.astype(str).str.upper().str.replace(r"\s+", "", regex=True)

        # Normalize IDs
        self.investor_entry_df["Normalized Investor ID"] = normalize(
            self.investor_entry_df["Investor ID"]
        )
        self.conn_df["Normalized Investor ID"] = normalize(self.conn_df["Investran Acct ID"])

        # Mapping BIN and Commitment
        bin_id_map = dict(
            zip(self.conn_df["Normalized Investor ID"], self.conn_df["Bin ID"])
        )
        commitment_map = dict(
            zip(self.conn_df["Normalized Investor ID"], self.conn_df["Commitment Amount"])
        )

        self.investor_entry_df["Bin ID"] = self.investor_entry_df[
            "Normalized Investor ID"
        ].map(bin_id_map)
        self.investor_entry_df["Commitment"] = self.investor_entry_df[
            "Normalized Investor ID"
        ].map(commitment_map)

        # Normalize Bin IDs
        self.investor_entry_df["Normalized Bin ID"] = normalize(
            self.investor_entry_df["Bin ID"]
        )
        self.cdr_summary_df["Normalized Account Number"] = normalize(
            self.cdr_summary_df["Account Number"]
        )

        # Drop totals and duplicates
        filtered = self.cdr_summary_df[
            ~self.cdr_summary_df["Investor Name"].astype(str).str.contains("Total", na=False)
        ].dropna(subset=["Account Number"])
        filtered = filtered.drop_duplicates(subset=["Normalized Account Number"])

        # Map contributions, distributions, external expenses
        contrib_cols = filtered.columns[9:17]
        dist_cols = filtered.columns[17:25]
        ext_cols = filtered.columns[25:33]

        # Merge into investor_entry_df
        for prefix, cols in {
            "Contributions_": contrib_cols,
            "Distributions_": dist_cols,
            "ExternalExpenses_": ext_cols,
        }.items():
            tmp = filtered[["Normalized Account Number"] + list(cols)].copy()
            tmp.rename(columns={"Normalized Account Number": "Normalized Bin ID"}, inplace=True)
            tmp.rename(columns={c: f"{prefix}{c}" for c in cols}, inplace=True)
            self.investor_entry_df = self.investor_entry_df.merge(
                tmp, on="Normalized Bin ID", how="left"
            )

        self.log_debug("Data preprocessing completed.")

    # ---------------------------------
    # Step 3: Update Commitment Sheet
    # ---------------------------------
    def update_commitment_sheet(self):
        self.log_debug("Updating Commitment sheet.")

        self.lookup_df = self.cdr_summary_df[
            ["Investor Name", "Account Number", "Investor ID", "Investor Commitment"]
        ]
        lookup_keys = self.normalize_keys(self.lookup_df["Account Number"])
        lookup_values = self.lookup_df["Investor Commitment"]
        conn_keys_series = self.normalize_keys(self.conn_df["Bin ID"])

        self.mapping = dict(zip(lookup_keys, lookup_values))

        def match_key(key):
            if key in self.mapping:
                return self.mapping[key], "-"
            for lk in lookup_keys:
                if lk.startswith(key):
                    return self.mapping[lk], "StartsWith"
                if key in lk:
                    return self.mapping[lk], "Contains"
            return None, "NA"

        matched_results = [match_key(key) for key in conn_keys_series]
        self.conn_df["GS Commitment"], match_types = zip(*matched_results)
        self.conn_df["GS Check"] = pd.to_numeric(
            self.conn_df["GS Commitment"], errors="coerce"
        ) - pd.to_numeric(self.conn_df["Commitment Amount"], errors="coerce")

        ws_commitment = self.wb.create_sheet(title="Commitment")
        for r_idx, row in enumerate(
            dataframe_to_rows(self.conn_df, index=False, header=True), start=1
        ):
            for c_idx, value in enumerate(row, start=1):
                ws_commitment.cell(row=r_idx, column=c_idx, value=value)
        self.apply_theme(ws_commitment)

        self.log_debug("Commitment sheet updated.")

    # ---------------------------------
    # Step 4: Update Entry Sheet
    # ---------------------------------
    def update_entry_sheet(self):
        ws_entry = self.wb.create_sheet(title="Entry")
        prefixes = ["Contributions_", "Distributions_", "ExternalExpenses_"]
        cols_to_check = [
            col for col in self.investor_entry_df.columns if any(col.startswith(p) for p in prefixes)
        ]
        cols_to_remove = [
            col
            for col in cols_to_check
            if pd.to_numeric(self.investor_entry_df[col], errors="coerce").fillna(0).sum() == 0
        ]
        self.investor_entry_df.drop(columns=cols_to_remove, inplace=True, errors="ignore")

        for r_idx, row in enumerate(
            dataframe_to_rows(self.investor_entry_df, index=False, header=True), start=1
        ):
            for c_idx, value in enumerate(row, start=1):
                ws_entry.cell(row=r_idx, column=c_idx, value=value)

        self.apply_theme(ws_entry)
        self.log_debug("Entry sheet updated.")

    # ---------------------------------
    # Step 5: Update Approx Matches
    # ---------------------------------
    def update_approx_matches_sheet(self):
        self.log_debug("Updating Approx Matches sheet.")
        conn_keys_series = self.normalize_keys(self.conn_df["Bin ID"])
        lookup_keys = self.normalize_keys(self.lookup_df["Account Number"])

        approx_matches = []
        for key in conn_keys_series:
            for lk in lookup_keys:
                if lk.startswith(key) or key in lk:
                    approx_matches.append(
                        {
                            "Conn Key": key,
                            "Matched Lookup Key": lk,
                            "Mapped Value": self.mapping.get(lk),
                        }
                    )
                    break

        if approx_matches:
            approx_df = pd.DataFrame(approx_matches)
            ws_approx = self.wb.create_sheet(title="Approx Matches")
            for r_idx, row in enumerate(
                dataframe_to_rows(approx_df, index=False, header=True), start=1
            ):
                for c_idx, value in enumerate(row, start=1):
                    ws_approx.cell(row=r_idx, column=c_idx, value=value)
            self.apply_theme(ws_approx)

        self.log_debug("Approx Matches sheet updated.")

    # ---------------------------------
    # Step 6: Run Full Processing
    # ---------------------------------
    def backend_process(self):
        self.log_debug("Starting Excel processing workflow.")
        self.load_data()
        self.preprocess_data()
        self.wb = load_workbook(filename=self.report_file)

        for sheet in ["Commitment", "Entry", "Approx Matches"]:
            if sheet in self.wb.sheetnames:
                self.wb.remove(self.wb[sheet])

        with ThreadPoolExecutor() as executor:
            executor.submit(self.update_commitment_sheet)
            executor.submit(self.update_entry_sheet)
            executor.submit(self.update_approx_matches_sheet)

        self.wb.save(self.output_file)
        self.log_debug(f"Updated workbook saved to {self.output_file}")
