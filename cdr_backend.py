# excel_processor.py
import logging
from concurrent.futures import ThreadPoolExecutor
from typing import Dict, Any, List, Tuple, Optional

import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment

logging.basicConfig(
    filename="excel_processor_debug.log",
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


class ExcelBackendProcessor:
    def __init__(
        self,
        report_file: str,
        cdr_file: str,
        output_file: str,
        headless: bool = False,
        log_callback: Optional[Any] = None,
    ):
        self.report_file = report_file
        self.cdr_file = cdr_file
        self.output_file = output_file
        self.headless = headless
        self.log_callback = log_callback

        # Data containers
        self.wb: Optional[Workbook] = None
        self.conn_df: Optional[pd.DataFrame] = None
        self.investor_entry_df: Optional[pd.DataFrame] = None
        self.cdr_summary_df: Optional[pd.DataFrame] = None
        self.lookup_df: Optional[pd.DataFrame] = None
        self.investor_allocation_df: Optional[pd.DataFrame] = None
        self.mapping: Dict[str, Any] = {}

        if self.headless:
            self.log_debug("Headless mode is enabled. Debug logging started.")

    def log_debug(self, message: str):
        logging.debug(message)
        if self.log_callback:
            try:
                self.log_callback("debug", message)
            except Exception:
                pass

    # ---- Normalizers ----
    @staticmethod
    def normalize_keys(series: pd.Series) -> pd.Series:
        return (
            series.astype(str)
            .str.upper()
            .str.replace(r"\s+", "", regex=True)
            .str.replace(r"[^0-9A-Z]", "", regex=True)
            .fillna("")
        )

    @staticmethod
    def normalize(name: str) -> str:
        try:
            nm = name.split("_")[-1]
        except Exception:
            nm = name
        return (
            str(nm)
            .strip()
            .upper()
            .replace(" ", "")
            .replace("_", "")
            .replace("/", "")
            .replace(".", "")
        )

    # ---- Formatting helpers ----
    def apply_theme(self, sheet):
        """Apply basic alignment, header/footer styling and zebra fill."""
        header_fill = PatternFill(start_color="093366", end_color="093366", fill_type="solid")
        footer_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        alt_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

        for r_idx, row in enumerate(sheet.iter_rows(values_only=False), start=1):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if cell.value in ("Blank1", "Blank2", None):
                    continue
                if r_idx == 1:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = header_fill
                elif r_idx == sheet.max_row:
                    cell.font = Font(bold=True)
                    cell.fill = footer_fill
                elif r_idx % 2 == 0:
                    cell.fill = alt_fill

    # ---- Load & Preprocess ----
    def load_data(self):
        self.log_debug("Loading data from Excel files.")
        # Adjust sheet names exactly as in your files
        self.cdr_summary_df = pd.read_excel(
            self.report_file, sheet_name="CDR Summary By Investor", skiprows=2, engine="openpyxl", dtype=object
        )
        self.investor_entry_df = pd.read_excel(
            self.cdr_file, sheet_name="investor_format", engine="openpyxl", dtype=object
        )
        # Try to load generic sheets; change indices/names as required
        self.conn_df = pd.read_excel(self.cdr_file, sheet_name=0, engine="openpyxl", dtype=object)
        self.conn_df = self.conn_df.loc[:, ~self.conn_df.columns.str.contains("^Unnamed")]
        try:
            self.investor_allocation_df = pd.read_excel(
                self.cdr_file, sheet_name="allocation_data", engine="openpyxl", dtype=object
            )
        except Exception:
            self.investor_allocation_df = pd.DataFrame()
        self.log_debug("Data loading completed.")

    def preprocess_data(self):
        self.log_debug("Starting data preprocessing.")
        # Normalization helpers
        normalize = lambda s: s.astype(str).str.upper().str.replace(r"\s+", "", regex=True)

        # Create normalized ID columns
        if "Investor ID" in self.investor_entry_df.columns:
            self.investor_entry_df["Normalized Investor ID"] = normalize(self.investor_entry_df["Investor ID"])
        if "Investran Acct ID" in self.conn_df.columns:
            self.conn_df["Normalized Investor ID"] = normalize(self.conn_df["Investran Acct ID"])

        # Example mapping: Bin ID and Commitment Amount from conn_df
        if {"Normalized Investor ID", "Bin ID"}.issubset(self.conn_df.columns):
            bin_id_map = dict(zip(self.conn_df["Normalized Investor ID"], self.conn_df["Bin ID"]))
            self.investor_entry_df["Bin ID"] = self.investor_entry_df.get("Normalized Investor ID", pd.Series()).map(bin_id_map)

        if {"Normalized Investor ID", "Commitment Amount"}.issubset(self.conn_df.columns):
            commitment_map = dict(zip(self.conn_df["Normalized Investor ID"], self.conn_df["Commitment Amount"]))
            self.investor_entry_df["Commitment"] = self.investor_entry_df.get("Normalized Investor ID", pd.Series()).map(commitment_map)

        # Normalize account numbers on CDR summary for joins
        acct_col = next((c for c in self.cdr_summary_df.columns if "Account" in str(c)), None)
        if acct_col:
            self.cdr_summary_df["Normalized Account Number"] = normalize(self.cdr_summary_df[acct_col].astype(str))

        # Filter and dedupe CDR summary (example)
        filtered = self.cdr_summary_df[~self.cdr_summary_df["Investor Name"].astype(str).str.contains("Total", na=False)]
        filtered = filtered.dropna(subset=[acct_col]).drop_duplicates(subset=["Normalized Account Number"])
        # TODO: split contribution/distribution/external columns as in your original code
        self.cdr_summary_df = filtered
        self.log_debug("Data preprocessing completed.")

    # ---- Sheets generation ----
    def update_commitment_sheet(self):
        self.log_debug("Updating Commitment sheet.")
        # Build lookup dataframe if available
        if self.cdr_summary_df is None or self.conn_df is None:
            self.log_debug("Missing dataframes for commitment update.")
            return
        # Example simple mapping: normalize lookup keys and create mapping
        try:
            lookup = self.cdr_summary_df[["Investor Name", "Account Number", "Investor ID", "Investor Commitment"]].copy()
        except Exception:
            lookup = pd.DataFrame()
        # Normalize and create mapping
        if not lookup.empty:
            lookup_keys = self.normalize_keys(lookup.iloc[:, 1])
            lookup_values = lookup.iloc[:, :3]
            self.mapping = dict(zip(lookup_keys, lookup_values.to_dict(orient="records")))
        # Example: add columns to conn_df
        self.conn_df["GS Commitment"] = self.conn_df.get("Bin ID", pd.Series()).map(self.mapping.get)
        # Create sheet and write dataframe
        ws = self.wb.create_sheet(title="Commitment")
        for r_idx, row in enumerate(dataframe_to_rows(self.conn_df.fillna(""), index=False, header=True), start=1):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)
        self.apply_theme(ws)
        self.log_debug("Commitment sheet updated.")

    def map_columns_to_summary(self, entry_columns: List[str], summary_df: pd.DataFrame) -> Dict[str, Any]:
        # Map entry columns to summary sums using normalized keys
        exclude_cols = ["Investor Name", "Investor ID", "Account Number", "Current Net Cashflow"]
        summary_df = summary_df.dropna(subset=["Account Number"], how="any")
        summary_sums = summary_df.drop(columns=[c for c in exclude_cols if c in summary_df.columns], errors="ignore").select_dtypes(include="number").sum()
        normalized_keys = {self.normalize(k): k for k in summary_sums.index}
        mapped_values = {}
        for col in entry_columns:
            norm_col = self.normalize(col)
            match = next((normalized_keys[k] for k in normalized_keys if (k in norm_col or norm_col in k)), None)
            mapped_values[col] = summary_sums.get(match, None)
        return mapped_values

    def update_entry_sheet(self):
        self.log_debug("Updating Entry sheet.")
        ws = self.wb.create_sheet(title="Entry")
        df = self.investor_entry_df.copy()
        # Remove zero-sum contribution/distribution/ext columns (example)
        prefixes = ["Contributions_", "Distributions_", "ExternalExpenses_"]
        cols_to_check = [col for col in df.columns if any(col.startswith(p) for p in prefixes)]
        cols_to_remove = [col for col in cols_to_check if pd.to_numeric(df[col], errors="coerce").fillna(0).sum() == 0]
        df.drop(columns=cols_to_remove, inplace=True, errors="ignore")

        for r_idx, row in enumerate(dataframe_to_rows(df.fillna(""), index=False, header=True), start=1):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        header_row = [cell.value for cell in ws[1]]
        # Add totals row
        sum_row_index = ws.max_row + 1
        ws.cell(row=sum_row_index, column=1, value="Total")
        for col_idx, col_name in enumerate(header_row, start=1):
            if col_name and (col_name == "Commitment" or any(col_name.startswith(p) for p in prefixes)):
                col_letter = get_column_letter(col_idx)
                formula = f"=SUM({col_letter}2:{col_letter}{sum_row_index-1})"
                ws.cell(row=sum_row_index, column=col_idx, value=formula)

        # Subtotals (example using mapping)
        summary_row_index = sum_row_index + 1
        ws.cell(row=summary_row_index, column=1, value="Subtotals")
        mapped_values = self.map_columns_to_summary(header_row, self.cdr_summary_df)
        for col_idx, col_name in enumerate(header_row, start=1):
            if col_name == "Commitment":
                try:
                    commitment_sum = pd.to_numeric(self.conn_df["Commitment Amount"][:-1], errors="coerce").sum()
                except Exception:
                    commitment_sum = None
                ws.cell(row=summary_row_index, column=col_idx, value=commitment_sum)
            elif col_name in mapped_values:
                ws.cell(row=summary_row_index, column=col_idx, value=mapped_values[col_name])

        # Validate row (simple equality check via formula)
        validate_row_index = summary_row_index + 1
        ws.cell(row=validate_row_index, column=1, value="Validate")
        for col_idx in range(2, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            total_cell = f"{col_letter}{sum_row_index}"
            subtotals_cell = f"{col_letter}{summary_row_index}"
            formula = f'=IF({total_cell}={subtotals_cell},"OK","Mismatch")'
            cell = ws.cell(row=validate_row_index, column=col_idx, value=formula)
            cell.alignment = Alignment(horizontal="center")

        # Apply fills and theme
        total_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        subtotals_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
        validate_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=sum_row_index, column=col_idx).fill = total_fill
            ws.cell(row=summary_row_index, column=col_idx).fill = subtotals_fill
            ws.cell(row=validate_row_index, column=col_idx).fill = validate_fill

        self.apply_theme(ws)
        self.log_debug("Entry sheet updated.")

    def update_approx_matches_sheet(self):
        self.log_debug("Updating Approx Matches sheet.")
        conn_keys_series = self.normalize_keys(self.conn_df.get("Bin ID", pd.Series()))
        lookup_keys = self.normalize_keys(self.lookup_df.iloc[:, 1]) if self.lookup_df is not None and not self.lookup_df.empty else pd.Series()
        approx_matches = []
        for key in conn_keys_series:
            found = next((lk for lk in lookup_keys if lk.startswith(key) or key in lk), None)
            if found and self.mapping.get(found):
                approx_matches.append({"Conn Key": key, "Matched Lookup Key": found, "Mapped Value": self.mapping.get(found)})

        if approx_matches:
            approx_df = pd.DataFrame(approx_matches)
            ws = self.wb.create_sheet(title="Approx Matches")
            for r_idx, row in enumerate(dataframe_to_rows(approx_df.fillna(""), index=False, header=True), start=1):
                for c_idx, value in enumerate(row, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=value)
            self.apply_theme(ws)
            self.log_debug("Approx Matches sheet updated.")

    # ---- Orchestration ----
    def backend_process(self):
        self.log_debug("Starting Excel processing workflow.")
        self.load_data()
        self.preprocess_data()

        # Load workbook (copy of report file to be modified)
        try:
            self.wb = load_workbook(filename=self.report_file)
        except Exception:
            self.wb = Workbook()

        # Remove existing target sheets if present
        for sheet in ["Commitment", "Approx Matches", "Entry"]:
            if sheet in self.wb.sheetnames:
                self.wb.remove(self.wb[sheet])

        # Run updates in threads (they share self.wb; openpyxl writes are done inside each function)
        with ThreadPoolExecutor(max_workers=3) as executor:
            futures = [
                executor.submit(self.update_commitment_sheet),
                executor.submit(self.update_entry_sheet),
                executor.submit(self.update_approx_matches_sheet),
            ]
            # Wait for completion (exceptions will surface here)
            for f in futures:
                try:
                    f.result()
                except Exception as e:
                    self.log_debug(f"Worker raised: {e}")

        # Save output
        self.wb.save(self.output_file)
        self.log_debug(f"Updated workbook saved to {self.output_file}")
