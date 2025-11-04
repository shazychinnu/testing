import os
import pandas as pd
import logging
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

class ExcelBackendProcessor:
    def __init__(self, report_file, cdr_file, output_file):
        self.report_file = report_file
        self.cdr_file = cdr_file
        self.output_file = output_file
        self.wb = None
        self.section_dfs = []
        self.investor_entry_df = None
        self.conn_df = None
        self.investor_allocation_df = None

        logging.basicConfig(
            filename='excel_processor_debug.log',
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.log_debug("Processor initialized.")

    def log_debug(self, message):
        logging.debug(message)
        print(f"[DEBUG] {message}")

    def normalize_keys(self, series):
        return series.astype(str).str.upper().str.replace(r'\s+', '', regex=True).fillna("")

    def normalize(self, name):
        return str(name).strip().upper().replace(" ", "").replace("_", "").replace("/", "").replace(".", "")

    def apply_theme(self, sheet):
        """Apply consistent styling to Excel sheet."""
        for r_idx, row in enumerate(sheet.iter_rows(), start=1):
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if r_idx == 1:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
                elif r_idx == sheet.max_row:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
                elif r_idx % 2 == 0:
                    cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    def load_data(self):
        """Load report and CDR files, split into multiple section DataFrames."""
        self.log_debug("Loading data...")
        # Load CDR summary (update name if it differs)
        self.cdr_summary_df = pd.read_excel(
            self.report_file,
            sheet_name="CDR Summary By Investor",
            skiprows=2,
            engine="openpyxl",
            dtype=object
        )

        # Split by subtotal markers
        subtotal_rows = self.cdr_summary_df[
            self.cdr_summary_df.iloc[:, 0].astype(str).str.startswith("Subtotal:", na=False)
        ].index.tolist()
        split_points = [0] + subtotal_rows + [len(self.cdr_summary_df)]

        for i in range(len(split_points) - 1):
            start, end = split_points[i], split_points[i + 1]
            part = self.cdr_summary_df.iloc[start:end].dropna(how='all')
            part = part[~part.iloc[:, 0].astype(str).str.startswith("Subtotal:", na=False)]
            if not part.empty:
                self.section_dfs.append(part)

        if not self.section_dfs:
            self.section_dfs = [self.cdr_summary_df]

        # Load other sheets
        self.investor_entry_df = pd.read_excel(
            self.cdr_file, sheet_name="investor_format", engine="openpyxl", dtype=object
        )
        self.conn_df = pd.read_excel(
            self.cdr_file, sheet_name=0, engine="openpyxl", dtype=object
        )
        self.investor_allocation_df = pd.read_excel(
            self.cdr_file, sheet_name="allocation_data", engine="openpyxl", dtype=object
        )
        self.log_debug("Data loaded successfully.")

    # -------------------- COMMITMENT SHEET FIX --------------------

    def create_commitment_sheet(self, conn_df, start_row=1):
        """Write conn_df to Commitment sheet with subtotals per investor."""
        ws_name = "Commitment"
        ws = self.wb[ws_name] if ws_name in self.wb.sheetnames else self.wb.create_sheet(title=ws_name)
        row_offset = start_row

        if conn_df is None or conn_df.shape[0] == 0:
            return row_offset

        header = list(conn_df.columns)
        for c_idx, h in enumerate(header, start=1):
            ws.cell(row=row_offset, column=c_idx, value=h)
        row_offset += 1

        group_col = "Investor Name"
        value_col = "Commitment Amount"

        prev_investor = None
        running_sum = 0.0

        for _, row in conn_df.iterrows():
            investor = row.get(group_col, "")
            if prev_investor is None:
                prev_investor = investor

            # subtotal when investor changes
            if investor != prev_investor:
                subtotal_row = ["Subtotal"] + [""] * (len(header) - 1)
                if value_col in header:
                    subtotal_row[header.index(value_col)] = running_sum
                for c_idx, val in enumerate(subtotal_row, start=1):
                    ws.cell(row=row_offset, column=c_idx, value=val)
                for c_idx in range(1, len(header) + 1):
                    ws.cell(row=row_offset, column=c_idx).fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
                row_offset += 1
                running_sum = 0.0
                prev_investor = investor

            for c_idx, col in enumerate(header, start=1):
                ws.cell(row=row_offset, column=c_idx, value=row.get(col))
            try:
                running_sum += float(row.get(value_col, 0) or 0)
            except Exception:
                pass
            row_offset += 1

        # subtotal for last investor
        subtotal_row = ["Subtotal"] + [""] * (len(header) - 1)
        if value_col in header:
            subtotal_row[header.index(value_col)] = running_sum
        for c_idx, val in enumerate(subtotal_row, start=1):
            ws.cell(row=row_offset, column=c_idx, value=val)
        for c_idx in range(1, len(header) + 1):
            ws.cell(row=row_offset, column=c_idx).fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        self.apply_theme(ws)
        return row_offset + 1

    # -------------------- ENTRY SHEET --------------------

    def create_entry_sheet(self, entry_df, start_row=1):
        ws_name = "Entry"
        ws = self.wb[ws_name] if ws_name in self.wb.sheetnames else self.wb.create_sheet(title=ws_name)
        row_offset = start_row

        if entry_df is None or entry_df.shape[0] == 0:
            return row_offset

        header = list(entry_df.columns)
        for c_idx, h in enumerate(header, start=1):
            ws.cell(row=row_offset, column=c_idx, value=h)
        row_offset += 1

        for _, row in entry_df.iterrows():
            for c_idx, col in enumerate(header, start=1):
                ws.cell(row=row_offset, column=c_idx, value=row.get(col))
            row_offset += 1

        subtotal_row = ["Subtotal"] + [""] * (len(header) - 1)
        for i, col in enumerate(header):
            if any(k in col for k in ["Contributions", "Distributions", "Expenses"]):
                subtotal_row[i] = pd.to_numeric(entry_df[col], errors='coerce').sum()
        for c_idx, val in enumerate(subtotal_row, start=1):
            ws.cell(row=row_offset, column=c_idx, value=val)
        for c_idx in range(1, len(header) + 1):
            ws.cell(row=row_offset, column=c_idx).fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        self.apply_theme(ws)
        return row_offset + 1

    # -------------------- FIXED APPROX MATCHES SHEET --------------------

    def create_approx_matches_sheet(self, conn_df, start_row=1):
        ws_name = "ApproxMatches"
        ws = self.wb[ws_name] if ws_name in self.wb.sheetnames else self.wb.create_sheet(title=ws_name)
        row_offset = start_row

        # Handle safe column extraction
        cols = []
        if "Bin ID" in conn_df.columns:
            cols.append("Bin ID")
        if "Investor Acct ID" in conn_df.columns:
            cols.append("Investor Acct ID")
        elif "Investor ID" in conn_df.columns:
            cols.append("Investor ID")

        if not cols:
            self.log_debug("No matching columns found for ApproxMatches — skipping.")
            return row_offset

        approx_df = conn_df[cols].copy()
        approx_df.columns = ["Conn Key", "Investor ID"][:len(approx_df.columns)]

        for r_idx, row in enumerate(dataframe_to_rows(approx_df, index=False, header=True), start=row_offset):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        self.apply_theme(ws)
        return row_offset + len(approx_df) + 2

    # -------------------- MAIN WORKFLOW --------------------

    def process_sections(self):
        self.load_data()
        try:
            self.wb = load_workbook(filename=self.report_file)
        except Exception:
            self.wb = Workbook()

        for sheet_name in ["Commitment", "Entry", "ApproxMatches"]:
            if sheet_name in self.wb.sheetnames:
                del self.wb[sheet_name]

        self.log_debug("Processing sections...")

        self.create_commitment_sheet(self.conn_df, start_row=1)
        self.create_entry_sheet(self.investor_entry_df, start_row=1)
        self.create_approx_matches_sheet(self.conn_df, start_row=1)

        self.wb.save(self.output_file)
        self.log_debug(f"Workbook saved to {self.output_file}")


# Example usage:
# processor = ExcelBackendProcessor("report.xlsx", "cdr.xlsx", "output.xlsx")
# processor.process_sections()
