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
        self.cdr_summary_df = pd.read_excel(
            self.report_file,
            sheet_name="CDR Summary By Investor",
            skiprows=2,
            engine="openpyxl",
            dtype=object
        )

        # Split by Subtotal rows
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

    # -------------------- SHEET CREATION FUNCTIONS --------------------

    def create_commitment_sheet(self, conn_df, section_df, start_row=1):
        ws_name = "Commitment"
        ws = self.wb[ws_name] if ws_name in self.wb.sheetnames else self.wb.create_sheet(title=ws_name)
        row_offset = start_row

        # Write section
        for r_idx, row in enumerate(dataframe_to_rows(conn_df, index=False, header=True), start=0):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=row_offset + r_idx, column=c_idx, value=value)
        row_offset += len(conn_df)

        # Subtotal row
        subtotal_row = ['Subtotal'] + [""] * (conn_df.shape[1] - 1)
        for col in ['Commitment Amount', 'Investor Commitment']:
            if col in conn_df.columns:
                subtotal_row[conn_df.columns.get_loc(col)] = pd.to_numeric(conn_df[col], errors='coerce').sum()
        for c_idx, value in enumerate(subtotal_row, start=1):
            ws.cell(row=row_offset + 1, column=c_idx, value=value)

        # Style subtotal
        total_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=row_offset + 1, column=col_idx).fill = total_fill

        self.apply_theme(ws)
        return row_offset + 3

    def create_entry_sheet(self, entry_df, section_df, start_row=1):
        ws_name = "Entry"
        ws = self.wb[ws_name] if ws_name in self.wb.sheetnames else self.wb.create_sheet(title=ws_name)
        row_offset = start_row

        for r_idx, row in enumerate(dataframe_to_rows(entry_df, index=False, header=True), start=0):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=row_offset + r_idx, column=c_idx, value=value)
        row_offset += len(entry_df)

        subtotal_row = ['Subtotal'] + [""] * (entry_df.shape[1] - 1)
        for col in entry_df.columns:
            if any(x in col for x in ["Contributions", "Distributions", "Expenses"]):
                subtotal_row[entry_df.columns.get_loc(col)] = pd.to_numeric(entry_df[col], errors='coerce').sum()
        for c_idx, value in enumerate(subtotal_row, start=1):
            ws.cell(row=row_offset + 1, column=c_idx, value=value)

        total_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=row_offset + 1, column=col_idx).fill = total_fill

        self.apply_theme(ws)
        return row_offset + 3

    def create_approx_matches_sheet(self, conn_df, section_df, start_row=1):
        ws_name = "ApproxMatches"
        ws = self.wb[ws_name] if ws_name in self.wb.sheetnames else self.wb.create_sheet(title=ws_name)
        row_offset = start_row

        approx_df = pd.DataFrame({
            "Conn Key": conn_df.get("Bin ID", []),
            "Investor ID": conn_df.get("Investor Acct ID", [])
        })

        for r_idx, row in enumerate(dataframe_to_rows(approx_df, index=False, header=True), start=0):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=row_offset + r_idx, column=c_idx, value=value)
        row_offset += len(approx_df)

        subtotal_row = ["Subtotal"] + [""] * (approx_df.shape[1] - 1)
        for c_idx, value in enumerate(subtotal_row, start=1):
            ws.cell(row=row_offset + 1, column=c_idx, value=value)

        total_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=row_offset + 1, column=col_idx).fill = total_fill

        self.apply_theme(ws)
        return row_offset + 3

    # -------------------- MAIN PROCESS FUNCTION --------------------

    def process_sections(self):
        self.load_data()
        self.wb = load_workbook(filename=self.report_file)

        # Remove old sheets
        for sheet_name in ["Commitment", "Entry", "ApproxMatches"]:
            if sheet_name in self.wb.sheetnames:
                del self.wb[sheet_name]

        self.log_debug("Processing sections...")
        row_offsets = {"Commitment": 1, "Entry": 1, "ApproxMatches": 1}

        for idx, section_df in enumerate(self.section_dfs, start=1):
            self.log_debug(f"Processing section {idx}")
            row_offsets["Commitment"] = self.create_commitment_sheet(self.conn_df, section_df, start_row=row_offsets["Commitment"])
            row_offsets["Entry"] = self.create_entry_sheet(self.investor_entry_df, section_df, start_row=row_offsets["Entry"])
            row_offsets["ApproxMatches"] = self.create_approx_matches_sheet(self.conn_df, section_df, start_row=row_offsets["ApproxMatches"])

        self.wb.save(self.output_file)
        self.log_debug(f"Workbook saved to {self.output_file}")


# Example usage:
# processor = ExcelBackendProcessor("report.xlsx", "cdr.xlsx", "output.xlsx")
# processor.process_sections()
