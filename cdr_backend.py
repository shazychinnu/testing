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
        # load CDR summary (adjust sheet_name if your file differs in exact naming)
        self.cdr_summary_df = pd.read_excel(
            self.report_file,
            sheet_name="CDR Summary By Investor",
            skiprows=2,
            engine="openpyxl",
            dtype=object
        )

        # Split by Subtotal marker rows (keeps original splitting behaviour)
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

        # Load other sheets from CDR file
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

    # -------------------- COMMITMENT SHEET (minimal-change, investor-subtotals) --------------------

    def create_commitment_sheet(self, conn_df, start_row=1):
        """
        Writes conn_df into a single Commitment sheet in the workbook.
        Inserts a subtotal row immediately after each Investor Name group,
        summing the 'Commitment Amount' (or other specified numeric columns if needed).
        This preserves row order from conn_df.
        """
        ws_name = "Commitment"
        ws = self.wb[ws_name] if ws_name in self.wb.sheetnames else self.wb.create_sheet(title=ws_name)
        row_offset = start_row

        # If conn_df is empty, just return
        if conn_df is None or conn_df.shape[0] == 0:
            return row_offset

        # Ensure relevant columns exist
        # Primary grouping key: 'Investor Name' (use exact column name from your source)
        group_col = "Investor Name"
        value_col = "Commitment Amount"  # column to sum for subtotal; adjust if needed

        # Write header once
        header = list(conn_df.columns)
        for c_idx, h in enumerate(header, start=1):
            ws.cell(row=row_offset, column=c_idx, value=h)
        header_row_idx = row_offset
        row_offset += 1

        # Iterate rows in order, write row-by-row, and detect group changes by Investor Name
        prev_investor = None
        running_sum = 0.0
        rows_written_for_group = 0

        for _, row in conn_df.iterrows():
            investor = row.get(group_col, "")
            # if first row or same investor:
            if prev_investor is None:
                prev_investor = investor

            # If investor changed, write subtotal for previous investor first
            if investor != prev_investor:
                # write subtotal row for prev_investor
                subtotal_row = ["Subtotal"] + [""] * (len(header) - 1)
                # place sum in the column corresponding to value_col if present
                if value_col in conn_df.columns:
                    idx = header.index(value_col)  # zero-based
                    subtotal_row[idx] = running_sum
                for c_idx, val in enumerate(subtotal_row, start=1):
                    ws.cell(row=row_offset, column=c_idx, value=val)
                # style subtotal fill immediately (optional)
                for c_idx in range(1, len(header) + 1):
                    ws.cell(row=row_offset, column=c_idx).fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
                row_offset += 1

                # reset running sum for new investor
                running_sum = 0.0
                rows_written_for_group = 0
                prev_investor = investor

            # Write the current row data
            for c_idx, col in enumerate(header, start=1):
                ws.cell(row=row_offset, column=c_idx, value=row.get(col))
            # Add to running sum if numeric
            if value_col in conn_df.columns:
                try:
                    val = float(row.get(value_col)) if row.get(value_col) not in (None, "") else 0.0
                except Exception:
                    val = 0.0
                running_sum += val

            row_offset += 1
            rows_written_for_group += 1

        # After loop, write subtotal for the last investor group
        if rows_written_for_group > 0:
            subtotal_row = ["Subtotal"] + [""] * (len(header) - 1)
            if value_col in conn_df.columns:
                idx = header.index(value_col)
                subtotal_row[idx] = running_sum
            for c_idx, val in enumerate(subtotal_row, start=1):
                ws.cell(row=row_offset, column=c_idx, value=val)
            for c_idx in range(1, len(header) + 1):
                ws.cell(row=row_offset, column=c_idx).fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
            row_offset += 1

        # Apply theme (header / alternate rows / last-row style)
        self.apply_theme(ws)
        return row_offset + 1  # small spacing after sheet content

    # -------------------- ENTRY & APPROX MATCHES (unchanged behavior) --------------------

    def create_entry_sheet(self, entry_df, start_row=1):
        ws_name = "Entry"
        ws = self.wb[ws_name] if ws_name in self.wb.sheetnames else self.wb.create_sheet(title=ws_name)
        row_offset = start_row

        if entry_df is None or entry_df.shape[0] == 0:
            return row_offset

        # Header
        header = list(entry_df.columns)
        for c_idx, h in enumerate(header, start=1):
            ws.cell(row=row_offset, column=c_idx, value=h)
        row_offset += 1

        # Write rows
        for _, row in entry_df.iterrows():
            for c_idx, col in enumerate(header, start=1):
                ws.cell(row=row_offset, column=c_idx, value=row.get(col))
            row_offset += 1

        # Add a subtotal row similar to earlier logic (if you used it before)
        subtotal_row = ["Subtotal"] + [""] * (len(header) - 1)
        # sum contributions/distributions/expenses if those columns exist
        for i, col in enumerate(header):
            if any(k in col for k in ["Contributions", "Distributions", "ExternalExpenses", "Expenses"]):
                subtotal_row[i] = pd.to_numeric(entry_df[col], errors='coerce').sum()
        for c_idx, val in enumerate(subtotal_row, start=1):
            ws.cell(row=row_offset, column=c_idx, value=val)
        for c_idx in range(1, len(header) + 1):
            ws.cell(row=row_offset, column=c_idx).fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        row_offset += 1

        self.apply_theme(ws)
        return row_offset + 1

    def create_approx_matches_sheet(self, conn_df, start_row=1):
        ws_name = "ApproxMatches"
        ws = self.wb[ws_name] if ws_name in self.wb.sheetnames else self.wb.create_sheet(title=ws_name)
        row_offset = start_row

        approx_df = pd.DataFrame({
            "Conn Key": conn_df.get("Bin ID", []),
            "Investor ID": conn_df.get("Investor Acct ID", [])
        })

        if approx_df.shape[0] == 0:
            return row_offset

        header = list(approx_df.columns)
        for c_idx, h in enumerate(header, start=1):
            ws.cell(row=row_offset, column=c_idx, value=h)
        row_offset += 1

        for _, row in approx_df.iterrows():
            for c_idx, col in enumerate(header, start=1):
                ws.cell(row=row_offset, column=c_idx, value=row.get(col))
            row_offset += 1

        subtotal_row = ["Subtotal"] + [""] * (len(header) - 1)
        for c_idx, val in enumerate(subtotal_row, start=1):
            ws.cell(row=row_offset, column=c_idx, value=val)
        for c_idx in range(1, len(header) + 1):
            ws.cell(row=row_offset, column=c_idx).fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

        self.apply_theme(ws)
        return row_offset + 2

    # -------------------- MAIN PROCESS FUNCTION --------------------

    def process_sections(self):
        self.load_data()
        # Load workbook (use report file as base) — create a new Workbook if you prefer
        try:
            self.wb = load_workbook(filename=self.report_file)
        except Exception:
            self.wb = Workbook()

        # Remove old sheets we will recreate
        for sheet_name in ["Commitment", "Entry", "ApproxMatches"]:
            if sheet_name in self.wb.sheetnames:
                del self.wb[sheet_name]

        self.log_debug("Processing sections...")

        # For Commitment sheet: we want to preserve conn_df ordering and insert subtotals per investor.
        # Use the raw connection dataframe exactly as source script did.
        # IMPORTANT: do not modify conn_df ordering here unless your source expects sorting.
        self.create_commitment_sheet(self.conn_df, start_row=1)

        # For Entry and ApproxMatches: write once using the loaded investor_entry_df and conn_df
        self.create_entry_sheet(self.investor_entry_df, start_row=1)
        self.create_approx_matches_sheet(self.conn_df, start_row=1)

        self.wb.save(self.output_file)
        self.log_debug(f"Workbook saved to {self.output_file}")


# Example usage:
# processor = ExcelBackendProcessor("report.xlsx", "cdr.xlsx", "output.xlsx")
# processor.process_sections()
