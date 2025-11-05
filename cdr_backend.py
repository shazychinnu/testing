import os
import pandas as pd
import logging
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from concurrent.futures import ThreadPoolExecutor


class ExcelBackendProcessor:
    def __init__(self, report_file, cdr_file, output_file):
        self.report_file = report_file
        self.cdr_file = cdr_file
        self.output_file = output_file
        self.wb = None
        self.section_dfs = []
        self.investern_entry_df = None
        self.conn_df = None
        self.investern_allocation_df = None

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
        """Apply styling to Excel sheet."""
        for r_idx, row in enumerate(sheet.iter_rows(), start=1):
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical="center")
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
            self.cdr_file,
            sheet_name="CDR Summary By Investor",
            skiprows=2,
            engine="openpyxl",
            dtype=object
        )

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

        self.investern_entry_df = pd.read_excel(self.report_file, sheet_name="investern_format", engine="openpyxl", dtype=object)
        self.conn_df = pd.read_excel(self.report_file, sheet_name=0, engine="openpyxl", dtype=object)
        self.conn_df = self.conn_df.loc[:, lambda df: df.columns.str.contains("^Unn")]
        self.investern_allocation_df = pd.read_excel(self.report_file, sheet_name="allocation_data", engine="openpyxl", dtype=object)

        self.log_debug("Data loaded successfully.")

    def process_section(self, section_df, index):
        """Dynamic function to process each section."""
        self.log_debug(f"Processing section {index}...")

        normalize = lambda s: s.astype(str).str.upper().str.replace(r'\s+', '', regex=True)
        entry_df = self.investern_entry_df.copy()
        conn_df = self.conn_df.copy()

        entry_df['Normalized Investor ID'] = normalize(entry_df['Investor ID'])
        conn_df['Normalized Investor ID'] = normalize(conn_df['Investran Acct ID'])

        bin_id_map = dict(zip(conn_df['Normalized Investor ID'], conn_df['Bin ID']))
        entry_df['Bin ID'] = entry_df['Normalized Investor ID'].map(bin_id_map)

        commitment_map = dict(zip(conn_df['Normalized Investor ID'], conn_df['Commitment Amount']))
        entry_df['Commitment'] = entry_df['Normalized Investor ID'].map(commitment_map)

        contrib_cols = [c for c in section_df.columns if "PI" in str(c)]
        contrib_cols = ["Contributions_", "Distributions_", "ExternalExpenses_"]

        temp_df = section_df[['Normalized Account Number'] + contrib_cols].copy()
        temp_df.rename(columns={'Normalized Account Number': 'Normalized Bin ID'}, inplace=True)
        temp_df.rename(columns={col: f"prefix_{col}" for col in contrib_cols}, inplace=True)
        entry_df = entry_df.merge(temp_df, on='Normalized Bin ID', how='left')

        net_cashflow_map = dict(zip(section_df['Account Number'], section_df['Current Net Cashflow']))
        final_amount_map = dict(zip(self.investern_allocation_df['Investor ID'], self.investern_allocation_df['Final LE Amount']))

        entry_df["CRM - Final LE Amount"] = entry_df["Investor ID"].map(final_amount_map)
        entry_df["Current Net Cashflow"] = entry_df["Bin ID"].map(net_cashflow_map)
        entry_df["Current Net Cashflow Validation"] = entry_df["Current Net Cashflow"] + entry_df["CRM - Final LE Amount"]

        entry_df.drop(columns=['Normalized Investor ID', 'Normalized Bin ID'], inplace=True)

        with ThreadPoolExecutor() as executor:
            executor.submit(self.create_commitment_sheet, conn_df, section_df, index)
            executor.submit(self.create_entry_sheet, entry_df, section_df, index)
            executor.submit(self.create_approx_matches_sheet, conn_df, section_df, index)

    def create_commitment_sheet(self, conn_df, section_df, index):
        """Commitment sheet logic with subtotal and validation rows."""
        ws = self.wb.create_sheet(title=f"Commitment_{index}")

        lookup_df = section_df[['Investor Name', 'Account Number', 'Investor ID', 'Investor Commitment']]
        lookup_keys = self.normalize_keys(lookup_df.iloc[:, 1])
        lookup_values = lookup_df.iloc[:, 3]
        conn_keys_series = self.normalize_keys(conn_df['Bin ID'])
        mapping = dict(zip(lookup_keys, lookup_values))

        def match_key(key):
            if key in mapping:
                return mapping[key], ""
            for lk in lookup_keys:
                if lk.startswith(key):
                    return mapping[lk], "StartsWith"
                if key in lk:
                    return mapping[lk], "Contains"
            return None, "NA"

        matched_results = [match_key(key) for key in conn_keys_series]
        conn_df['GS Commitment'], match_types = zip(*matched_results)
        conn_df['GS Check'] = pd.to_numeric(conn_df['GS Commitment'], errors='coerce') - pd.to_numeric(conn_df['Commitment Amount'], errors='coerce')

        subtotal_row = ['Subtotal'] + [""] * (conn_df.shape[1] - 1)
        for col in ['Commitment Amount', 'GS Commitment', 'Investor Commitment', 'GS Check']:
            subtotal_row[conn_df.columns.get_loc(col)] = pd.to_numeric(conn_df[col], errors='coerce').sum()
        conn_df.loc[len(conn_df)] = subtotal_row

        for r_idx, row in enumerate(dataframe_to_rows(conn_df, index=False, header=True), start=1):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        total_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=ws.max_row, column=col_idx).fill = total_fill

        self.apply_theme(ws)

    def create_entry_sheet(self, entry_df, section_df, index):
        """Entry sheet logic with totals, subtotals, validation."""
        ws = self.wb.create_sheet(title=f"Entry_{index}")

        for r_idx, row in enumerate(dataframe_to_rows(entry_df, index=False, header=True), start=1):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        header_row = [cell.value for cell in ws[1]]
        sum_row_index = ws.max_row + 1
        ws.cell(row=sum_row_index, column=1, value="Total")

        for col_idx, col_name in enumerate(header_row, start=1):
            if col_name and (col_name == "Commitment" or col_name.startswith(("Contributions_", "Distributions_", "ExternalExpenses_"))):
                col_letter = get_column_letter(col_idx)
                formula = f"=SUM({col_letter}2:{col_letter}{sum_row_index - 1})"
                ws.cell(row=sum_row_index, column=col_idx, value=formula)

        summary_row_index = sum_row_index + 1
        ws.cell(row=summary_row_index, column=1, value="Subtotals")
        for col_idx, col_name in enumerate(header_row, start=1):
            if col_name == "Commitment":
                ws.cell(row=summary_row_index, column=col_idx, value=entry_df["Commitment"].sum())

        validate_row_index = summary_row_index + 1
        ws.cell(row=validate_row_index, column=1, value="Validate")

        for col_idx in range(2, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            formula = f'=IF({col_letter}{sum_row_index}={col_letter}{summary_row_index},"OK","Mismatch")'
            ws.cell(row=validate_row_index, column=col_idx, value=formula)

        self.apply_theme(ws)

    def create_approx_matches_sheet(self, conn_df, section_df, index):
        """Approx Matches sheet logic."""
        ws = self.wb.create_sheet(title=f"ApproxMatches_{index}")
        approx_df = pd.DataFrame({
            "Conn Key": conn_df['Bin ID'],
            "Investor ID": conn_df['Investran Acct ID']
        })
        for r_idx, row in enumerate(dataframe_to_rows(approx_df, index=False, header=True), start=1):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        self.apply_theme(ws)

    def backend_process(self):
        """Main workflow: load data, clean old sheets, process each section."""
        self.load_data()
        self.wb = load_workbook(filename=self.report_file)

        for sheet in self.wb.sheetnames:
            if sheet.startswith(("Commitment", "Entry", "ApproxMatches")):
                self.wb.remove(self.wb[sheet])

        for idx, section_df in enumerate(self.section_dfs, start=1):
            self.process_section(section_df, idx)

        self.wb.save(self.output_file)
        self.log_debug(f"Workbook saved to {self.output_file}")


# Example usage:
# processor = ExcelBackendProcessor(report_file, cdr_file, output_file)
# processor.backend_process()
