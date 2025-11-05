import os
import pandas as pd
import logging
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

class ExcelBackendProcessor:
    def __init__(self, report_file, cdr_file, output_file):
        # keep same signature as your source
        self.report_file = report_file
        self.cdr_file = cdr_file
        self.output_file = output_file

        # workbook and data placeholders
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
        # remove whitespace and uppercase to create normalized keys
        return series.astype(str).str.upper().str.replace(r'\s+', '', regex=True).fillna("")

    def normalize(self, name):
        return str(name).strip().upper().replace(" ", "").replace("_", "").replace("/", "").replace(".", "")

    def apply_theme(self, sheet):
        """Apply header, alternating row color, and last-row highlight."""
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
        """Load report and CDR files and split the CDR summary into sections (same behavior as source)."""
        self.log_debug("Loading data...")

        # Read CDR summary sheet (ensure spelling matches your file)
        self.cdr_summary_df = pd.read_excel(
            self.report_file,
            sheet_name="CDR Summary By Investor",
            skiprows=2,
            engine="openpyxl",
            dtype=object
        )

        # Split by Subtotal marker rows (keeps original behavior)
        subtotal_rows = self.cdr_summary_df[
            self.cdr_summary_df.iloc[:, 0].astype(str).str.startswith("Subtotal:", na=False)
        ].index.tolist()
        split_points = [0] + subtotal_rows + [len(self.cdr_summary_df)]

        for i in range(len(split_points) - 1):
            start, end = split_points[i], split_points[i + 1]
            part = self.cdr_summary_df.iloc[start:end].dropna(how='all')
            # remove lines that start with Subtotal: if present
            part = part[~part.iloc[:, 0].astype(str).str.startswith("Subtotal:", na=False)]
            if not part.empty:
                # reset index for each part to ensure simple iteration later
                part = part.reset_index(drop=True)
                self.section_dfs.append(part)

        if not self.section_dfs:
            # fallback: keep whole df as single section
            self.section_dfs = [self.cdr_summary_df.reset_index(drop=True)]

        # Load other sheets from CDR file (same as your source)
        self.investor_entry_df = pd.read_excel(
            self.cdr_file, sheet_name="investor_format", engine="openpyxl", dtype=object
        )
        self.conn_df = pd.read_excel(
            self.cdr_file, sheet_name=0, engine="openpyxl", dtype=object
        )
        self.investor_allocation_df = pd.read_excel(
            self.cdr_file, sheet_name="allocation_data", engine="openpyxl", dtype=object
        )

        self.log_debug(f"Loaded {len(self.section_dfs)} section(s) from CDR summary.")
        # Insert subtotal rows into each section based on Legal Entity
        self._insert_subtotals_into_sections(group_col="Legal Entity", sum_col="Commitment Amount", subtotal_label="Subtotal")
        self.log_debug("Data loaded and subtotals inserted into sections successfully.")

    # -------------------- Subtotal insertion (new) --------------------
    def _insert_subtotals_into_sections(self, group_col='Legal Entity', sum_col='Commitment Amount', subtotal_label='Subtotal'):
        """
        For each section in self.section_dfs, insert subtotal rows after each group of group_col.
        Also add a boolean column 'IsSubtotal' to mark inserted subtotal rows.
        The subtotal value is placed in sum_col column.
        """
        updated_sections = []
        for sec_idx, sec_df in enumerate(self.section_dfs, start=1):
            # Ensure sec_df is DataFrame
            sec = sec_df.copy().reset_index(drop=True)

            # If group_col not present, skip insertion for this section (but add IsSubtotal column)
            if group_col not in sec.columns or sum_col not in sec.columns:
                sec['IsSubtotal'] = False
                updated_sections.append(sec)
                continue

            new_rows = []
            prev_group = None
            running_sum = 0.0
            rows_in_group = 0

            # iterate rows in original order
            for _, row in sec.iterrows():
                current_group = row.get(group_col, "")
                # if new group starts (and not the first row), append subtotal row for prev_group
                if prev_group is not None and current_group != prev_group:
                    # construct subtotal row: place label in first column, sum in sum_col
                    subtotal_row = {c: "" for c in sec.columns}
                    subtotal_row[list(sec.columns)[0]] = subtotal_label
                    # put numeric sum in sum_col
                    subtotal_row[sum_col] = running_sum
                    subtotal_row['IsSubtotal'] = True
                    new_rows.append(subtotal_row)
                    # reset running_sum
                    running_sum = 0.0
                    rows_in_group = 0

                # append the actual row
                row_dict = row.to_dict()
                row_dict['IsSubtotal'] = False
                # safe numeric addition
                try:
                    val = row_dict.get(sum_col)
                    # handle strings with commas or empty values
                    if pd.isna(val) or val == "":
                        val_num = 0.0
                    else:
                        val_num = float(str(val).replace(',', ''))
                except Exception:
                    val_num = 0.0
                running_sum += val_num
                rows_in_group += 1
                new_rows.append(row_dict)
                prev_group = current_group

            # after loop add subtotal for last group if rows exist
            if rows_in_group > 0:
                subtotal_row = {c: "" for c in sec.columns}
                subtotal_row[list(sec.columns)[0]] = subtotal_label
                subtotal_row[sum_col] = running_sum
                subtotal_row['IsSubtotal'] = True
                new_rows.append(subtotal_row)

            # create DataFrame from new_rows preserving columns order and ensuring IsSubtotal exists
            new_sec_df = pd.DataFrame(new_rows, columns=list(sec.columns) + (['IsSubtotal'] if 'IsSubtotal' not in sec.columns else []))
            # If 'IsSubtotal' got duplicated in columns list above, fix:
            if 'IsSubtotal' not in new_sec_df.columns:
                new_sec_df['IsSubtotal'] = [r.get('IsSubtotal', False) for r in new_rows]

            # Ensure sum_col dtype numeric where possible
            if sum_col in new_sec_df.columns:
                new_sec_df[sum_col] = pd.to_numeric(new_sec_df[sum_col], errors='coerce')

            updated_sections.append(new_sec_df)

            self.log_debug(f"Section {sec_idx}: inserted subtotals grouped by '{group_col}' into section (rows -> {len(new_sec_df)})")

        # replace section_dfs with updated versions
        self.section_dfs = updated_sections

    # -------------------- Commitment sheet (write sections with subtotals) --------------------
    def create_commitment_sheet(self, start_row=1):
        """
        Writes every section in self.section_dfs (which already contains inserted subtotal rows)
        to a single 'Commitment' sheet in the workbook. Preserves section order and row order.
        """
        ws_name = "Commitment"
        ws = self.wb.create_sheet(title=ws_name)
        row_offset = start_row

        if not self.section_dfs:
            self.log_debug("No section dfs to write to Commitment sheet.")
            return row_offset

        # Determine header: union of columns across sections to be safe,
        # but prefer columns from first section (so output matches original layout)
        first_sec = self.section_dfs[0]
        header = list(first_sec.columns)
        # If IsSubtotal was added, keep it as last column (optional)
        if 'IsSubtotal' in header and header[-1] != 'IsSubtotal':
            header = [c for c in header if c != 'IsSubtotal'] + ['IsSubtotal']

        # write header
        for c_idx, h in enumerate(header, start=1):
            ws.cell(row=row_offset, column=c_idx, value=h)
        row_offset += 1

        # Write sections sequentially
        for sec_idx, sec in enumerate(self.section_dfs, start=1):
            if sec is None or sec.shape[0] == 0:
                continue
            # ensure sec columns include header cols
            for _, row in sec.iterrows():
                for c_idx, col in enumerate(header, start=1):
                    # Use empty string for missing columns
                    val = row.get(col, "")
                    ws.cell(row=row_offset, column=c_idx, value=val)
                row_offset += 1
            # optional blank line between sections for readability
            # (do not change row_offset if you must exactly match original; comment out next two lines)
            # row_offset += 0

        # Apply styling
        self.apply_theme(ws)
        return row_offset

    # -------------------- Entry sheet (unchanged from source) --------------------
    def create_entry_sheet(self, entry_df, start_row=1):
        ws_name = "Entry"
        ws = self.wb.create_sheet(title=ws_name)
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

        # Insert subtotal row similar to source (if required)
        subtotal_row = ["Subtotal"] + [""] * (len(header) - 1)
        for i, col in enumerate(header):
            if any(k in col for k in ["Contributions", "Distributions", "Expenses", "ExternalExpenses"]):
                subtotal_row[i] = pd.to_numeric(entry_df[col], errors='coerce').sum()
        for c_idx, val in enumerate(subtotal_row, start=1):
            ws.cell(row=row_offset, column=c_idx, value=val)
        for c_idx in range(1, len(header) + 1):
            ws.cell(row=row_offset, column=c_idx).fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

        self.apply_theme(ws)
        return row_offset + 1

    # -------------------- ApproxMatches (safe extraction) --------------------
    def create_approx_matches_sheet(self, conn_df, start_row=1):
        ws_name = "ApproxMatches"
        ws = self.wb.create_sheet(title=ws_name)
        row_offset = start_row

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

    # -------------------- Main workflow (use report_file as base, save to output_file) --------------------
    def process_sections(self):
        # load cdr summary and other required data, and insert subtotals into section_dfs
        self.load_data()

        # load report_file as base workbook (preserve other sheets)
        self.wb = load_workbook(filename=self.report_file)
        self.log_debug(f"Loaded workbook from {self.report_file}")

        # remove only the sheets we will recreate (keep other sheets intact)
        for sheet_name in ["Commitment", "Entry", "ApproxMatches"]:
            if sheet_name in self.wb.sheetnames:
                del self.wb[sheet_name]

        self.log_debug("Writing updated sheets to workbook...")

        # write Commitment using the updated section_dfs (subtotals already inserted)
        self.create_commitment_sheet(start_row=1)

        # write Entry and ApproxMatches as in source
        self.create_entry_sheet(self.investor_entry_df, start_row=1)
        self.create_approx_matches_sheet(self.conn_df, start_row=1)

        # save out to output_file (creates or overwrites)
        self.wb.save(self.output_file)
        self.log_debug(f"Workbook saved to {self.output_file}")


# Example usage:
# processor = ExcelBackendProcessor("report.xlsx", "cdr.xlsx", "output.xlsx")
# processor.process_sections()
