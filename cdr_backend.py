import pandas as pd
from openpyxl import load_workbook, Workbook


cdr_file = "CDR_VREP.xlsx"
wizard_file = "report_file.xlsx"
output_file = "output.xlsx"


def ensure_output_file_exists():
    try:
        with open(output_file, 'rb'):
            pass
    except FileNotFoundError:
        wb = load_workbook(filename=cdr_file)
        wb.save(output_file)


def create_commitment_sheet():
    ensure_output_file_exists()

    cdr_summary_df = pd.read_excel(
        cdr_file,
        sheet_name='CDR Summary By Investor',
        engine='openpyxl',
        skiprows=2
    )
    cdr_summary_df.columns = cdr_summary_df.columns.str.strip()
    cdr_summary_df['Account Number'] = cdr_summary_df['Account Number'].astype(str)
    cdr_summary_df = cdr_summary_df.drop_duplicates(subset=['Investor ID'])

    account_to_commitment = cdr_summary_df.set_index('Account Number')['Investor Commitment'].to_dict()
    investorid_to_commitment = cdr_summary_df.set_index('Investor ID')['Investor Commitment'].to_dict()

    df = pd.read_excel(wizard_file, sheet_name='Data_format', engine='openpyxl')
    df.columns = df.columns.str.strip()
    df['Bin ID'] = df['Bin ID'].astype(str)
    df['Commitment Amount'] = pd.to_numeric(df['Commitment Amount'], errors='coerce')
    df['GS Commitment'] = df['Bin ID'].map(account_to_commitment)

    subtotal_idx = df['Legal Entity Name'].astype(str).str.contains("Subtotal", case=False, na=False)
    subtotal_indices = df.index[subtotal_idx].tolist()
    start_idx = 0

    for idx in subtotal_indices:
        block = df.iloc[start_idx:idx]
        commitment_sum = block['Commitment Amount'].sum()
        gs_commitment_sum = block['GS Commitment'].sum()

        df.at[idx, 'Commitment Amount'] = commitment_sum
        df.at[idx, 'GS Commitment'] = gs_commitment_sum
        start_idx = idx + 1

    df['GS Check'] = df['Commitment Amount'] - df['GS Commitment']

    investern_df = pd.read_excel(wizard_file, sheet_name='Investern_format', engine='openpyxl')
    investern_df.columns = investern_df.columns.str.strip()
    investern_df['Investor Id'] = investern_df['Investor Id'].astype(str)

    investern_df['SS Commitment'] = investern_df['Investor Id'].map(investorid_to_commitment)
    investern_df['SS Check'] = pd.to_numeric(investern_df['SS Commitment'], errors='coerce') - pd.to_numeric(
        investern_df['Invester Commitment'], errors='coerce'
    )

    max_rows = max(len(df), len(investern_df))
    empty_columns = pd.DataFrame({f'Empty_{i}': [''] * max_rows for i in range(3)})

    df = df.reindex(range(max_rows)).reset_index(drop=True)
    investern_df = investern_df.reindex(range(max_rows)).reset_index(drop=True)
    empty_columns = empty_columns.reindex(range(max_rows)).reset_index(drop=True)

    combined_df = pd.concat([df, empty_columns, investern_df], axis=1)

    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        combined_df.to_excel(writer, sheet_name='Commitment Sheet', index=False)


def create_entry_sheet_with_subtotals():
    df_raw = pd.read_excel(wizard_file, sheet_name='allocation_data', engine='openpyxl', header=None)

    header_rows = df_raw.index[df_raw.iloc[:, 0] == "Vehicle/Investor"].tolist()
    tables = []

    for i, header_idx in enumerate(header_rows):
        start_idx = header_idx
        end_idx = header_rows[i + 1] if i + 1 < len(header_rows) else len(df_raw)
        block = df_raw.iloc[start_idx:end_idx].reset_index(drop=True)
        block.columns = block.iloc[0]
        block = block.drop(0).reset_index(drop=True)

        numeric_cols = ['Final LE Amount']
        subtotal = block[numeric_cols].sum()

        subtotal_row = {
            col: subtotal[col] if col in numeric_cols else ('Subtotal' if col == block.columns[0] else '')
            for col in block.columns
        }
        block = pd.concat([block, pd.DataFrame([subtotal_row])], ignore_index=True)
        tables.append(block)

    processed_tables = []
    for df in tables:
        if set(df.iloc[0]) == set(df.columns):
            df = df.iloc[1:]
        processed_tables.append(df)

    final_df = pd.concat(processed_tables, ignore_index=True)

    commitment_df = pd.read_excel(output_file, sheet_name='Commitment Sheet', engine='openpyxl')
    commitment_mapping = commitment_df[['Investor Acct ID', 'Bin ID', 'Commitment Amount']].drop_duplicates(
        subset=['Investor Acct ID']
    )

    investorid_to_binid = commitment_mapping.set_index('Investor Acct ID')['Bin ID'].to_dict()
    investorid_to_commitment = commitment_mapping.set_index('Investor Acct ID')['Commitment Amount'].to_dict()

    final_df['Bin ID'] = final_df['Investor ID'].map(investorid_to_binid)
    final_df['Commitment Amount'] = final_df['Investor ID'].map(investorid_to_commitment)

    try:
        with open(output_file, 'rb'):
            pass
    except FileNotFoundError:
        wb = Workbook()
        wb.save(output_file)

    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        final_df.to_excel(writer, sheet_name='Entry', index=False)


# Example usage
if __name__ == "__main__":
    create_commitment_sheet()
    create_entry_sheet_with_subtotals()
    print("✅ Automation completed successfully!")
