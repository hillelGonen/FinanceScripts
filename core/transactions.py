import glob
import os

import pandas as pd


def clean_dates(date_series):
    date_series = date_series.astype(str).str.strip()
    date_series = date_series.str.replace('.', '-', regex=False)
    date_series = date_series.str.replace('/', '-', regex=False)
    return pd.to_datetime(date_series, dayfirst=True, errors='coerce')


def deduplicate_transactions(df):
    dupes = df.duplicated(subset=['Date', 'Merchant', 'Amount'], keep='first')
    return df[~dupes].copy(), int(dupes.sum())


def _detect_header_row(df_preview):
    for i, row in df_preview.iterrows():
        row_str = row.astype(str).str.cat(sep=' ')
        if "תאריך" in row_str and ("רכישה" in row_str or "עסקה" in row_str):
            return i
    return 0


def _parse_excel_dataframe(df, source_name):
    df.columns = df.columns.str.strip()

    merchant_col = next((c for c in ['שם בית עסק', 'שם בית העסק'] if c in df.columns), None)
    amount_col = 'סכום חיוב' if 'סכום חיוב' in df.columns else None
    date_col = next((c for c in ['תאריך רכישה', 'תאריך עסקה', 'תאריך'] if c in df.columns), None)

    if not merchant_col or not amount_col:
        return None

    temp_df = pd.DataFrame()
    temp_df['Date'] = clean_dates(df[date_col]) if date_col else pd.NaT
    temp_df['Merchant'] = df[merchant_col]
    temp_df['Amount'] = df[amount_col]
    temp_df['Source_File'] = source_name

    summary_phrases = ['TOTAL FOR DATE', 'סה"כ לחיוב', 'סה"כ', 'סך הכל']
    mask = ~temp_df['Merchant'].astype(str).str.contains('|'.join(summary_phrases), regex=True, na=False)
    temp_df = temp_df[mask]

    temp_df = temp_df.dropna(subset=['Amount'])
    temp_df['Amount'] = temp_df['Amount'].astype(str).str.replace(r'[₪,]', '', regex=True).str.strip()
    temp_df['Amount'] = pd.to_numeric(temp_df['Amount'], errors='coerce')
    temp_df = temp_df.dropna(subset=['Amount'])

    return temp_df


def read_transactions_from_file(filepath):
    try:
        try:
            preview = pd.read_excel(filepath, header=None, nrows=30)
        except Exception:
            return None

        header_row = _detect_header_row(preview)
        df = pd.read_excel(filepath, skiprows=header_row)
        return _parse_excel_dataframe(df, os.path.basename(filepath))

    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return None


def find_excel_files(pattern="**/*.xlsx", folder_to_ignore=None):
    all_files = glob.glob(pattern, recursive=True)
    valid_files = []
    for f in all_files:
        if os.path.basename(f).startswith('~$'):
            continue
        if folder_to_ignore and folder_to_ignore in f:
            continue
        valid_files.append(f)
    return valid_files
