"""银行对账工具 — 文件解析器

解析银行流水(.xlsx)和公司日记账(.xls)为统一格式的 DataFrame。
"""

import pandas as pd
import openpyxl
import xlrd
import re
from datetime import datetime

from config import BANK_COLUMN_ALIASES, LEDGER_COLUMN_ALIASES


# ═══════════════════════════════════════════════════════════════════════
# Column auto-detection
# ═══════════════════════════════════════════════════════════════════════

def auto_detect_columns(df, aliases_dict):
    """根据列名别名映射自动匹配 DataFrame 的实际列名到标准列名。

    Args:
        df: pandas DataFrame
        aliases_dict: {标准列名: [别名列表]}

    Returns:
        dict: {标准列名: 实际列名}，未匹配到的值为 None
    """
    result = {}
    actual_cols = [str(c).strip() for c in df.columns]

    for std_name, aliases in aliases_dict.items():
        matched = None
        for alias in aliases:
            if alias in actual_cols:
                matched = alias
                break
        result[std_name] = matched

    return result


# ═══════════════════════════════════════════════════════════════════════
# Bank statement parser (.xlsx)
# ═══════════════════════════════════════════════════════════════════════

def _find_bank_header_row(ws):
    """Find the header row in a bank statement worksheet.

    Scans rows looking for one that contains common bank statement column names
    like '交易日', '借方金额', etc. Returns the 1-indexed row number.
    """
    # Keywords that a bank statement header row should contain
    header_keywords = ['交易日', '交易时间', '借方金额', '贷方金额']

    for row_idx in range(1, ws.max_row + 1):
        row_vals = [str(cell.value) if cell.value is not None else '' for cell in ws[row_idx]]
        row_text = ' '.join(row_vals)

        matches = sum(1 for kw in header_keywords if kw in row_text)
        if matches >= 3:  # At least 3 of the keywords must be present
            return row_idx

    return 1  # Fallback: assume first row is header


def parse_bank_statement(filepath):
    """解析银行流水文件 (.xlsx)。

    自动识别列名并标准化输出：
    date, debit, credit, balance, summary, counterparty, counterparty_acct

    Args:
        filepath: .xlsx 文件路径

    Returns:
        pandas DataFrame
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    # Find the header row
    header_row = _find_bank_header_row(ws)

    # Read header values (all columns)
    header_vals = []
    for cell in ws[header_row]:
        val = cell.value
        header_vals.append(str(val).strip() if val is not None else '')

    # Read all data rows after header
    data_rows = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        row_vals = []
        for cell in ws[row_idx]:
            row_vals.append(cell.value)
        data_rows.append(row_vals)

    # Build DataFrame with all columns
    df = pd.DataFrame(data_rows, columns=header_vals)

    # Auto-detect column mappings
    col_map = auto_detect_columns(df, BANK_COLUMN_ALIASES)

    # Build standardized DataFrame
    result = pd.DataFrame()

    # Date
    date_col = col_map.get('date')
    if date_col and date_col in df.columns:
        result['date'] = pd.to_datetime(df[date_col], errors='coerce')
    else:
        result['date'] = pd.NaT

    # Debit (借方金额)
    debit_col = col_map.get('debit')
    if debit_col and debit_col in df.columns:
        result['debit'] = pd.to_numeric(df[debit_col], errors='coerce').fillna(0)
    else:
        result['debit'] = 0.0

    # Credit (贷方金额)
    credit_col = col_map.get('credit')
    if credit_col and credit_col in df.columns:
        result['credit'] = pd.to_numeric(df[credit_col], errors='coerce').fillna(0)
    else:
        result['credit'] = 0.0

    # Balance (余额)
    balance_col = col_map.get('balance')
    if balance_col and balance_col in df.columns:
        result['balance'] = pd.to_numeric(df[balance_col], errors='coerce').fillna(0)
    else:
        result['balance'] = 0.0

    # Summary (摘要)
    summary_col = col_map.get('summary')
    if summary_col and summary_col in df.columns:
        result['summary'] = df[summary_col].fillna('').astype(str)
    else:
        result['summary'] = ''

    # Counterparty (收(付)方名称)
    cp_col = col_map.get('counterparty')
    if cp_col and cp_col in df.columns:
        result['counterparty'] = df[cp_col].fillna('').astype(str)
    else:
        result['counterparty'] = ''

    # Counterparty account (收(付)方账号)
    cp_acct_col = col_map.get('counterparty_acct')
    if cp_acct_col and cp_acct_col in df.columns:
        result['counterparty_acct'] = df[cp_acct_col].fillna('').astype(str)
    else:
        result['counterparty_acct'] = ''

    # Drop rows where date is NaT (metadata/empty rows)
    result = result.dropna(subset=['date']).reset_index(drop=True)

    # Fill any remaining NaN amounts with 0
    for col in ['debit', 'credit', 'balance']:
        result[col] = result[col].fillna(0)

    return result


# ═══════════════════════════════════════════════════════════════════════
# Ledger parser (.xls)
# ═══════════════════════════════════════════════════════════════════════

def _is_summary_row(row_values):
    """Check if a ledger row is a summary row (本日合计/本月合计/本年累计/当前合计/当前累计/上年结转).

    Args:
        row_values: list of cell values for a row

    Returns:
        bool: True if row should be filtered out
    """
    # Convert all values to strings for pattern matching
    row_str = '|'.join([str(v) if v is not None else '' for v in row_values])

    # Check for summary keywords
    summary_keywords = ['合计', '累计', '上年结转']
    for kw in summary_keywords:
        if kw in row_str:
            return True

    return False


def parse_ledger(filepath):
    """解析公司日记账文件 (.xls)。

    过滤日合计/月合计/累计行，合并年/月/日为 date 列。
    标准化输出：
    date, summary, voucher_no, counterparty_subject, debit, credit, direction, balance

    Args:
        filepath: .xls 文件路径

    Returns:
        pandas DataFrame
    """
    wb = xlrd.open_workbook(filepath)
    ws = wb.sheet_by_index(0)

    # Read headers from first row
    headers = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]

    # Read all data rows
    data_rows = []
    for r in range(1, ws.nrows):
        row_vals = [ws.cell_value(r, c) for c in range(ws.ncols)]

        # Skip summary rows
        if _is_summary_row(row_vals):
            continue

        data_rows.append(row_vals)

    # Build DataFrame
    df = pd.DataFrame(data_rows, columns=headers)

    # Auto-detect column mappings
    col_map = auto_detect_columns(df, LEDGER_COLUMN_ALIASES)

    # Build standardized DataFrame
    result = pd.DataFrame()

    # Combine year, month, day into date
    year_col = col_map.get('year')
    month_col = col_map.get('month')
    day_col = col_map.get('day')

    year_vals = pd.to_numeric(df[year_col], errors='coerce').fillna(1).astype(int) if year_col else 2026
    month_vals = pd.to_numeric(df[month_col], errors='coerce').fillna(1).astype(int) if month_col else 1
    day_vals = pd.to_numeric(df[day_col], errors='coerce').fillna(1).astype(int) if day_col else 1

    # Build datetime strings
    def build_date(y, m, d):
        try:
            return pd.Timestamp(year=int(y), month=int(m), day=int(d))
        except (ValueError, TypeError):
            return pd.NaT

    dates = []
    for i in range(len(df)):
        dates.append(build_date(year_vals.iloc[i] if hasattr(year_vals, 'iloc') else year_vals[i] if isinstance(year_vals, (list, pd.Series)) else year_vals,
                                 month_vals.iloc[i] if hasattr(month_vals, 'iloc') else month_vals[i] if isinstance(month_vals, (list, pd.Series)) else month_vals,
                                 day_vals.iloc[i] if hasattr(day_vals, 'iloc') else day_vals[i] if isinstance(day_vals, (list, pd.Series)) else day_vals))

    result['date'] = dates

    # Summary
    summary_col = col_map.get('summary')
    if summary_col and summary_col in df.columns:
        result['summary'] = df[summary_col].fillna('').astype(str)
    else:
        result['summary'] = ''

    # Voucher number
    voucher_col = col_map.get('voucher_no')
    if voucher_col and voucher_col in df.columns:
        result['voucher_no'] = df[voucher_col].fillna('').astype(str)
    else:
        result['voucher_no'] = ''

    # Counterparty subject
    cp_subj_col = col_map.get('counterparty_subject')
    if cp_subj_col and cp_subj_col in df.columns:
        result['counterparty_subject'] = df[cp_subj_col].fillna('').astype(str)
    else:
        result['counterparty_subject'] = ''

    # Debit
    debit_col = col_map.get('debit')
    if debit_col and debit_col in df.columns:
        result['debit'] = pd.to_numeric(df[debit_col], errors='coerce').fillna(0)
    else:
        result['debit'] = 0.0

    # Credit
    credit_col = col_map.get('credit')
    if credit_col and credit_col in df.columns:
        result['credit'] = pd.to_numeric(df[credit_col], errors='coerce').fillna(0)
    else:
        result['credit'] = 0.0

    # Direction
    direction_col = col_map.get('direction')
    if direction_col and direction_col in df.columns:
        result['direction'] = df[direction_col].fillna('').astype(str)
    else:
        result['direction'] = ''

    # Balance
    balance_col = col_map.get('balance')
    if balance_col and balance_col in df.columns:
        result['balance'] = pd.to_numeric(df[balance_col], errors='coerce').fillna(0)
    else:
        result['balance'] = 0.0

    # Drop rows where date is NaT
    result = result.dropna(subset=['date']).reset_index(drop=True)

    # Fill NaN amounts
    for col in ['debit', 'credit', 'balance']:
        result[col] = result[col].fillna(0)

    return result


# ═══════════════════════════════════════════════════════════════════════
# Transaction normalization
# ═══════════════════════════════════════════════════════════════════════

def normalize_transaction(df, source):
    """统一标准化交易记录。

    - 添加 source 列
    - 添加 normalized_amount（统一符号：正数=收入，负数=支出）
      - 银行：贷方 - 借方
      - 日记账：借方 - 贷方
    - 添加 month 列（YYYY-MM）
    - 清理空值和异常值

    Args:
        df: parse_bank_statement 或 parse_ledger 输出的 DataFrame
        source: 'bank' 或 'ledger'

    Returns:
        标准化后的 DataFrame（添加了 source, normalized_amount, month 列）
    """
    df = df.copy()

    # Add source
    df['source'] = source

    # Compute normalized amount
    if source == 'bank':
        # 银行：贷方=收入(正), 借方=支出(负)
        df['normalized_amount'] = df['credit'] - df['debit']
    elif source == 'ledger':
        # 日记账：借方=收入(正), 贷方=支出(负) — 与银行相反
        df['normalized_amount'] = df['debit'] - df['credit']
    else:
        raise ValueError(f"Unknown source: {source}. Must be 'bank' or 'ledger'.")

    # Add month column (YYYY-MM)
    if 'date' in df.columns:
        df['month'] = df['date'].apply(
            lambda d: d.strftime('%Y-%m') if pd.notna(d) else ''
        )
    else:
        df['month'] = ''

    # Clean abnormal values (extremely large amounts that might be errors)
    # Keep this simple - just ensure amounts are within reasonable bounds
    # Could be expanded later

    return df


# ═══════════════════════════════════════════════════════════════════════
# Main entry point
# ═══════════════════════════════════════════════════════════════════════

def load_and_parse(bank_path, ledger_path):
    """加载并解析银行流水和公司日记账。

    Args:
        bank_path: 银行流水 .xlsx 文件路径
        ledger_path: 公司日记账 .xls 文件路径

    Returns:
        (bank_df, ledger_df): 标准化后的两个 DataFrame
    """
    bank_df = parse_bank_statement(bank_path)
    ledger_df = parse_ledger(ledger_path)

    bank_df = normalize_transaction(bank_df, 'bank')
    ledger_df = normalize_transaction(ledger_df, 'ledger')

    return bank_df, ledger_df
