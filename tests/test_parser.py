"""Tests for src/parser.py - Bank statement and ledger parsing"""
import pytest
import pandas as pd
import openpyxl
import xlwt
import os
import tempfile
from datetime import date

# Add src to path
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from parser import (
    auto_detect_columns,
    parse_bank_statement,
    parse_ledger,
    normalize_transaction,
    load_and_parse,
)


# ── Fixtures ──────────────────────────────────────────────────────────

@pytest.fixture
def sample_bank_aliases():
    """Alias dict matching BANK_COLUMN_ALIASES from config."""
    return {
        "date": ["交易日", "日期", "交易日期", "记账日期"],
        "debit": ["借方金额", "支出金额", "付款金额", "借方"],
        "credit": ["贷方金额", "收入金额", "收款金额", "贷方"],
        "balance": ["余额", "账户余额", "本次余额"],
        "summary": ["摘要", "交易摘要", "用途", "备注", "交易说明"],
        "counterparty": ["收(付)方名称", "对方户名", "交易对方", "对方名称", "对手名称"],
        "counterparty_acct": ["收(付)方账号", "对方账号", "对方账户"],
    }


@pytest.fixture
def sample_ledger_aliases():
    """Alias dict matching LEDGER_COLUMN_ALIASES from config."""
    return {
        "year": ["年"],
        "month": ["月"],
        "day": ["日"],
        "date": ["日期", "记账日期"],
        "summary": ["摘要", "备注", "说明"],
        "voucher_no": ["结算号", "凭证号", "结算方式", "票号"],
        "counterparty_subject": ["对方科目", "对方单位", "科目"],
        "debit": ["借方金额", "收入金额", "借方"],
        "credit": ["贷方金额", "支出金额", "贷方"],
        "direction": ["方向"],
        "balance": ["余额金额", "余额"],
    }


@pytest.fixture
def bank_xlsx_path():
    """Create a minimal .xlsx bank statement for testing."""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Simulate a bank statement with a metadata header area, then a data table
    # Metadata rows (rows 1-3)
    ws.append(["银行对账单"])
    ws.append(["账号: 1234567890"])
    ws.append([])  # empty row

    # Header row (row 4)
    headers = ["交易日", "交易时间", "交易类型", "借方金额", "贷方金额", "余额", "摘要", "收(付)方名称", "收(付)方账号"]
    ws.append(headers)

    # Data rows
    data = [
        ["2026-01-02", "14:37:29", "贷记", None, 340000, 10096724.64, "货款", "河南芯动力科技有限公司", "1702020609200865910"],
        ["2026-01-04", "06:37:17", "费用", 202.41, None, 10096522.23, "网银支付-跨行-异地手续费", "网上电子汇划收入", "911052750071101810"],
        ["2026-01-04", "06:37:17", "费用", 138.42, None, 10096383.81, "网银支付-行内-异地手续费", None, None],
        ["2026-01-04", "16:24:01", "贷记", None, 36250, 10132608.81, "货款", "沈阳嘉睿电子科技有限公司", "2202020609200865911"],
        ["2026-01-05", "10:00:00", "费用", 500.00, None, 10132108.81, "手续费", "某银行", "123456"],
    ]
    for row in data:
        ws.append(row)

    wb.save(path)
    yield path
    os.unlink(path)


@pytest.fixture
def ledger_xls_path():
    """Create a minimal .xls ledger for testing."""
    with tempfile.NamedTemporaryFile(suffix='.xls', delete=False) as f:
        path = f.name

    wb = xlwt.Workbook()
    ws = wb.add_sheet("Sheet1")

    # Headers
    headers = ["年", "月", "日", "凭证号数", "摘要", "结算号", "对方科目", "借方金额", "贷方金额", "方向", "余额金额"]
    for col, h in enumerate(headers):
        ws.write(0, col, h)

    # Data rows
    data = [
        [2026, 1, 2, "回-0001", "收河南芯动力科技有限公司货款", "网银转账", "112201|人民币", 340000.0, 0.0, "借", 10096724.64],
        [2026, 1, 4, "付-0001", "支付网银服务费", "网银转账", "660301", 0.0, 202.41, "借", 10096522.23],
        [2026, 1, 4, "", "本日合计", "", "", 340000.0, 202.41, "借", 10096522.23],  # 日合计行 - should be filtered
        [2026, 1, "", "", "本月合计", "", "", 340000.0, 202.41, "借", 10096522.23],  # 月合计行 - should be filtered
        [2026, 1, 5, "回-0002", "收货款", "网银转账", "112201|人民币", 160500.0, 0.0, "借", 10257022.23],
        [2026, 1, "", "", "本年累计", "", "", 500500.0, 202.41, "借", 10257022.23],  # 累计行 - should be filtered
    ]
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            ws.write(r + 1, c, val)

    wb.save(path)
    yield path
    os.unlink(path)


# ── Tests: auto_detect_columns ────────────────────────────────────────

def test_auto_detect_columns_exact_match(sample_bank_aliases):
    """Should map standard column names to actual column names when they match exactly."""
    df = pd.DataFrame(columns=["交易日", "借方金额", "贷方金额", "余额", "摘要", "收(付)方名称", "收(付)方账号"])
    result = auto_detect_columns(df, sample_bank_aliases)
    assert result["date"] == "交易日"
    assert result["debit"] == "借方金额"
    assert result["credit"] == "贷方金额"
    assert result["balance"] == "余额"
    assert result["summary"] == "摘要"
    assert result["counterparty"] == "收(付)方名称"
    assert result["counterparty_acct"] == "收(付)方账号"


def test_auto_detect_columns_alias_match(sample_bank_aliases):
    """Should match via aliases when column names differ from primary alias."""
    df = pd.DataFrame(columns=["日期", "支出金额", "收入金额", "账户余额", "备注", "对方户名", "对方账号"])
    result = auto_detect_columns(df, sample_bank_aliases)
    assert result["date"] == "日期"
    assert result["debit"] == "支出金额"
    assert result["credit"] == "收入金额"
    assert result["balance"] == "账户余额"
    assert result["summary"] == "备注"
    assert result["counterparty"] == "对方户名"
    assert result["counterparty_acct"] == "对方账号"


def test_auto_detect_columns_missing_column(sample_bank_aliases):
    """Should return None for columns that don't exist in the dataframe."""
    df = pd.DataFrame(columns=["交易日", "借方金额", "贷方金额"])
    result = auto_detect_columns(df, sample_bank_aliases)
    assert result["date"] == "交易日"
    assert result["debit"] == "借方金额"
    assert result["credit"] == "贷方金额"
    assert result["balance"] is None
    assert result["summary"] is None


def test_auto_detect_columns_ledger(sample_ledger_aliases):
    """Should match ledger-specific columns."""
    df = pd.DataFrame(columns=["年", "月", "日", "摘要", "结算号", "对方科目", "借方金额", "贷方金额", "方向", "余额金额"])
    result = auto_detect_columns(df, sample_ledger_aliases)
    assert result["year"] == "年"
    assert result["month"] == "月"
    assert result["day"] == "日"
    assert result["summary"] == "摘要"
    assert result["voucher_no"] == "结算号"
    assert result["counterparty_subject"] == "对方科目"
    assert result["debit"] == "借方金额"
    assert result["credit"] == "贷方金额"
    assert result["direction"] == "方向"
    assert result["balance"] == "余额金额"


# ── Tests: parse_bank_statement ───────────────────────────────────────

def test_parse_bank_statement_structure(bank_xlsx_path):
    """Should parse bank statement into a DataFrame with correct standardized columns."""
    df = parse_bank_statement(bank_xlsx_path)
    expected_cols = ["date", "debit", "credit", "balance", "summary", "counterparty", "counterparty_acct"]
    for col in expected_cols:
        assert col in df.columns, f"Missing column: {col}"

    assert len(df) == 5  # 5 data rows (metadata rows skipped)
    assert df["debit"].dtype in ('float64', 'int64')
    assert df["credit"].dtype in ('float64', 'int64')


def test_parse_bank_statement_values(bank_xlsx_path):
    """Should correctly parse values from bank statement."""
    df = parse_bank_statement(bank_xlsx_path)

    # First row: credit transaction
    row0 = df.iloc[0]
    assert row0["date"] == pd.Timestamp("2026-01-02") or str(row0["date"]) == "2026-01-02"
    assert row0["debit"] == 0.0  # NaN should be filled with 0
    assert row0["credit"] == 340000.0
    assert row0["counterparty"] == "河南芯动力科技有限公司"

    # Second row: debit transaction
    row1 = df.iloc[1]
    assert row1["debit"] == 202.41
    assert row1["credit"] == 0.0
    assert row1["summary"] == "网银支付-跨行-异地手续费"


def test_parse_bank_statement_nan_filled(bank_xlsx_path):
    """NaN values in debit/credit should be filled with 0."""
    df = parse_bank_statement(bank_xlsx_path)
    assert (df["debit"].isna().sum()) == 0
    assert (df["credit"].isna().sum()) == 0


# ── Tests: parse_ledger ───────────────────────────────────────────────

def test_parse_ledger_structure(ledger_xls_path):
    """Should parse ledger into a DataFrame with correct standardized columns."""
    df = parse_ledger(ledger_xls_path)
    expected_cols = ["date", "summary", "voucher_no", "counterparty_subject", "debit", "credit", "direction", "balance"]
    for col in expected_cols:
        assert col in df.columns, f"Missing column: {col}"


def test_parse_ledger_filter_summary_rows(ledger_xls_path):
    """Should filter out 本日合计, 本月合计, 本年累计 rows."""
    df = parse_ledger(ledger_xls_path)
    # 6 data rows total, but 3 are summary rows (本日合计, 本月合计, 本年累计)
    # Should be 3 regular transaction rows
    assert len(df) == 3

    # Verify no summary rows remain
    summary_mask = df["summary"].str.contains("合计|累计", na=False)
    assert summary_mask.sum() == 0, f"Summary rows not filtered: {df[summary_mask]['summary'].tolist()}"


def test_parse_ledger_date_combination(ledger_xls_path):
    """Should combine year, month, day columns into a single date column."""
    df = parse_ledger(ledger_xls_path)
    # First data row: 2026-01-02
    assert str(df.iloc[0]["date"])[:10] == "2026-01-02"
    # Second data row: 2026-01-04
    assert str(df.iloc[1]["date"])[:10] == "2026-01-04"
    # Third data row: 2026-01-05
    assert str(df.iloc[2]["date"])[:10] == "2026-01-05"


def test_parse_ledger_values(ledger_xls_path):
    """Should correctly parse values from ledger."""
    df = parse_ledger(ledger_xls_path)

    # First row: income (debit > 0)
    row0 = df.iloc[0]
    assert row0["debit"] == 340000.0
    assert row0["credit"] == 0.0
    assert row0["summary"] == "收河南芯动力科技有限公司货款"
    assert row0["voucher_no"] == "网银转账"

    # Second row: expense (credit > 0)
    row1 = df.iloc[1]
    assert row1["debit"] == 0.0
    assert row1["credit"] == 202.41


def test_parse_ledger_nan_filled(ledger_xls_path):
    """NaN values in debit/credit should be filled with 0."""
    df = parse_ledger(ledger_xls_path)
    assert df["debit"].isna().sum() == 0
    assert df["credit"].isna().sum() == 0


# ── Tests: normalize_transaction ──────────────────────────────────────

@pytest.fixture
def sample_bank_df():
    """Create a sample bank DataFrame as output by parse_bank_statement."""
    return pd.DataFrame({
        "date": pd.to_datetime(["2026-01-02", "2026-01-04", "2026-01-05"]),
        "debit": [0.0, 202.41, 500.0],
        "credit": [340000.0, 0.0, 0.0],
        "balance": [10096724.64, 10096522.23, 10096022.23],
        "summary": ["货款", "手续费", "手续费"],
        "counterparty": ["公司A", "银行B", "银行C"],
        "counterparty_acct": ["123", "456", "789"],
    })


@pytest.fixture
def sample_ledger_df():
    """Create a sample ledger DataFrame as output by parse_ledger."""
    return pd.DataFrame({
        "date": pd.to_datetime(["2026-01-02", "2026-01-04", "2026-01-05"]),
        "summary": ["收货款", "支付费用", "收货款"],
        "voucher_no": ["回-0001", "付-0001", "回-0002"],
        "counterparty_subject": ["112201", "660301", "112201"],
        "debit": [340000.0, 0.0, 160500.0],
        "credit": [0.0, 202.41, 0.0],
        "direction": ["借", "借", "借"],
        "balance": [10096724.64, 10096522.23, 10257022.23],
    })


def test_normalize_bank_adds_source(sample_bank_df):
    """Should add source='bank' column."""
    df = normalize_transaction(sample_bank_df, "bank")
    assert "source" in df.columns
    assert (df["source"] == "bank").all()


def test_normalize_ledger_adds_source(sample_ledger_df):
    """Should add source='ledger' column."""
    df = normalize_transaction(sample_ledger_df, "ledger")
    assert "source" in df.columns
    assert (df["source"] == "ledger").all()


def test_normalize_bank_amount_sign(sample_bank_df):
    """Bank: normalized_amount = credit - debit (positive=income, negative=expense)."""
    df = normalize_transaction(sample_bank_df, "bank")
    # Row 0: credit=340000, debit=0 -> income (+340000)
    assert df.iloc[0]["normalized_amount"] == pytest.approx(340000.0)
    # Row 1: credit=0, debit=202.41 -> expense (-202.41)
    assert df.iloc[1]["normalized_amount"] == pytest.approx(-202.41)
    # Row 2: credit=0, debit=500 -> expense (-500)
    assert df.iloc[2]["normalized_amount"] == pytest.approx(-500.0)


def test_normalize_ledger_amount_sign(sample_ledger_df):
    """Ledger: normalized_amount = debit - credit (positive=income, negative=expense)."""
    df = normalize_transaction(sample_ledger_df, "ledger")
    # Row 0: debit=340000, credit=0 -> income (+340000)
    assert df.iloc[0]["normalized_amount"] == pytest.approx(340000.0)
    # Row 1: debit=0, credit=202.41 -> expense (-202.41)
    assert df.iloc[1]["normalized_amount"] == pytest.approx(-202.41)
    # Row 2: debit=160500, credit=0 -> income (+160500)
    assert df.iloc[2]["normalized_amount"] == pytest.approx(160500.0)


def test_normalize_adds_month(sample_bank_df):
    """Should add month column in YYYY-MM format."""
    df = normalize_transaction(sample_bank_df, "bank")
    assert "month" in df.columns
    assert df.iloc[0]["month"] == "2026-01"
    assert df.iloc[1]["month"] == "2026-01"


# ── Tests: load_and_parse ─────────────────────────────────────────────

def test_load_and_parse(bank_xlsx_path, ledger_xls_path):
    """Integration test: load_and_parse should return both DataFrames with proper structure."""
    bank_df, ledger_df = load_and_parse(bank_xlsx_path, ledger_xls_path)

    # Bank checks
    assert "source" in bank_df.columns
    assert (bank_df["source"] == "bank").all()
    assert "normalized_amount" in bank_df.columns
    assert "month" in bank_df.columns
    assert len(bank_df) == 5

    # Ledger checks
    assert "source" in ledger_df.columns
    assert (ledger_df["source"] == "ledger").all()
    assert "normalized_amount" in ledger_df.columns
    assert "month" in ledger_df.columns
    assert len(ledger_df) == 3  # 3 summary rows filtered out
