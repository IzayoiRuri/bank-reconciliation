"""Tests for src/normalizer.py - Summary text normalization"""
import pytest
import pandas as pd
import os
import sys

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from normalizer import (
    normalize_summary,
    extract_amounts,
    extract_counterparty,
    clean_dataframe,
)


# ── Tests: normalize_summary ──────────────────────────────────────────

def test_normalize_removes_voucher_prefix():
    """凭证前缀 '收123' 应被去除，保留 '付款'"""
    result = normalize_summary("收123付款")
    assert "付款" in result
    assert "收123" not in result


def test_normalize_removes_internal_number():
    """内部编号 _2_ 应被去除"""
    result = normalize_summary("摘要_2_尾缀")
    assert "摘要" in result
    assert "尾缀" in result
    assert "_2_" not in result


def test_normalize_removes_embedded_date():
    """嵌入日期 2025.09.04 应被去除"""
    result = normalize_summary("摘要2025.09.04")
    assert "摘要" in result
    assert "2025" not in result


def test_normalize_removes_bank_code():
    """银行代号 '招行0101' 应被去除"""
    result = normalize_summary("支付招行0101手续费")
    assert "招行0101" not in result


def test_normalize_removes_separators():
    """分隔符 - 应被替换为空格"""
    result = normalize_summary("跨行-异地")
    assert "-" not in result
    assert "异地" in result


def test_normalize_real_case():
    """真实案例：银行和日记账标准化后应有高相似度"""
    bank_summary = "网银支付-跨行-异地手续费"
    ledger_summary = "支付招行0101网银跨行-异地手续费_2_2025.09.04"

    bank_norm = normalize_summary(bank_summary)
    ledger_norm = normalize_summary(ledger_summary)

    # Both should normalize to similar core keywords
    # The key is that absurdly long garbage is stripped from both
    assert len(bank_norm) > 0, f"Bank normalization produced empty string: {bank_summary!r}"
    assert len(ledger_norm) > 0, f"Ledger normalization produced empty string: {ledger_summary!r}"

    # Check that common meaningful words remain
    # After normalization, both should contain similar concepts
    # At minimum, neither should contain raw dates or internal IDs
    assert "2025.09.04" not in ledger_norm
    assert "_2_" not in ledger_norm
    assert "招行0101" not in ledger_norm

    # Check similarity: split into words and compute overlap
    bank_words = set(bank_norm.split())
    ledger_words = set(ledger_norm.split())
    overlap = bank_words & ledger_words
    assert len(overlap) > 0, (
        f"No word overlap:\n  bank: {bank_norm!r}\n  ledger: {ledger_norm!r}"
    )


def test_normalize_idempotent():
    """标准化函数应是幂等的"""
    cases = [
        "收123付款",
        "摘要_2_尾缀",
        "摘要2025.09.04",
        "支付招行0101手续费",
        "跨行-异地",
        "网银支付-跨行-异地手续费",
        "支付招行0101网银跨行-异地手续费_2_2025.09.04",
        "货款",
        "手续费",
        "",
    ]
    for text in cases:
        first = normalize_summary(text)
        second = normalize_summary(first)
        assert first == second, (
            f"Not idempotent for {text!r}:\n  first:  {first!r}\n  second: {second!r}"
        )


def test_normalize_lowercase():
    """标准化结果应为小写"""
    # Even though Chinese doesn't have case, if there's any ASCII it should be lowercase
    result = normalize_summary("ABC支付DEF")
    assert result == result.lower()


def test_normalize_strips_whitespace():
    """结果不应有首尾空格"""
    result = normalize_summary("  支付  货款  ")
    assert result == result.strip()
    # No double spaces
    assert "  " not in result


# ── Tests: extract_amounts ────────────────────────────────────────────

def test_extract_amounts_simple():
    """提取简单金额"""
    result = extract_amounts("支付货款50000.00元")
    assert 50000.0 in result


def test_extract_amounts_multiple():
    """提取多个金额"""
    result = extract_amounts("支付货款50000.00元，手续费15.50元")
    assert 50000.0 in result
    assert 15.50 in result


def test_extract_amounts_no_amount():
    """无金额时返回空列表"""
    result = extract_amounts("支付货款")
    assert result == []


def test_extract_amounts_with_comma():
    """带千分位的金额"""
    result = extract_amounts("支付1,234,567.89元")
    assert 1234567.89 in result


# ── Tests: extract_counterparty ────────────────────────────────────────

def test_extract_counterparty_company():
    """提取对方公司名称"""
    result = extract_counterparty("支付河南芯动力科技有限公司货款")
    assert result == "河南芯动力科技有限公司"


def test_extract_counterparty_factory():
    """提取厂名"""
    result = extract_counterparty("收款XX机械厂货款")
    assert result == "xx机械厂"


def test_extract_counterparty_shop():
    """提取店名"""
    result = extract_counterparty("支付XX建材店费用")
    assert result == "xx建材店"


def test_extract_counterparty_none():
    """无法识别时返回 None"""
    result = extract_counterparty("手续费")
    assert result is None


# ── Tests: clean_dataframe ────────────────────────────────────────────

def test_clean_dataframe_bank():
    """对银行 DataFrame 应用标准化"""
    df = pd.DataFrame({
        "summary": [
            "网银支付-跨行-异地手续费",
            "货款",
            "手续费",
        ],
        "amount": [202.41, 340000.0, 15.0],
    })
    result = clean_dataframe(df, source="bank")

    assert "summary" in result.columns  # Original preserved
    assert "normalized_summary" in result.columns  # New column added
    assert len(result) == 3
    assert result["normalized_summary"].iloc[0] == normalize_summary("网银支付-跨行-异地手续费")
    assert result["normalized_summary"].iloc[1] == normalize_summary("货款")
    assert result["summary"].iloc[0] == "网银支付-跨行-异地手续费"  # Original unchanged


def test_clean_dataframe_ledger():
    """对日记账 DataFrame 应用标准化"""
    df = pd.DataFrame({
        "summary": [
            "支付招行0101网银跨行-异地手续费_2_2025.09.04",
            "收河南芯动力科技有限公司货款",
        ],
        "amount": [202.41, 340000.0],
    })
    result = clean_dataframe(df, source="ledger")

    assert "normalized_summary" in result.columns
    assert len(result) == 2
    assert result["summary"].iloc[0] == "支付招行0101网银跨行-异地手续费_2_2025.09.04"


def test_clean_dataframe_empty():
    """空 DataFrame 不应崩溃"""
    df = pd.DataFrame({"summary": []})
    result = clean_dataframe(df, source="bank")
    assert "normalized_summary" in result.columns
    assert len(result) == 0


def test_clean_dataframe_missing_summary():
    """缺少 summary 列时应妥善处理"""
    df = pd.DataFrame({"other_col": [1, 2, 3]})
    result = clean_dataframe(df, source="bank")
    # Should not crash; should add a normalized_summary column
    assert "normalized_summary" in result.columns
