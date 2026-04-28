"""Tests for config module."""

from src.config import (
    BANK_COLUMN_ALIASES,
    LEDGER_COLUMN_ALIASES,
    EXACT_MATCH_DAYS,
    FUZZY_MATCH_DAYS,
    AMOUNT_TOLERANCE,
    FUZZY_SCORE_THRESHOLD,
    SPLIT_MATCH_DAYS,
    DUPLICATE_CHECK_DAYS,
    REPORT_WORKSHEETS,
    DB_PATH,
    SUMMARY_STOP_WORDS,
    SUMMARY_REMOVE_PATTERNS,
)


class TestConfig:
    """Verify all config attributes exist and have correct types."""

    def test_bank_column_aliases(self):
        assert isinstance(BANK_COLUMN_ALIASES, dict)
        assert "date" in BANK_COLUMN_ALIASES
        assert "debit" in BANK_COLUMN_ALIASES
        assert "credit" in BANK_COLUMN_ALIASES
        assert "balance" in BANK_COLUMN_ALIASES
        assert "summary" in BANK_COLUMN_ALIASES
        assert "counterparty" in BANK_COLUMN_ALIASES
        assert "counterparty_acct" in BANK_COLUMN_ALIASES

    def test_ledger_column_aliases(self):
        assert isinstance(LEDGER_COLUMN_ALIASES, dict)
        assert "date" in LEDGER_COLUMN_ALIASES
        assert "summary" in LEDGER_COLUMN_ALIASES
        assert "debit" in LEDGER_COLUMN_ALIASES
        assert "credit" in LEDGER_COLUMN_ALIASES

    def test_matching_params(self):
        assert EXACT_MATCH_DAYS == 3
        assert FUZZY_MATCH_DAYS == 7
        assert AMOUNT_TOLERANCE == 0.01
        assert FUZZY_SCORE_THRESHOLD == 70
        assert SPLIT_MATCH_DAYS == 7
        assert DUPLICATE_CHECK_DAYS == 3

    def test_report_config(self):
        assert isinstance(REPORT_WORKSHEETS, list)
        assert len(REPORT_WORKSHEETS) == 5

    def test_db_path(self):
        assert DB_PATH == "reconciliation_history.db"

    def test_summary_config(self):
        assert isinstance(SUMMARY_STOP_WORDS, list)
        assert len(SUMMARY_STOP_WORDS) > 0
        assert isinstance(SUMMARY_REMOVE_PATTERNS, list)
        assert len(SUMMARY_REMOVE_PATTERNS) > 0
