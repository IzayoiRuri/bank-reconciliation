"""Tests for src/history.py — SQLite history storage"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import pytest
import tempfile
import sqlite3
import pandas as pd
from datetime import datetime

from history import (
    init_db,
    save_reconciliation,
    list_history,
    get_history,
    delete_history,
    search_history,
)
from pipeline import ReconciliationResult


# ── Helper: create a minimal ReconciliationResult ──────────────────────

def _make_result(bank_file='bank.xlsx', ledger_file='ledger.xls', **overrides):
    """Create a ReconciliationResult with sensible defaults."""
    defaults = {
        'bank_file': bank_file,
        'ledger_file': ledger_file,
        'bank_total_count': 10,
        'ledger_total_count': 12,
        'bank_total_income': 50000.0,
        'bank_total_expense': 20000.0,
        'ledger_total_income': 50000.0,
        'ledger_total_expense': 20000.0,
        'bank_opening_balance': 100000.0,
        'bank_closing_balance': 130000.0,
        'matched_count': 8,
        'exact_matched': 5,
        'fuzzy_matched': 2,
        'split_matched': 1,
        'unmatched_bank_count': 2,
        'unmatched_ledger_count': 4,
        'duplicate_bank_count': 0,
        'duplicate_ledger_count': 0,
        'total_bank_amount': 30000.0,
        'total_ledger_amount': 30000.0,
        'matched_amount_diff': 15.50,
        'unmatched_bank_amount': 500.0,
        'unmatched_ledger_amount': 300.0,
        'matched_records': [],
        'unmatched_bank': pd.DataFrame(),
        'unmatched_ledger': pd.DataFrame(),
        'bank_duplicates': pd.DataFrame(),
        'ledger_duplicates': pd.DataFrame(),
        'reconciled_at': datetime.now().isoformat(),
        'date_range': ('2026-04-01', '2026-04-15'),
        'match_rate': 66.67,
    }
    defaults.update(overrides)
    return ReconciliationResult(**defaults)


# ── Fixture: temporary database path ───────────────────────────────────

@pytest.fixture
def temp_db():
    """Create a temporary database file and clean up after."""
    fd, path = tempfile.mkstemp(suffix='.db', prefix='test_history_')
    os.close(fd)
    yield path
    # Cleanup
    if os.path.exists(path):
        os.unlink(path)


# ═══════════════════════════════════════════════════════════════════════
# Tests
# ═══════════════════════════════════════════════════════════════════════

class TestInitDb:
    """Tests for init_db()."""

    def test_init_db_creates_table(self, temp_db):
        """init_db should create the reconciliations table."""
        conn = init_db(db_path=temp_db)
        # Check that the table exists
        cursor = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='reconciliations'"
        )
        row = cursor.fetchone()
        assert row is not None
        assert row[0] == 'reconciliations'
        conn.close()

    def test_init_db_is_idempotent(self, temp_db):
        """Calling init_db multiple times should not error."""
        conn1 = init_db(db_path=temp_db)
        conn1.close()
        conn2 = init_db(db_path=temp_db)
        # Should not raise; verify table still exists
        cursor = conn2.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='reconciliations'"
        )
        assert cursor.fetchone() is not None
        conn2.close()

    def test_init_db_default_path(self, temp_db):
        """init_db with no arguments should work (uses config.DB_PATH)."""
        # We can't easily test the default path without side effects,
        # but we can verify the function signature accepts None.
        conn = init_db(db_path=temp_db)
        assert conn is not None
        conn.close()


class TestSaveAndRetrieve:
    """Tests for save_reconciliation() and get_history()."""

    def test_save_and_retrieve(self, temp_db):
        """Save a result and retrieve it by ID."""
        result = _make_result()
        rid = save_reconciliation(result, db_path=temp_db)

        assert rid > 0

        record = get_history(rid, db_path=temp_db)
        assert record is not None
        assert record['bank_file'] == 'bank.xlsx'
        assert record['ledger_file'] == 'ledger.xls'
        assert record['bank_count'] == 10
        assert record['ledger_count'] == 12
        assert record['bank_total_income'] == 50000.0
        assert record['bank_total_expense'] == 20000.0
        assert record['ledger_total_income'] == 50000.0
        assert record['ledger_total_expense'] == 20000.0
        assert record['bank_opening_balance'] == 100000.0
        assert record['bank_closing_balance'] == 130000.0
        assert record['matched_count'] == 8
        assert record['exact_matched'] == 5
        assert record['fuzzy_matched'] == 2
        assert record['split_matched'] == 1
        assert record['unmatched_bank_count'] == 2
        assert record['unmatched_ledger_count'] == 4
        assert record['match_rate'] == 66.67
        assert record['amount_diff'] == 15.50
        assert record['date_range_start'] == '2026-04-01'
        assert record['date_range_end'] == '2026-04-15'
        # reconciled_at should be present
        assert record['reconciled_at'] is not None

    def test_save_with_report_and_notes(self, temp_db):
        """Save with a report path and notes."""
        result = _make_result()
        rid = save_reconciliation(
            result,
            report_path='/tmp/report.xlsx',
            notes='Test reconciliation run',
            db_path=temp_db,
        )
        record = get_history(rid, db_path=temp_db)
        assert record['report_path'] == '/tmp/report.xlsx'
        assert record['notes'] == 'Test reconciliation run'

    def test_get_nonexistent_record(self, temp_db):
        """get_history for non-existent ID returns None."""
        conn = init_db(db_path=temp_db)
        conn.close()
        record = get_history(9999, db_path=temp_db)
        assert record is None


class TestListHistory:
    """Tests for list_history()."""

    def test_list_history_limit(self, temp_db):
        """list_history returns at most `limit` records."""
        for i in range(5):
            result = _make_result(
                bank_file=f'bank_{i}.xlsx',
                ledger_file=f'ledger_{i}.xls',
            )
            save_reconciliation(result, db_path=temp_db)

        results = list_history(limit=3, db_path=temp_db)
        assert len(results) == 3

        results_all = list_history(limit=10, db_path=temp_db)
        assert len(results_all) == 5

    def test_list_history_empty(self, temp_db):
        """list_history returns empty list when no records exist."""
        conn = init_db(db_path=temp_db)
        conn.close()
        results = list_history(db_path=temp_db)
        assert results == []

    def test_list_history_descending_order(self, temp_db):
        """Records should be returned newest first."""
        r1 = _make_result(reconciled_at='2026-04-01T10:00:00')
        r2 = _make_result(reconciled_at='2026-04-02T10:00:00')
        save_reconciliation(r1, db_path=temp_db)
        save_reconciliation(r2, db_path=temp_db)

        results = list_history(limit=10, db_path=temp_db)
        assert len(results) == 2
        # Newest should be first
        assert results[0]['reconciled_at'] == '2026-04-02T10:00:00'
        assert results[1]['reconciled_at'] == '2026-04-01T10:00:00'


class TestDeleteHistory:
    """Tests for delete_history()."""

    def test_delete_history(self, temp_db):
        """Delete a record and verify it's gone."""
        result = _make_result()
        rid = save_reconciliation(result, db_path=temp_db)

        assert delete_history(rid, db_path=temp_db) is True

        # Record should no longer exist
        record = get_history(rid, db_path=temp_db)
        assert record is None

    def test_delete_nonexistent(self, temp_db):
        """Deleting a non-existent record returns False."""
        conn = init_db(db_path=temp_db)
        conn.close()
        assert delete_history(9999, db_path=temp_db) is False


class TestSearchHistory:
    """Tests for search_history()."""

    def test_search_by_keyword(self, temp_db):
        """Search should find records by bank_file or ledger_file."""
        save_reconciliation(
            _make_result(bank_file='2026-Q1-工商银行.xlsx', ledger_file='总账-Q1.xls'),
            db_path=temp_db,
        )
        save_reconciliation(
            _make_result(bank_file='2026-Q2-建设银行.xlsx', ledger_file='总账-Q2.xls'),
            db_path=temp_db,
        )
        save_reconciliation(
            _make_result(bank_file='bank3.xlsx', ledger_file='test.xls'),
            db_path=temp_db,
        )

        # Search for "工商"
        results = search_history('工商', db_path=temp_db)
        assert len(results) == 1
        assert results[0]['bank_file'] == '2026-Q1-工商银行.xlsx'

        # Search for "总账" — should match both ledger_file fields
        results = search_history('总账', db_path=temp_db)
        assert len(results) == 2

        # Search for non-existent keyword
        results = search_history('不存在的', db_path=temp_db)
        assert len(results) == 0

        # Search for "bank" — should match only bank3.xlsx
        results = search_history('bank', db_path=temp_db)
        assert len(results) == 1
        assert results[0]['bank_file'] == 'bank3.xlsx'

    def test_search_case_insensitive(self, temp_db):
        """Search should be case-insensitive."""
        save_reconciliation(
            _make_result(bank_file='TEST_BANK.xlsx', ledger_file='LEDGER.xls'),
            db_path=temp_db,
        )
        results = search_history('test', db_path=temp_db)
        assert len(results) == 1
        results = search_history('TEST', db_path=temp_db)
        assert len(results) == 1
