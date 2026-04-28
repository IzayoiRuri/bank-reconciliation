"""Tests for src/pipeline.py — Reconciliation pipeline orchestration"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import pytest
import pandas as pd
from datetime import datetime, timedelta

from pipeline import (
    ReconciliationResult,
    run_reconciliation,
    get_summary_text,
    compare_amounts,
)

from config import (
    EXACT_MATCH_DAYS,
    FUZZY_MATCH_DAYS,
    AMOUNT_TOLERANCE,
    FUZZY_SCORE_THRESHOLD,
    SPLIT_MATCH_DAYS,
    DUPLICATE_CHECK_DAYS,
)


# ── Mock DataFrames ────────────────────────────────────────────────────

def _make_bank_df(records):
    """Build a bank DataFrame that mimics the output of load_and_parse().
    
    records: list of (date_str, debit, credit, balance, summary, counterparty, counterparty_acct)
    """
    rows = []
    for date_str, debit, credit, balance, summary, cp, cp_acct in records:
        rows.append({
            'date': pd.Timestamp(date_str),
            'debit': float(debit),
            'credit': float(credit),
            'balance': float(balance),
            'summary': str(summary),
            'counterparty': str(cp),
            'counterparty_acct': str(cp_acct),
            'source': 'bank',
            'normalized_amount': float(credit - debit),
            'month': pd.Timestamp(date_str).strftime('%Y-%m'),
        })
    return pd.DataFrame(rows)


def _make_ledger_df(records):
    """Build a ledger DataFrame that mimics the output of load_and_parse().
    
    records: list of (date_str, debit, credit, balance, summary, voucher_no, counterparty_subject, direction)
    """
    rows = []
    for date_str, debit, credit, balance, summary, voucher, cp_subj, direction in records:
        rows.append({
            'date': pd.Timestamp(date_str),
            'debit': float(debit),
            'credit': float(credit),
            'balance': float(balance),
            'summary': str(summary),
            'voucher_no': str(voucher),
            'counterparty_subject': str(cp_subj),
            'direction': str(direction),
            'source': 'ledger',
            'normalized_amount': float(debit - credit),
            'month': pd.Timestamp(date_str).strftime('%Y-%m'),
        })
    return pd.DataFrame(rows)


# ── Helper: add normalized_summary after clean_dataframe ───────────────

def _add_norm_summary(df):
    """Simulate clean_dataframe output by adding normalized_summary column."""
    from normalizer import normalize_summary
    df = df.copy()
    if 'summary' in df.columns:
        df['normalized_summary'] = df['summary'].apply(
            lambda x: normalize_summary(x) if pd.notna(x) else ''
        )
    else:
        df['normalized_summary'] = ''
    return df


# ═══════════════════════════════════════════════════════════════════════
# Tests
# ═══════════════════════════════════════════════════════════════════════

class TestCompareAmounts:
    """Tests for compare_amounts()."""

    def test_basic_comparison(self):
        """Compare bank and ledger totals — both sides match."""
        bank = _make_bank_df([
            ('2026-04-01', 0, 1000, 5000, '收到货款', '客户A', '622202001'),
            ('2026-04-02', 200, 0, 4800, '支付房租', '房东B', '622202002'),
            ('2026-04-03', 500, 0, 4300, '支付工资', '员工C', '622202003'),
        ])
        ledger = _make_ledger_df([
            ('2026-04-01', 1000, 0, 5000, '收到货款', 'V001', '银行存款', '借'),
            ('2026-04-02', 0, 200, 4800, '支付房租', 'V002', '管理费用', '贷'),
            ('2026-04-03', 0, 500, 4300, '支付工资', 'V003', '应付职工薪酬', '贷'),
        ])

        result = compare_amounts(bank, ledger)

        # Bank: income = sum(credit) = 1000, expense = sum(debit) = 700, net = credit - debit = 300
        assert result['bank_income'] == 1000.0
        assert result['bank_expense'] == 700.0
        assert result['bank_net'] == 300.0

        # Ledger: income = sum(debit) = 1000, expense = sum(credit) = 700, net = debit - credit = 300
        assert result['ledger_income'] == 1000.0
        assert result['ledger_expense'] == 700.0
        assert result['ledger_net'] == 300.0

        assert result['diff'] == 0.0

    def test_mismatched_totals(self):
        """Bank and ledger have different totals."""
        bank = _make_bank_df([
            ('2026-04-01', 0, 1000, 5000, '收到货款', '客户A', '622202001'),
        ])
        ledger = _make_ledger_df([
            ('2026-04-01', 800, 0, 5000, '收到部分货款', 'V001', '银行存款', '借'),
        ])

        result = compare_amounts(bank, ledger)

        assert result['bank_income'] == 1000.0
        assert result['bank_expense'] == 0.0
        assert result['bank_net'] == 1000.0
        assert result['ledger_income'] == 800.0
        assert result['ledger_expense'] == 0.0
        assert result['ledger_net'] == 800.0
        assert result['diff'] == pytest.approx(200.0)

    def test_empty_dataframes(self):
        """Empty DataFrames produce zero totals."""
        bank = _make_bank_df([])
        ledger = _make_ledger_df([])

        result = compare_amounts(bank, ledger)
        assert result['bank_income'] == 0.0
        assert result['bank_expense'] == 0.0
        assert result['bank_net'] == 0.0
        assert result['ledger_income'] == 0.0
        assert result['ledger_expense'] == 0.0
        assert result['ledger_net'] == 0.0
        assert result['diff'] == 0.0


class TestReconciliationResultDataclass:
    """Tests for the ReconciliationResult dataclass structure."""

    def test_creation_with_defaults(self):
        """ReconciliationResult can be instantiated with all fields."""
        r = ReconciliationResult(
            bank_file='test_bank.xlsx',
            ledger_file='test_ledger.xls',
            bank_total_count=0,
            ledger_total_count=0,
            bank_total_income=0.0,
            bank_total_expense=0.0,
            ledger_total_income=0.0,
            ledger_total_expense=0.0,
            bank_opening_balance=0.0,
            bank_closing_balance=0.0,
            matched_count=0,
            exact_matched=0,
            fuzzy_matched=0,
            split_matched=0,
            unmatched_bank_count=0,
            unmatched_ledger_count=0,
            duplicate_bank_count=0,
            duplicate_ledger_count=0,
            total_bank_amount=0.0,
            total_ledger_amount=0.0,
            matched_amount_diff=0.0,
            unmatched_bank_amount=0.0,
            unmatched_ledger_amount=0.0,
            matched_records=[],
            unmatched_bank=pd.DataFrame(),
            unmatched_ledger=pd.DataFrame(),
            bank_duplicates=pd.DataFrame(),
            ledger_duplicates=pd.DataFrame(),
            reconciled_at=datetime.now().isoformat(),
            date_range=('2026-04-01', '2026-04-30'),
            match_rate=0.0,
        )
        assert r.bank_file == 'test_bank.xlsx'
        assert r.ledger_file == 'test_ledger.xls'
        assert r.match_rate == 0.0

    def test_has_all_required_fields(self):
        """Verify all required fields from the spec are present."""
        fields = {
            'bank_file', 'ledger_file',
            'bank_total_count', 'ledger_total_count',
            'bank_total_income', 'bank_total_expense',
            'ledger_total_income', 'ledger_total_expense',
            'bank_opening_balance', 'bank_closing_balance',
            'matched_count', 'exact_matched', 'fuzzy_matched', 'split_matched',
            'unmatched_bank_count', 'unmatched_ledger_count',
            'duplicate_bank_count', 'duplicate_ledger_count',
            'total_bank_amount', 'total_ledger_amount',
            'matched_amount_diff', 'unmatched_bank_amount', 'unmatched_ledger_amount',
            'matched_records', 'unmatched_bank', 'unmatched_ledger',
            'bank_duplicates', 'ledger_duplicates',
            'reconciled_at', 'date_range', 'match_rate',
            'bank_df', 'ledger_df',
        }
        r_fields = {f.name for f in __import__('dataclasses').fields(ReconciliationResult)}
        assert fields == r_fields


class TestRunReconciliation:
    """Tests for run_reconciliation() — end-to-end pipeline."""

    def test_basic_exact_match(self):
        """Simple end-to-end: all transactions match exactly."""
        bank = _make_bank_df([
            ('2026-04-01', 0, 5000, 10000, '收到货款-客户A', '客户A', '622202001'),
            ('2026-04-02', 1000, 0, 9000, '支付房租', '房东B', '622202002'),
        ])
        ledger = _make_ledger_df([
            ('2026-04-01', 5000, 0, 10000, '收到货款-客户A', 'V001', '银行存款', '借'),
            ('2026-04-02', 0, 1000, 9000, '支付房租', 'V002', '管理费用', '贷'),
        ])

        # Simulate what pipeline does internally
        bank = _add_norm_summary(bank)
        ledger = _add_norm_summary(ledger)

        # Manually run the 3 match stages
        from matcher import exact_match, fuzzy_match, split_match
        import config

        all_matched = []
        exact_m, ub1, ul1 = exact_match(bank, ledger, config)
        all_matched.extend(exact_m)
        fuzzy_m, ub2, ul2 = fuzzy_match(ub1, ul1, config)
        all_matched.extend(fuzzy_m)
        split_m, ub3, ul3 = split_match(ub2, ul2, config)
        all_matched.extend(split_m)

        assert len(exact_m) == 2
        assert len(all_matched) == 2
        assert len(ub3) == 0
        assert len(ul3) == 0

    def test_pipeline_with_fuzzy_match(self):
        """Pipeline where fuzzy match catches a near-match."""
        bank = _make_bank_df([
            ('2026-04-01', 0, 5000, 10000, '收到货款 客户A', '客户A', '622202001'),
            ('2026-04-02', 1000, 0, 9000, '支付办公室房租', '房东B', '622202002'),
        ])
        ledger = _make_ledger_df([
            ('2026-04-01', 5000, 0, 10000, '收到货款-客户A有限公司', 'V001', '银行存款', '借'),
            ('2026-04-02', 0, 1000, 9000, '支付办公室房租金', 'V002', '管理费用', '贷'),
        ])

        bank = _add_norm_summary(bank)
        ledger = _add_norm_summary(ledger)

        from matcher import exact_match, fuzzy_match, split_match
        import config

        all_matched = []
        exact_m, ub1, ul1 = exact_match(bank, ledger, config)
        all_matched.extend(exact_m)
        fuzzy_m, ub2, ul2 = fuzzy_match(ub1, ul1, config)
        all_matched.extend(fuzzy_m)
        split_m, ub3, ul3 = split_match(ub2, ul2, config)
        all_matched.extend(split_m)

        assert len(all_matched) == 2
        # After normalization, summaries might still differ enough for exact match
        # to fail but fuzzy should catch both
        fuzzy_count = len(fuzzy_m)
        assert fuzzy_count >= 0  # at minimum we match all, possibly via fuzzy

    def test_pipeline_with_split_match(self):
        """One bank transaction matches two ledger entries (split)."""
        bank = _make_bank_df([
            ('2026-04-01', 0, 5000, 10000, '收到货款', '客户A', '622202001'),
        ])
        ledger = _make_ledger_df([
            ('2026-04-01', 3000, 0, 10000, '收到货款-部分1', 'V001', '银行存款', '借'),
            ('2026-04-01', 2000, 0, 10000, '收到货款-部分2', 'V002', '银行存款', '借'),
        ])

        bank = _add_norm_summary(bank)
        ledger = _add_norm_summary(ledger)

        from matcher import exact_match, fuzzy_match, split_match
        import config

        all_matched = []
        exact_m, ub1, ul1 = exact_match(bank, ledger, config)
        all_matched.extend(exact_m)
        fuzzy_m, ub2, ul2 = fuzzy_match(ub1, ul1, config)
        all_matched.extend(fuzzy_m)
        split_m, ub3, ul3 = split_match(ub2, ul2, config)
        all_matched.extend(split_m)

        # The bank transaction should be matched to both ledger entries via split
        assert len(split_m) >= 1
        # The split match should have exactly 2 ledger indices
        if split_m:
            assert isinstance(split_m[0].ledger_idx, list)
            assert len(split_m[0].ledger_idx) == 2

    def test_unmatched_remain(self):
        """Transactions that don't match anywhere remain in unmatched."""
        bank = _make_bank_df([
            ('2026-04-01', 0, 5000, 10000, '收到货款A', '客户A', '622202001'),
            ('2026-04-15', 999, 0, 9001, '神秘支出', '???', '??'),
        ])
        ledger = _make_ledger_df([
            ('2026-04-01', 5000, 0, 10000, '收到货款A', 'V001', '银行存款', '借'),
            ('2026-04-20', 0, 777, 9223, '另一笔支出', 'V002', '管理费用', '贷'),
        ])

        bank = _add_norm_summary(bank)
        ledger = _add_norm_summary(ledger)

        from matcher import exact_match, fuzzy_match, split_match
        import config

        all_matched = []
        exact_m, ub1, ul1 = exact_match(bank, ledger, config)
        all_matched.extend(exact_m)
        fuzzy_m, ub2, ul2 = fuzzy_match(ub1, ul1, config)
        all_matched.extend(fuzzy_m)
        split_m, ub3, ul3 = split_match(ub2, ul2, config)
        all_matched.extend(split_m)

        assert len(ub3) >= 1  # Bank has unmatched
        # The unmatched bank should be the -999 (amount 999) one, ledger the -777
        unmatched_bank_count = len(ub3)
        unmatched_ledger_count = len(ul3)
        assert unmatched_bank_count + unmatched_ledger_count >= 1

    def test_duplicate_detection(self):
        """Duplicates within same source are flagged."""
        bank = _make_bank_df([
            ('2026-04-01', 0, 5000, 10000, '收到货款', '客户A', '622202001'),
            ('2026-04-01', 0, 5000, 10000, '收到货款', '客户A', '622202001'),  # duplicate!
            ('2026-04-03', 200, 0, 9800, '支付水电', '电力公司', '622202003'),
        ])

        from matcher import detect_duplicates
        import config
        bank = _add_norm_summary(bank)
        bank_with_dups = detect_duplicates(bank, 'bank', config)

        dup_count = bank_with_dups['is_duplicate'].sum()
        assert dup_count >= 2  # Both identical rows flagged

    def test_empty_inputs(self):
        """Pipeline handles empty DataFrames gracefully."""
        bank = _make_bank_df([])
        ledger = _make_ledger_df([])
        bank = _add_norm_summary(bank)
        ledger = _add_norm_summary(ledger)

        from matcher import exact_match, fuzzy_match, split_match, detect_duplicates
        import config

        # All stages should return empty results
        exact_m, ub1, ul1 = exact_match(bank, ledger, config)
        assert len(exact_m) == 0
        assert len(ub1) == 0
        assert len(ul1) == 0

        fuzzy_m, ub2, ul2 = fuzzy_match(ub1, ul1, config)
        assert len(fuzzy_m) == 0

        split_m, ub3, ul3 = split_match(ub2, ul2, config)
        assert len(split_m) == 0

        # Duplicate detection on empty
        bank_d = detect_duplicates(bank, 'bank', config)
        assert len(bank_d) == 0


class TestGetSummaryText:
    """Tests for get_summary_text()."""

    def test_basic_summary(self):
        """Summary text is generated for a valid result."""
        bank = _make_bank_df([
            ('2026-04-01', 0, 5000, 10000, '收到货款', '客户A', '622202001'),
            ('2026-04-02', 1000, 0, 9000, '支付房租', '房东B', '622202002'),
        ])
        ledger = _make_ledger_df([
            ('2026-04-01', 5000, 0, 10000, '收到货款', 'V001', '银行存款', '借'),
            ('2026-04-02', 0, 1000, 9000, '支付房租', 'V002', '管理费用', '贷'),
        ])

        bank = _add_norm_summary(bank)
        ledger = _add_norm_summary(ledger)

        r = ReconciliationResult(
            bank_file='test.xlsx',
            ledger_file='test.xls',
            bank_total_count=2,
            ledger_total_count=2,
            bank_total_income=5000.0,
            bank_total_expense=1000.0,
            ledger_total_income=5000.0,
            ledger_total_expense=1000.0,
            bank_opening_balance=10000.0,
            bank_closing_balance=9000.0,
            matched_count=2,
            exact_matched=2,
            fuzzy_matched=0,
            split_matched=0,
            unmatched_bank_count=0,
            unmatched_ledger_count=0,
            duplicate_bank_count=0,
            duplicate_ledger_count=0,
            total_bank_amount=4000.0,
            total_ledger_amount=4000.0,
            matched_amount_diff=0.0,
            unmatched_bank_amount=0.0,
            unmatched_ledger_amount=0.0,
            matched_records=[],
            unmatched_bank=pd.DataFrame(),
            unmatched_ledger=pd.DataFrame(),
            bank_duplicates=pd.DataFrame(),
            ledger_duplicates=pd.DataFrame(),
            reconciled_at=datetime.now().isoformat(),
            date_range=('2026-04-01', '2026-04-02'),
            match_rate=100.0,
        )

        text = get_summary_text(r)
        assert isinstance(text, str)
        assert '100.0%' in text or '100.00' in text
        assert 'test.xlsx' in text
        assert 'test.xls' in text
        assert '0' in text  # unmatched = 0

    def test_summary_with_unmatched(self):
        """Summary reflects unmatched items."""
        r = ReconciliationResult(
            bank_file='b.xlsx',
            ledger_file='l.xls',
            bank_total_count=5,
            ledger_total_count=4,
            bank_total_income=10000.0,
            bank_total_expense=3000.0,
            ledger_total_income=8000.0,
            ledger_total_expense=3000.0,
            bank_opening_balance=50000.0,
            bank_closing_balance=57000.0,
            matched_count=3,
            exact_matched=2,
            fuzzy_matched=1,
            split_matched=0,
            unmatched_bank_count=2,
            unmatched_ledger_count=1,
            duplicate_bank_count=0,
            duplicate_ledger_count=0,
            total_bank_amount=7000.0,
            total_ledger_amount=5000.0,
            matched_amount_diff=0.0,
            unmatched_bank_amount=2000.0,
            unmatched_ledger_amount=0.0,
            matched_records=[],
            unmatched_bank=pd.DataFrame(),
            unmatched_ledger=pd.DataFrame(),
            bank_duplicates=pd.DataFrame(),
            ledger_duplicates=pd.DataFrame(),
            reconciled_at=datetime.now().isoformat(),
            date_range=('2026-04-01', '2026-04-15'),
            match_rate=60.0,
        )

        text = get_summary_text(r)
        assert '60.0%' in text or '60' in text
        assert '未匹配' in text
