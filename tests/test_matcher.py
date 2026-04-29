"""Tests for src/matcher.py — 4-stage transaction matching engine"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import pytest
import pandas as pd
from datetime import datetime, timedelta
from dataclasses import dataclass

from matcher import (
    MatchRecord,
    exact_match,
    fuzzy_match,
    split_match,
    detect_duplicates,
)


# ── Config helper ─────────────────────────────────────────────────────

@dataclass
class MatchConfig:
    """Minimal config for testing; mirrors config.py parameters."""
    EXACT_MATCH_DAYS: int = 3
    FUZZY_MATCH_DAYS: int = 7
    AMOUNT_TOLERANCE: float = 0.01
    FUZZY_SCORE_THRESHOLD: int = 70
    SPLIT_MATCH_DAYS: int = 7
    DUPLICATE_CHECK_DAYS: int = 3


# ── Test DataFrames helpers ────────────────────────────────────────────

def _make_df(records, source='bank'):
    """records: list of (date_str, amount, summary)."""
    rows = []
    for d, amt, summary in records:
        rows.append({
            'date': pd.Timestamp(d),
            'normalized_amount': float(amt),
            'normalized_summary': str(summary),
            'source': source,
        })
    return pd.DataFrame(rows)


def make_bank_df(records):
    return _make_df(records, 'bank')


def make_ledger_df(records):
    return _make_df(records, 'ledger')


# ── Fixtures ──────────────────────────────────────────────────────────

@pytest.fixture
def config():
    return MatchConfig()


@pytest.fixture
def simple_bank():
    return make_bank_df([
        ("2026-01-05", 1000.00, "货款 北京科技公司"),
        ("2026-01-10", -200.00, "网银手续费"),
        ("2026-01-15", 5000.00, "货款 上海电子"),
        ("2026-01-20", -50.00, "账户管理费"),
    ])


@pytest.fixture
def simple_ledger():
    return make_ledger_df([
        ("2026-01-05", 1000.00, "收北京科技公司货款"),
        ("2026-01-11", -200.00, "支付网银异地手续费"),
        ("2026-01-15", 5000.00, "收上海电子科技有限公司货款"),
        ("2026-01-21", -50.00, "账户管理费支出"),
    ])


# ═══════════════════════════════════════════════════════════════════════
# Tests: MatchRecord
# ═══════════════════════════════════════════════════════════════════════

class TestMatchRecord:
    def test_creation_exact(self):
        rec = MatchRecord(
            bank_idx=0, ledger_idx=1, match_type='exact',
            amount_diff=0.0, date_diff=0, score=100.0,
            bank_amount=1000.0, ledger_amount=1000.0,
        )
        assert rec.bank_idx == 0
        assert rec.ledger_idx == 1
        assert rec.match_type == 'exact'
        assert rec.amount_diff == 0.0
        assert rec.date_diff == 0
        assert rec.score == 100.0

    def test_creation_split(self):
        rec = MatchRecord(
            bank_idx=0, ledger_idx=[1, 2], match_type='split',
            amount_diff=0.005, date_diff=2, score=100.0,
            bank_amount=1000.0, ledger_amount=999.995,
        )
        assert rec.ledger_idx == [1, 2]
        assert rec.match_type == 'split'

    def test_to_dict(self):
        rec = MatchRecord(
            bank_idx=3, ledger_idx=5, match_type='fuzzy',
            amount_diff=0.003, date_diff=1, score=85.0,
            bank_amount=-200.0, ledger_amount=-200.003,
        )
        d = rec.to_dict()
        assert d['bank_idx'] == 3
        assert d['ledger_idx'] == 5
        assert d['match_type'] == 'fuzzy'
        assert d['score'] == 85.0
        assert d['amount_diff'] == 0.003


# ═══════════════════════════════════════════════════════════════════════
# Tests: exact_match
# ═══════════════════════════════════════════════════════════════════════

class TestExactMatch:
    def test_basic(self, config, simple_bank, simple_ledger):
        """Same date, same amount → should match exactly."""
        matched, un_bank, un_ledger = exact_match(simple_bank, simple_ledger, config)
        # All 4 should match
        assert len(matched) == 4
        for m in matched:
            assert m.match_type == 'exact'

    def test_within_tolerance(self, config):
        """Amount diff 0.005 and date diff 2 days → should match."""
        bank = make_bank_df([("2026-01-05", 1000.00, "货款")])
        ledger = make_ledger_df([("2026-01-07", 1000.005, "收货款")])
        matched, un_bank, un_ledger = exact_match(bank, ledger, config)
        assert len(matched) == 1
        m = matched[0]
        assert m.bank_idx == 0
        assert m.ledger_idx == 0
        assert m.amount_diff == pytest.approx(0.005, abs=0.001)
        assert m.date_diff == 2

    def test_amount_outside_tolerance(self, config):
        """Amount diff 0.02 > 0.01 tolerance → should NOT match."""
        bank = make_bank_df([("2026-01-05", 1000.00, "货款")])
        ledger = make_ledger_df([("2026-01-05", 1000.02, "收货款")])
        matched, un_bank, un_ledger = exact_match(bank, ledger, config)
        assert len(matched) == 0
        assert len(un_bank) == 1
        assert len(un_ledger) == 1

    def test_date_outside_window(self, config):
        """Date diff 5 days > 3 day window → should NOT match."""
        bank = make_bank_df([("2026-01-01", 1000.00, "货款")])
        ledger = make_ledger_df([("2026-01-06", 1000.00, "收货款")])
        matched, un_bank, un_ledger = exact_match(bank, ledger, config)
        assert len(matched) == 0

    def test_one_to_one(self, config):
        """Two similar transactions each match different counterparts."""
        bank = make_bank_df([
            ("2026-01-05", 1000.00, "货款A"),
            ("2026-01-10", 2000.00, "货款B"),
        ])
        ledger = make_ledger_df([
            ("2026-01-06", 1000.00, "收货款A"),
            ("2026-01-11", 2000.00, "收货款B"),
        ])
        matched, un_bank, un_ledger = exact_match(bank, ledger, config)
        assert len(matched) == 2
        idx_pairs = sorted((m.bank_idx, m.ledger_idx) for m in matched)
        assert idx_pairs == [(0, 0), (1, 1)]

    def test_greedy_closest_date(self, config):
        """Two ledger entries for same bank entry; pick closest date."""
        bank = make_bank_df([("2026-01-05", 1000.00, "货款")])
        ledger = make_ledger_df([
            ("2026-01-07", 1000.00, "收货款远"),   # date diff 2
            ("2026-01-06", 1000.00, "收货款近"),   # date diff 1 → should be chosen
        ])
        matched, un_bank, un_ledger = exact_match(bank, ledger, config)
        assert len(matched) == 1
        assert matched[0].ledger_idx == 1  # closer date

    def test_greedy_same_date_closest_amount(self, config):
        """Same date diff, pick closest amount."""
        bank = make_bank_df([("2026-01-05", 1000.00, "货款")])
        ledger = make_ledger_df([
            ("2026-01-06", 1000.008, "收货款A"),   # diff 0.008
            ("2026-01-06", 1000.002, "收货款B"),   # diff 0.002 → closer amount
        ])
        matched, un_bank, un_ledger = exact_match(bank, ledger, config)
        assert len(matched) == 1
        assert matched[0].ledger_idx == 1  # closer amount


# ═══════════════════════════════════════════════════════════════════════
# Tests: fuzzy_match
# ═══════════════════════════════════════════════════════════════════════

class TestFuzzyMatch:
    def test_by_summary_similarity(self, config):
        """Same amount/date, different but similar summaries → fuzzy match."""
        bank = make_bank_df([("2026-01-05", 1000.00, "货款 北京科技有限公司")])
        ledger = make_ledger_df([("2026-01-06", 1000.00, "收北京科技公司货款")])
        matched, un_bank, un_ledger = fuzzy_match(bank, ledger, config)
        assert len(matched) == 1
        assert matched[0].match_type == 'fuzzy'
        assert matched[0].score >= config.FUZZY_SCORE_THRESHOLD

    def test_score_below_threshold(self, config):
        """Completely different summaries below threshold → no match."""
        bank = make_bank_df([("2026-01-05", 1000.00, "货款")])
        ledger = make_ledger_df([("2026-01-06", 1000.00, "退税")])
        matched, un_bank, un_ledger = fuzzy_match(bank, ledger, config)
        # "货款" vs "退税" — very different, score should be < 70
        assert len(matched) == 0

    def test_amount_outside_tolerance_no_fuzzy(self, config):
        """Amount diff too large → not even a candidate for fuzzy."""
        bank = make_bank_df([("2026-01-05", 1000.00, "货款北京科技公司")])
        ledger = make_ledger_df([("2026-01-06", 1000.02, "收北京科技公司货款")])
        matched, un_bank, un_ledger = fuzzy_match(bank, ledger, config)
        assert len(matched) == 0

    def test_date_outside_window_no_fuzzy(self, config):
        """Date diff > FUZZY_MATCH_DAYS → not a candidate."""
        bank = make_bank_df([("2026-01-01", 1000.00, "货款北京科技公司")])
        ledger = make_ledger_df([("2026-01-15", 1000.00, "收北京科技公司货款")])
        matched, un_bank, un_ledger = fuzzy_match(bank, ledger, config)
        assert len(matched) == 0


# ═══════════════════════════════════════════════════════════════════════
# Tests: split_match
# ═══════════════════════════════════════════════════════════════════════

class TestSplitMatch:
    def test_two_entries(self, config):
        """Bank 1000 = ledger 600 + 400."""
        bank = make_bank_df([("2026-01-05", 1000.00, "多笔货款合并")])
        ledger = make_ledger_df([
            ("2026-01-06", 600.00, "货款A"),
            ("2026-01-06", 400.00, "货款B"),
        ])
        matched, un_bank, un_ledger = split_match(bank, ledger, config)
        assert len(matched) == 1
        m = matched[0]
        assert m.match_type == 'split'
        assert isinstance(m.ledger_idx, list)
        assert set(m.ledger_idx) == {0, 1}

    def test_two_entries_small_deviation(self, config):
        """Bank 1000 vs ledger 600.005 + 400.003 = 1000.008, diff 0.008 ≤ tolerance*2=0.02."""
        bank = make_bank_df([("2026-01-05", 1000.00, "多笔")])
        ledger = make_ledger_df([
            ("2026-01-06", 600.005, "A"),
            ("2026-01-06", 400.003, "B"),
        ])
        matched, un_bank, un_ledger = split_match(bank, ledger, config)
        assert len(matched) == 1

    def test_exceeds_max_combinations(self, config):
        """More than 5 ledger entries needed → no split match."""
        bank = make_bank_df([("2026-01-05", 600.00, "多笔")])
        ledger = make_ledger_df([
            ("2026-01-06", 100.00, "A"),
            ("2026-01-06", 100.00, "B"),
            ("2026-01-06", 100.00, "C"),
            ("2026-01-06", 100.00, "D"),
            ("2026-01-06", 100.00, "E"),
            ("2026-01-06", 100.00, "F"),  # 6th entry → too many
        ])
        matched, un_bank, un_ledger = split_match(bank, ledger, config)
        # 6 entries each 100 doesn't sum to 600 with ≤5 entries? Wait:
        # 100*6 can be done with 5 entries (100+100+100+100+100=500) or 6 entries...
        # But 6 entries creates >5 combinations. With 5 entries we get 500, not 600.
        # So no valid 2-5 combination sums to 600 ± (0.01*n)
        assert len(matched) == 0

    def test_date_outside_window(self, config):
        """Ledger entries outside SPLIT_MATCH_DAYS are not candidates."""
        bank = make_bank_df([("2026-01-05", 1000.00, "多笔")])
        ledger = make_ledger_df([
            ("2026-01-15", 600.00, "A"),   # 10 days away
            ("2026-01-06", 400.00, "B"),
        ])
        matched, un_bank, un_ledger = split_match(bank, ledger, config)
        # Entry A at 10 days is outside window, only B at 400 remains
        assert len(matched) == 0

    def test_no_candidate_ledgers(self, config):
        """No ledger entries at all."""
        bank = make_bank_df([("2026-01-05", 1000.00, "多笔")])
        ledger = make_ledger_df([])
        matched, un_bank, un_ledger = split_match(bank, ledger, config)
        assert len(matched) == 0
        assert len(un_bank) == 1


# ═══════════════════════════════════════════════════════════════════════
# Tests: detect_duplicates
# ═══════════════════════════════════════════════════════════════════════

class TestDetectDuplicates:
    def test_same_date_same_amount(self, config):
        """Two entries same date, same amount → marked as duplicate."""
        df = make_bank_df([
            ("2026-01-05", 1000.00, "货款"),
            ("2026-01-05", 1000.00, "货款"),  # duplicate
            ("2026-01-10", 500.00, "手续费"),
        ])
        result = detect_duplicates(df, 'bank', config)
        assert 'is_duplicate' in result.columns
        assert 'duplicate_group_id' in result.columns
        # First two should be duplicates
        assert result.iloc[0]['is_duplicate'] == True
        assert result.iloc[1]['is_duplicate'] == True
        assert result.iloc[2]['is_duplicate'] == False
        # Same group_id
        assert result.iloc[0]['duplicate_group_id'] == result.iloc[1]['duplicate_group_id']

    def test_different_date_within_window(self, config):
        """Same amount, dates within 3 days → duplicate."""
        df = make_bank_df([
            ("2026-01-05", 1000.00, "货款"),
            ("2026-01-07", 1000.00, "货款"),  # 2 days apart, within window
        ])
        result = detect_duplicates(df, 'bank', config)
        assert result.iloc[0]['is_duplicate'] == True
        assert result.iloc[1]['is_duplicate'] == True

    def test_different_date_outside_window(self, config):
        """Same amount, dates beyond 3 days → not duplicate."""
        df = make_bank_df([
            ("2026-01-01", 1000.00, "货款"),
            ("2026-01-10", 1000.00, "货款"),  # 9 days apart
        ])
        result = detect_duplicates(df, 'bank', config)
        assert result.iloc[0]['is_duplicate'] == False
        assert result.iloc[1]['is_duplicate'] == False

    def test_different_amount_not_duplicate(self, config):
        """Same date, different amount → not duplicate."""
        df = make_bank_df([
            ("2026-01-05", 1000.00, "货款"),
            ("2026-01-05", 500.00, "费用"),
        ])
        result = detect_duplicates(df, 'bank', config)
        assert not result['is_duplicate'].any()

    def test_amount_within_tolerance(self, config):
        """Amounts differ by less than tolerance → still duplicate."""
        df = make_bank_df([
            ("2026-01-05", 1000.00, "货款"),
            ("2026-01-05", 1000.005, "货款"),
        ])
        result = detect_duplicates(df, 'bank', config)
        assert result.iloc[0]['is_duplicate'] == True
        assert result.iloc[1]['is_duplicate'] == True


# ═══════════════════════════════════════════════════════════════════════
# Tests: greedy prevents double match
# ═══════════════════════════════════════════════════════════════════════

class TestGreedyPreventsDoubleMatch:
    def test_matched_not_available_for_later(self, config):
        """Exact-matched entries should not be available for fuzzy."""
        bank = make_bank_df([
            ("2026-01-05", 1000.00, "货款北京科技"),
            ("2026-01-10", 2000.00, "货款上海电子"),
        ])
        ledger = make_ledger_df([
            ("2026-01-05", 1000.00, "收货款北京科技"),  # exact match bank[0]
            ("2026-01-14", 2000.00, "收上海电子货款"),  # 4 days away → outside exact, inside fuzzy
        ])

        # Phase 1: exact — only bank[0]→ledger[0] (bank[1]→ledger[1] date diff=4 > 3)
        matched_1, un_bank, un_ledger = exact_match(bank, ledger, config)
        assert len(matched_1) == 1
        assert matched_1[0].bank_idx == 0
        assert matched_1[0].ledger_idx == 0

        # Phase 2: fuzzy on unmatched
        matched_2, un_bank2, un_ledger2 = fuzzy_match(un_bank, un_ledger, config)
        assert len(matched_2) == 1  # bank[1]→ledger[1] fuzzy

        # ledger[0] should not be matched twice
        all_matched_ledger = [m.ledger_idx for m in matched_1] + [m.ledger_idx for m in matched_2]
        assert len(all_matched_ledger) == len(set(all_matched_ledger))

    def test_each_transaction_matched_once(self, config):
        """Verify each bank/ledger transaction is matched at most once across phases."""
        bank = make_bank_df([
            ("2026-01-05", 1000.00, "A"),
            ("2026-01-06", 500.00, "B"),
        ])
        ledger = make_ledger_df([
            ("2026-01-05", 1000.00, "A"),
            ("2026-01-06", 500.00, "B"),
        ])

        matched, un_bank, un_ledger = exact_match(bank, ledger, config)
        assert len(matched) == 2

        bank_indices = set(m.bank_idx for m in matched)
        ledger_indices = set(m.ledger_idx for m in matched)
        assert len(bank_indices) == 2  # both bank rows matched
        assert len(ledger_indices) == 2  # both ledger rows matched


# ═══════════════════════════════════════════════════════════════════════
# Tests: edge cases
# ═══════════════════════════════════════════════════════════════════════

class TestEdgeCases:
    def test_empty_bank(self, config):
        bank = make_bank_df([])
        ledger = make_ledger_df([("2026-01-05", 1000.00, "货款")])
        matched, un_bank, un_ledger = exact_match(bank, ledger, config)
        assert len(matched) == 0
        assert len(un_bank) == 0
        assert len(un_ledger) == 1

    def test_empty_ledger(self, config):
        bank = make_bank_df([("2026-01-05", 1000.00, "货款")])
        ledger = make_ledger_df([])
        matched, un_bank, un_ledger = exact_match(bank, ledger, config)
        assert len(matched) == 0
        assert len(un_bank) == 1
        assert len(un_ledger) == 0

    def test_empty_both(self, config):
        bank = make_bank_df([])
        ledger = make_ledger_df([])
        matched, un_bank, un_ledger = exact_match(bank, ledger, config)
        assert len(matched) == 0

