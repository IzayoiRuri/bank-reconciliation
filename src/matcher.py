"""银行对账工具 — 4阶段交易匹配引擎

精确匹配 → 模糊匹配 → 拆分匹配 → 重复检测
"""

import itertools
from dataclasses import dataclass, field, asdict
from typing import List, Union, Tuple

import pandas as pd
from thefuzz import fuzz

from config import (
    EXACT_MATCH_DAYS,
    FUZZY_MATCH_DAYS,
    AMOUNT_TOLERANCE,
    FUZZY_SCORE_THRESHOLD,
    SPLIT_MATCH_DAYS,
    DUPLICATE_CHECK_DAYS,
)


# ═══════════════════════════════════════════════════════════════════════
# MatchRecord
# ═══════════════════════════════════════════════════════════════════════

@dataclass
class MatchRecord:
    """Represents a single match between a bank transaction and one or more
    ledger transactions.

    Attributes:
        bank_idx: Index of the bank transaction in the bank DataFrame.
        ledger_idx: Index (int for exact/fuzzy) or list of indices (for split).
        match_type: One of 'exact', 'fuzzy', 'split'.
        amount_diff: Absolute difference between bank amount and ledger amount(s).
        date_diff: Days between bank date and (earliest) ledger date.
        score: Similarity score (100 for exact, fuzzy score for fuzzy, 100 for split).
        bank_amount: The normalized_amount of the bank transaction.
        ledger_amount: The sum of ledger normalized_amount(s).
    """
    bank_idx: int
    ledger_idx: Union[int, List[int]]
    match_type: str
    amount_diff: float
    date_diff: int
    score: float
    bank_amount: float
    ledger_amount: float

    def to_dict(self) -> dict:
        """Convert to a plain dictionary for serialization."""
        d = asdict(self)
        return d


# ═══════════════════════════════════════════════════════════════════════
# Internal helpers
# ═══════════════════════════════════════════════════════════════════════

def _get_config_value(config, key, default):
    """Extract a config value from a module, dataclass, or dict."""
    if hasattr(config, key):
        return getattr(config, key)
    if isinstance(config, dict) and key in config:
        return config[key]
    return default


def _build_candidates(bank_df, ledger_df, amount_tolerance, date_window, exclude_bank=None, exclude_ledger=None):
    """Build a list of candidate (bank_idx, ledger_idx, amount_diff, date_diff, score).

    Returns a list of dicts sorted by score descending (best match first).
    For exact matching, score is based on date proximity and amount proximity.
    """
    if exclude_bank is None:
        exclude_bank = set()
    if exclude_ledger is None:
        exclude_ledger = set()

    candidates = []

    for bi in bank_df.index:
        if bi in exclude_bank:
            continue
        b_date = bank_df.loc[bi, 'date']
        b_amt = bank_df.loc[bi, 'normalized_amount']

        for li in ledger_df.index:
            if li in exclude_ledger:
                continue
            l_date = ledger_df.loc[li, 'date']
            l_amt = ledger_df.loc[li, 'normalized_amount']

            date_diff = abs((b_date - l_date).days)
            if date_diff > date_window:
                continue

            amt_diff = abs(b_amt - l_amt)
            if amt_diff > amount_tolerance:
                continue

            candidates.append({
                'bank_idx': bi,
                'ledger_idx': li,
                'amount_diff': amt_diff,
                'date_diff': date_diff,
            })

    return candidates


# ═══════════════════════════════════════════════════════════════════════
# 1. exact_match
# ═══════════════════════════════════════════════════════════════════════

def exact_match(bank_df, ledger_df, config):
    """Match bank and ledger transactions on amount and date within tolerance.

    For each bank transaction, find ledger transactions where:
    - abs(amount_diff) ≤ AMOUNT_TOLERANCE
    - abs(date_diff) ≤ EXACT_MATCH_DAYS

    Greedy matching: closest date first; same date → closest amount.

    Returns:
        (matched_records, unmatched_bank_df, unmatched_ledger_df)
    """
    days = _get_config_value(config, 'EXACT_MATCH_DAYS', EXACT_MATCH_DAYS)
    tolerance = _get_config_value(config, 'AMOUNT_TOLERANCE', AMOUNT_TOLERANCE)

    candidates = _build_candidates(bank_df, ledger_df, tolerance, days)

    if not candidates:
        return [], bank_df.copy(), ledger_df.copy()

    # Greedy matching: process bank transactions one by one
    matched_bank = set()
    matched_ledger = set()
    matched_records = []

    # Sort bank indices for deterministic processing
    for bi in sorted(set(c['bank_idx'] for c in candidates)):
        if bi in matched_bank:
            continue

        # Find all candidates for this bank
        bi_cands = [c for c in candidates if c['bank_idx'] == bi and c['ledger_idx'] not in matched_ledger]
        if not bi_cands:
            continue

        # Sort: closest date first, then closest amount
        bi_cands.sort(key=lambda c: (c['date_diff'], c['amount_diff']))
        best = bi_cands[0]

        # Double-check: could another bank also want this ledger?
        # Check if any other bank has a better claim on this ledger
        other_claims = [
            c for c in candidates
            if c['ledger_idx'] == best['ledger_idx']
            and c['bank_idx'] != bi
            and c['bank_idx'] not in matched_bank
        ]
        # If there's a competing claim with better date_diff, let them have it
        better_claim = False
        for oc in other_claims:
            if (oc['date_diff'], oc['amount_diff']) < (best['date_diff'], best['amount_diff']):
                better_claim = True
                break

        if better_claim:
            continue

        # Record the match
        b_amt = bank_df.loc[bi, 'normalized_amount']
        l_amt = ledger_df.loc[best['ledger_idx'], 'normalized_amount']

        record = MatchRecord(
            bank_idx=bi,
            ledger_idx=best['ledger_idx'],
            match_type='exact',
            amount_diff=best['amount_diff'],
            date_diff=best['date_diff'],
            score=100.0,
            bank_amount=float(b_amt),
            ledger_amount=float(l_amt),
        )
        matched_records.append(record)
        matched_bank.add(bi)
        matched_ledger.add(best['ledger_idx'])

    # Build unmatched DataFrames
    unmatched_bank = bank_df.loc[~bank_df.index.isin(matched_bank)].copy()
    unmatched_ledger = ledger_df.loc[~ledger_df.index.isin(matched_ledger)].copy()

    return matched_records, unmatched_bank, unmatched_ledger


# ═══════════════════════════════════════════════════════════════════════
# 2. fuzzy_match
# ═══════════════════════════════════════════════════════════════════════

def fuzzy_match(bank_df, ledger_df, config):
    """Match remaining transactions using fuzzy string comparison on summaries.

    Candidate filter:
    - abs(amount_diff) ≤ AMOUNT_TOLERANCE
    - abs(date_diff) ≤ FUZZY_MATCH_DAYS
    - thefuzz token_sort_ratio ≥ FUZZY_SCORE_THRESHOLD

    Greedy matching: highest fuzzy score first.

    Returns:
        (matched_records, unmatched_bank_df, unmatched_ledger_df)
    """
    days = _get_config_value(config, 'FUZZY_MATCH_DAYS', FUZZY_MATCH_DAYS)
    tolerance = _get_config_value(config, 'AMOUNT_TOLERANCE', AMOUNT_TOLERANCE)
    threshold = _get_config_value(config, 'FUZZY_SCORE_THRESHOLD', FUZZY_SCORE_THRESHOLD)

    if bank_df.empty or ledger_df.empty:
        return [], bank_df.copy(), ledger_df.copy()

    # Build fuzzy candidates
    candidates = []
    for bi in bank_df.index:
        b_date = bank_df.loc[bi, 'date']
        b_amt = bank_df.loc[bi, 'normalized_amount']
        b_summary = str(bank_df.loc[bi, 'normalized_summary'])

        for li in ledger_df.index:
            l_date = ledger_df.loc[li, 'date']
            l_amt = ledger_df.loc[li, 'normalized_amount']
            l_summary = str(ledger_df.loc[li, 'normalized_summary'])

            date_diff = abs((b_date - l_date).days)
            if date_diff > days:
                continue

            amt_diff = abs(b_amt - l_amt)
            if amt_diff > tolerance:
                continue

            # Compute fuzzy score — use max of token_sort and partial_ratio
            # for robustness with both Chinese and token-based matching
            score = max(
                fuzz.token_sort_ratio(b_summary, l_summary),
                fuzz.partial_ratio(b_summary, l_summary),
            )
            if score < threshold:
                continue

            candidates.append({
                'bank_idx': bi,
                'ledger_idx': li,
                'amount_diff': amt_diff,
                'date_diff': date_diff,
                'score': float(score),
            })

    if not candidates:
        return [], bank_df.copy(), ledger_df.copy()

    # Greedy matching: highest score first
    candidates.sort(key=lambda c: c['score'], reverse=True)

    matched_bank = set()
    matched_ledger = set()
    matched_records = []

    for c in candidates:
        if c['bank_idx'] in matched_bank or c['ledger_idx'] in matched_ledger:
            continue

        b_amt = bank_df.loc[c['bank_idx'], 'normalized_amount']
        l_amt = ledger_df.loc[c['ledger_idx'], 'normalized_amount']

        record = MatchRecord(
            bank_idx=c['bank_idx'],
            ledger_idx=c['ledger_idx'],
            match_type='fuzzy',
            amount_diff=c['amount_diff'],
            date_diff=c['date_diff'],
            score=c['score'],
            bank_amount=float(b_amt),
            ledger_amount=float(l_amt),
        )
        matched_records.append(record)
        matched_bank.add(c['bank_idx'])
        matched_ledger.add(c['ledger_idx'])

    unmatched_bank = bank_df.loc[~bank_df.index.isin(matched_bank)].copy()
    unmatched_ledger = ledger_df.loc[~ledger_df.index.isin(matched_ledger)].copy()

    return matched_records, unmatched_bank, unmatched_ledger


# ═══════════════════════════════════════════════════════════════════════
# 3. split_match
# ═══════════════════════════════════════════════════════════════════════

def split_match(bank_df, ledger_df, config):
    """Match a single bank transaction to a combination of 2-5 ledger entries.

    For each bank transaction, search within SPLIT_MATCH_DAYS for combinations
    of ledger entries that sum to the bank amount (within tolerance * n).

    Returns:
        (matched_records, unmatched_bank_df, unmatched_ledger_df)
    """
    days = _get_config_value(config, 'SPLIT_MATCH_DAYS', SPLIT_MATCH_DAYS)
    tolerance = _get_config_value(config, 'AMOUNT_TOLERANCE', AMOUNT_TOLERANCE)
    max_comb = 5  # Max ledger entries per split

    if bank_df.empty or ledger_df.empty:
        return [], bank_df.copy(), ledger_df.copy()

    matched_bank = set()
    matched_ledger = set()
    matched_records = []

    for bi in bank_df.index:
        if bi in matched_bank:
            continue

        b_date = bank_df.loc[bi, 'date']
        b_amt = bank_df.loc[bi, 'normalized_amount']

        # Find candidate ledger entries within date window
        cand_indices = []
        cand_amounts = []
        for li in ledger_df.index:
            if li in matched_ledger:
                continue
            l_date = ledger_df.loc[li, 'date']
            if abs((b_date - l_date).days) <= days:
                cand_indices.append(li)
                cand_amounts.append(ledger_df.loc[li, 'normalized_amount'])

        if len(cand_indices) < 2:
            continue  # Need at least 2 entries to split

        # Try combinations of 2..min(max_comb, len(candidates))
        best_match = None
        best_n = None

        for n in range(2, min(max_comb, len(cand_indices)) + 1):
            for combo_indices in itertools.combinations(range(len(cand_indices)), n):
                combo_amt = sum(cand_amounts[i] for i in combo_indices)
                diff = abs(b_amt - combo_amt)
                allowed = tolerance * n

                if diff <= allowed:
                    # Found a match — prefer smaller n and smaller diff
                    if best_match is None or (n < best_n) or (n == best_n and diff < best_match['diff']):
                        best_match = {
                            'ledger_indices': [cand_indices[i] for i in combo_indices],
                            'combo_amt': combo_amt,
                            'diff': diff,
                        }
                        best_n = n

        if best_match is not None:
            li_list = best_match['ledger_indices']
            # Find earliest date among matched ledgers
            earliest_date = min(ledger_df.loc[li, 'date'] for li in li_list)
            date_diff = abs((b_date - earliest_date).days)

            record = MatchRecord(
                bank_idx=bi,
                ledger_idx=li_list,
                match_type='split',
                amount_diff=best_match['diff'],
                date_diff=date_diff,
                score=100.0,
                bank_amount=float(b_amt),
                ledger_amount=float(best_match['combo_amt']),
            )
            matched_records.append(record)
            matched_bank.add(bi)
            for li in li_list:
                matched_ledger.add(li)

    unmatched_bank = bank_df.loc[~bank_df.index.isin(matched_bank)].copy()
    unmatched_ledger = ledger_df.loc[~ledger_df.index.isin(matched_ledger)].copy()

    return matched_records, unmatched_bank, unmatched_ledger


# ═══════════════════════════════════════════════════════════════════════
# 4. detect_duplicates
# ═══════════════════════════════════════════════════════════════════════

def detect_duplicates(df, source, config):
    """Detect potential duplicate transactions within the same source.

    Two transactions are flagged as duplicates if:
    - Same source (bank or ledger)
    - Date within ±DUPLICATE_CHECK_DAYS
    - Amount difference ≤ AMOUNT_TOLERANCE

    Returns:
        DataFrame with 'is_duplicate' (bool) and 'duplicate_group_id' (int) columns added.
        Non-duplicates get duplicate_group_id = -1.
    """
    days = _get_config_value(config, 'DUPLICATE_CHECK_DAYS', DUPLICATE_CHECK_DAYS)
    tolerance = _get_config_value(config, 'AMOUNT_TOLERANCE', AMOUNT_TOLERANCE)

    df = df.copy()
    df['is_duplicate'] = False
    df['duplicate_group_id'] = -1

    n = len(df)
    if n < 2:
        return df

    # Union-find for grouping
    parent = list(range(n))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(x, y):
        rx, ry = find(x), find(y)
        if rx != ry:
            parent[rx] = ry

    # Compare all pairs
    for i in range(n):
        for j in range(i + 1, n):
            d_i = df.iloc[i]['date']
            d_j = df.iloc[j]['date']
            if abs((d_i - d_j).days) > days:
                continue

            amt_i = df.iloc[i]['normalized_amount']
            amt_j = df.iloc[j]['normalized_amount']
            if abs(amt_i - amt_j) > tolerance:
                continue

            union(i, j)

    # Assign group IDs
    group_map = {}
    next_gid = 0
    for i in range(n):
        root = find(i)
        if parent[i] != i or any(find(j) == i for j in range(n) if j != i):
            # This row is part of a group (size ≥ 2)
            if root not in group_map:
                group_map[root] = next_gid
                next_gid += 1
            df.loc[i, 'is_duplicate'] = True
            df.loc[i, 'duplicate_group_id'] = group_map[root]

    # Ensure duplicate_group_id is int type
    df['duplicate_group_id'] = df['duplicate_group_id'].astype(int)

    return df
