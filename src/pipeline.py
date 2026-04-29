"""银行对账工具 — 对账管道

编排 4 阶段匹配流程并生成汇总统计。
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional, Union

import pandas as pd

import config as _default_config
from parser import load_and_parse
from normalizer import clean_dataframe
from matcher import (
    MatchRecord,
    exact_match,
    fuzzy_match,
    split_match,
    detect_duplicates,
)


# ═══════════════════════════════════════════════════════════════════════
# ReconciliationResult
# ═══════════════════════════════════════════════════════════════════════

@dataclass
class ReconciliationResult:
    """对账结果"""
    # 输入摘要
    bank_file: str
    ledger_file: str
    bank_total_count: int
    ledger_total_count: int
    bank_total_income: float       # 银行总收入
    bank_total_expense: float      # 银行总支出
    ledger_total_income: float     # 日记账总收入
    ledger_total_expense: float    # 日记账总支出
    bank_opening_balance: float    # 银行期初余额
    bank_closing_balance: float    # 银行期末余额

    # 匹配统计
    matched_count: int             # 总匹配数
    exact_matched: int             # 精确匹配数
    fuzzy_matched: int             # 模糊匹配数
    split_matched: int             # 拆分匹配数

    # 差异
    unmatched_bank_count: int      # 银行独有
    unmatched_ledger_count: int    # 日记账独有
    duplicate_bank_count: int      # 银行重复
    duplicate_ledger_count: int    # 日记账重复

    # 金额差异
    total_bank_amount: float       # 银行交易净额（收入-支出）
    total_ledger_amount: float     # 日记账交易净额
    matched_amount_diff: float     # 已匹配交易金额差
    unmatched_bank_amount: float   # 未匹配银行金额
    unmatched_ledger_amount: float # 未匹配日记账金额

    # 详细记录
    matched_records: list          # MatchRecord 列表
    unmatched_bank: pd.DataFrame   # 银行独有交易
    unmatched_ledger: pd.DataFrame # 日记账独有交易
    bank_duplicates: pd.DataFrame  # 银行重复交易
    ledger_duplicates: pd.DataFrame # 日记账重复交易

    # 元数据
    reconciled_at: str             # 对账时间
    date_range: tuple              # (min_date, max_date)
    match_rate: float              # 匹配率

    # 原始数据（用于报告生成，可选）
    bank_df: object = None         # 银行原始DataFrame
    ledger_df: object = None       # 日记账原始DataFrame


# ═══════════════════════════════════════════════════════════════════════
# compare_amounts
# ═══════════════════════════════════════════════════════════════════════

def compare_amounts(bank_df, ledger_df):
    """比较银行和日记账的总收入、总支出、净额。

    Args:
        bank_df: 银行 DataFrame（含 credit, debit 列）
        ledger_df: 日记账 DataFrame（含 credit, debit 列）

    Returns:
        dict: {
            bank_income, bank_expense, bank_net,
            ledger_income, ledger_expense, ledger_net,
            diff
        }
    """
    # 银行：收入=贷方(credit)，支出=借方(debit)
    bank_income = round(float(bank_df['credit'].sum()), 2) if 'credit' in bank_df.columns else 0.0
    bank_expense = round(float(bank_df['debit'].sum()), 2) if 'debit' in bank_df.columns else 0.0
    bank_net = round(bank_income - bank_expense, 2)

    # 日记账：收入=借方(debit)，支出=贷方(credit)
    ledger_income = round(float(ledger_df['debit'].sum()), 2) if 'debit' in ledger_df.columns else 0.0
    ledger_expense = round(float(ledger_df['credit'].sum()), 2) if 'credit' in ledger_df.columns else 0.0
    ledger_net = round(ledger_income - ledger_expense, 2)

    diff = round(bank_net - ledger_net, 2)

    return {
        'bank_income': bank_income,
        'bank_expense': bank_expense,
        'bank_net': bank_net,
        'ledger_income': ledger_income,
        'ledger_expense': ledger_expense,
        'ledger_net': ledger_net,
        'diff': diff,
    }


# ═══════════════════════════════════════════════════════════════════════
# _resolve_config
# ═══════════════════════════════════════════════════════════════════════

def _resolve_config(config_module=None):
    """Return the config module/object to use. Falls back to default config."""
    if config_module is not None:
        return config_module
    return _default_config


# ═══════════════════════════════════════════════════════════════════════
# run_reconciliation
# ═══════════════════════════════════════════════════════════════════════

def run_reconciliation(bank_path, ledger_path, config_module=None):
    """执行完整对账流程。

    流程:
        1. 加载并解析银行流水和日记账
        2. 标准化摘要文本
        3. 检测内部重复
        4. 4阶段匹配：精确 → 模糊 → 拆分
        5. 汇总统计

    Args:
        bank_path: 银行流水 .xlsx 文件路径
        ledger_path: 公司日记账 .xls 文件路径
        config_module: 可选的配置模块（默认使用 config.py）

    Returns:
        ReconciliationResult
    """
    cfg = _resolve_config(config_module)

    # ── 1. Load & parse ──────────────────────────────────────────
    bank_df, ledger_df = load_and_parse(bank_path, ledger_path)

    # ── 2. Clean / normalize summaries ───────────────────────────
    bank_df = clean_dataframe(bank_df, 'bank')
    ledger_df = clean_dataframe(ledger_df, 'ledger')

    # ── 3. Detect duplicates ─────────────────────────────────────
    bank_with_dups = detect_duplicates(bank_df, 'bank', cfg)
    ledger_with_dups = detect_duplicates(ledger_df, 'ledger', cfg)

    bank_duplicates = bank_with_dups[bank_with_dups['is_duplicate']].copy()
    ledger_duplicates = ledger_with_dups[ledger_with_dups['is_duplicate']].copy()

    # ── 4. 4-stage matching ──────────────────────────────────────
    # Stage 1: exact
    exact_records, bank_rem, ledger_rem = exact_match(bank_df, ledger_df, cfg)

    # Stage 2: fuzzy (on remaining)
    fuzzy_records, bank_rem, ledger_rem = fuzzy_match(bank_rem, ledger_rem, cfg)

    # Stage 3: split (on remaining)
    split_records, bank_unmatched, ledger_unmatched = split_match(bank_rem, ledger_rem, cfg)

    # Collect all matched records
    all_matched = exact_records + fuzzy_records + split_records

    # ── 5. Compute summary statistics ────────────────────────────

    # Amount comparison
    amount_cmp = compare_amounts(bank_df, ledger_df)

    # Opening / closing balances
    if len(bank_df) > 0:
        bank_opening = round(float(
            bank_df['balance'].iloc[0] - bank_df['normalized_amount'].iloc[0]
        ), 2)
        bank_closing = round(float(bank_df['balance'].iloc[-1]), 2)
    else:
        bank_opening = 0.0
        bank_closing = 0.0

    # Date range
    all_dates = pd.concat([
        bank_df['date'].dropna(),
        ledger_df['date'].dropna(),
    ])
    if len(all_dates) > 0:
        min_date = all_dates.min().strftime('%Y-%m-%d')
        max_date = all_dates.max().strftime('%Y-%m-%d')
        date_range = (min_date, max_date)
    else:
        date_range = ('', '')

    # Match rate
    max_count = max(len(bank_df), len(ledger_df))
    if max_count > 0:
        match_rate = round(len(all_matched) / max_count, 4)
    else:
        match_rate = 1.0

    # Total bank/ledger net amounts
    total_bank_amount = round(float(bank_df['normalized_amount'].sum()), 2)
    total_ledger_amount = round(float(ledger_df['normalized_amount'].sum()), 2)

    # Matched amount diff (sum of absolute differences)
    matched_amount_diff = round(
        sum(abs(r.amount_diff) for r in all_matched), 2
    )

    # Unmatched amounts
    unmatched_bank_amount = round(
        float(bank_unmatched['normalized_amount'].sum()), 2
    ) if len(bank_unmatched) > 0 else 0.0
    unmatched_ledger_amount = round(
        float(ledger_unmatched['normalized_amount'].sum()), 2
    ) if len(ledger_unmatched) > 0 else 0.0

    # ── 6. Build and return result ───────────────────────────────

    return ReconciliationResult(
        bank_file=bank_path,
        ledger_file=ledger_path,
        bank_total_count=len(bank_df),
        ledger_total_count=len(ledger_df),
        bank_total_income=amount_cmp['bank_income'],
        bank_total_expense=amount_cmp['bank_expense'],
        ledger_total_income=amount_cmp['ledger_income'],
        ledger_total_expense=amount_cmp['ledger_expense'],
        bank_opening_balance=bank_opening,
        bank_closing_balance=bank_closing,
        matched_count=len(all_matched),
        exact_matched=len(exact_records),
        fuzzy_matched=len(fuzzy_records),
        split_matched=len(split_records),
        unmatched_bank_count=len(bank_unmatched),
        unmatched_ledger_count=len(ledger_unmatched),
        duplicate_bank_count=len(bank_duplicates),
        duplicate_ledger_count=len(ledger_duplicates),
        total_bank_amount=total_bank_amount,
        total_ledger_amount=total_ledger_amount,
        matched_amount_diff=matched_amount_diff,
        unmatched_bank_amount=unmatched_bank_amount,
        unmatched_ledger_amount=unmatched_ledger_amount,
        matched_records=all_matched,
        unmatched_bank=bank_unmatched,
        unmatched_ledger=ledger_unmatched,
        bank_duplicates=bank_duplicates,
        ledger_duplicates=ledger_duplicates,
        reconciled_at=datetime.now().isoformat(),
        date_range=date_range,
        match_rate=match_rate,
        bank_df=bank_with_dups,
        ledger_df=ledger_with_dups,
    )


# ═══════════════════════════════════════════════════════════════════════
# get_summary_text
# ═══════════════════════════════════════════════════════════════════════

def get_summary_text(result):
    """生成人类可读的对账摘要文本。

    Args:
        result: ReconciliationResult 实例

    Returns:
        str: 格式化的摘要文本
    """
    if not isinstance(result, ReconciliationResult):
        raise TypeError(
            f"Expected ReconciliationResult, got {type(result).__name__}"
        )

    lines = []
    lines.append("=" * 60)
    lines.append("  银行对账报告")
    lines.append("=" * 60)
    lines.append(f"  对账时间:    {result.reconciled_at}")
    lines.append(f"  银行文件:    {result.bank_file}")
    lines.append(f"  日记账文件:  {result.ledger_file}")
    lines.append(f"  日期范围:    {result.date_range[0]} ~ {result.date_range[1]}")
    lines.append("")

    lines.append("-" * 60)
    lines.append("  交易统计")
    lines.append("-" * 60)
    lines.append(f"  银行交易数:      {result.bank_total_count}")
    lines.append(f"  日记账交易数:    {result.ledger_total_count}")
    lines.append(f"  银行总收入:      {result.bank_total_income:,.2f}")
    lines.append(f"  银行总支出:      {result.bank_total_expense:,.2f}")
    lines.append(f"  日记账总收入:    {result.ledger_total_income:,.2f}")
    lines.append(f"  日记账总支出:    {result.ledger_total_expense:,.2f}")
    if result.bank_opening_balance or result.bank_closing_balance:
        lines.append(f"  银行期初余额:    {result.bank_opening_balance:,.2f}")
        lines.append(f"  银行期末余额:    {result.bank_closing_balance:,.2f}")
    lines.append("")

    lines.append("-" * 60)
    lines.append("  匹配结果")
    lines.append("-" * 60)
    lines.append(f"  匹配率:          {result.match_rate:.1%}")
    lines.append(f"  总匹配数:        {result.matched_count}")
    lines.append(f"    - 精确匹配:    {result.exact_matched}")
    lines.append(f"    - 模糊匹配:    {result.fuzzy_matched}")
    lines.append(f"    - 拆分匹配:    {result.split_matched}")
    lines.append(f"  银行未匹配:      {result.unmatched_bank_count}")
    lines.append(f"  日记账未匹配:    {result.unmatched_ledger_count}")
    lines.append("")

    lines.append("-" * 60)
    lines.append("  差异分析")
    lines.append("-" * 60)
    lines.append(f"  银行交易净额:    {result.total_bank_amount:,.2f}")
    lines.append(f"  日记账交易净额:  {result.total_ledger_amount:,.2f}")
    lines.append(f"  已匹配金额差:    {result.matched_amount_diff:,.2f}")
    lines.append(f"  未匹配银行金额:  {result.unmatched_bank_amount:,.2f}")
    lines.append(f"  未匹配日记账金额:{result.unmatched_ledger_amount:,.2f}")
    lines.append("")

    if result.duplicate_bank_count > 0 or result.duplicate_ledger_count > 0:
        lines.append("-" * 60)
        lines.append("  疑似重复")
        lines.append("-" * 60)
        lines.append(f"  银行重复:        {result.duplicate_bank_count}")
        lines.append(f"  日记账重复:      {result.duplicate_ledger_count}")
        lines.append("")

    lines.append("=" * 60)
    lines.append("  对账完成")
    lines.append("=" * 60)

    return '\n'.join(lines)
