"""银行对账工具 — SQLite 历史存储

记录每次对账结果并支持回溯查询。
"""

import sqlite3
import os
from datetime import datetime

import config as _default_config


# ═══════════════════════════════════════════════════════════════════════
# Database table creation
# ═══════════════════════════════════════════════════════════════════════

CREATE_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS reconciliations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    reconciled_at TEXT NOT NULL,
    bank_file TEXT NOT NULL,
    ledger_file TEXT NOT NULL,
    date_range_start TEXT,
    date_range_end TEXT,
    bank_count INTEGER,
    ledger_count INTEGER,
    bank_total_income REAL,
    bank_total_expense REAL,
    ledger_total_income REAL,
    ledger_total_expense REAL,
    bank_opening_balance REAL,
    bank_closing_balance REAL,
    matched_count INTEGER,
    exact_matched INTEGER,
    fuzzy_matched INTEGER,
    split_matched INTEGER,
    unmatched_bank_count INTEGER,
    unmatched_ledger_count INTEGER,
    match_rate REAL,
    amount_diff REAL,
    report_path TEXT,
    notes TEXT
);
"""


# ═══════════════════════════════════════════════════════════════════════
# _resolve_db_path
# ═══════════════════════════════════════════════════════════════════════

def _resolve_db_path(db_path=None):
    """Return the database path to use. Falls back to config.DB_PATH."""
    if db_path is not None:
        return db_path
    return _default_config.DB_PATH


# ═══════════════════════════════════════════════════════════════════════
# _dict_from_row
# ═══════════════════════════════════════════════════════════════════════

def _dict_from_row(row, description):
    """Convert a sqlite3.Row to a plain dict using cursor description."""
    if row is None:
        return None
    return dict(zip([col[0] for col in description], row))


# ═══════════════════════════════════════════════════════════════════════
# init_db
# ═══════════════════════════════════════════════════════════════════════

def init_db(db_path=None):
    """Initialize the database — creates the reconciliations table if it
    does not exist. Idempotent.

    Args:
        db_path: Path to the SQLite database file. Defaults to config.DB_PATH.

    Returns:
        sqlite3.Connection: An open connection to the database.
    """
    path = _resolve_db_path(db_path)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    conn.execute(CREATE_TABLE_SQL)
    conn.commit()
    return conn


# ═══════════════════════════════════════════════════════════════════════
# save_reconciliation
# ═══════════════════════════════════════════════════════════════════════

def save_reconciliation(result, report_path='', notes='', db_path=None):
    """Save a ReconciliationResult to the history database.

    Args:
        result: A ReconciliationResult instance (from pipeline.py).
        report_path: Optional path to the generated report file.
        notes: Optional notes/remarks about this reconciliation run.
        db_path: Optional database path override.

    Returns:
        int: The ID of the inserted row.
    """
    conn = init_db(db_path=db_path)

    date_range_start = result.date_range[0] if result.date_range else None
    date_range_end = result.date_range[1] if result.date_range else None

    sql = """
    INSERT INTO reconciliations (
        reconciled_at, bank_file, ledger_file,
        date_range_start, date_range_end,
        bank_count, ledger_count,
        bank_total_income, bank_total_expense,
        ledger_total_income, ledger_total_expense,
        bank_opening_balance, bank_closing_balance,
        matched_count, exact_matched, fuzzy_matched, split_matched,
        unmatched_bank_count, unmatched_ledger_count,
        match_rate, amount_diff,
        report_path, notes
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    cursor = conn.execute(sql, (
        result.reconciled_at,
        result.bank_file,
        result.ledger_file,
        date_range_start,
        date_range_end,
        result.bank_total_count,
        result.ledger_total_count,
        result.bank_total_income,
        result.bank_total_expense,
        result.ledger_total_income,
        result.ledger_total_expense,
        result.bank_opening_balance,
        result.bank_closing_balance,
        result.matched_count,
        result.exact_matched,
        result.fuzzy_matched,
        result.split_matched,
        result.unmatched_bank_count,
        result.unmatched_ledger_count,
        result.match_rate,
        result.matched_amount_diff,
        report_path,
        notes,
    ))

    conn.commit()
    row_id = cursor.lastrowid
    conn.close()
    return row_id


# ═══════════════════════════════════════════════════════════════════════
# list_history
# ═══════════════════════════════════════════════════════════════════════

def list_history(limit=20, db_path=None):
    """List recent reconciliation records, newest first.

    Args:
        limit: Maximum number of records to return (default: 20).
        db_path: Optional database path override.

    Returns:
        list[dict]: List of records with keys matching the DB columns.
    """
    conn = init_db(db_path=db_path)

    cursor = conn.execute(
        "SELECT * FROM reconciliations ORDER BY id DESC LIMIT ?",
        (limit,),
    )
    rows = cursor.fetchall()
    results = [_dict_from_row(r, cursor.description) for r in rows]
    conn.close()
    return results


# ═══════════════════════════════════════════════════════════════════════
# get_history
# ═══════════════════════════════════════════════════════════════════════

def get_history(record_id, db_path=None):
    """Retrieve a single reconciliation record by ID.

    Args:
        record_id: The integer ID of the record.
        db_path: Optional database path override.

    Returns:
        dict or None: The record as a dictionary, or None if not found.
    """
    conn = init_db(db_path=db_path)

    cursor = conn.execute(
        "SELECT * FROM reconciliations WHERE id = ?",
        (record_id,),
    )
    row = cursor.fetchone()
    conn.close()
    return _dict_from_row(row, cursor.description)


# ═══════════════════════════════════════════════════════════════════════
# delete_history
# ═══════════════════════════════════════════════════════════════════════

def delete_history(record_id, db_path=None):
    """Delete a reconciliation record by ID.

    Args:
        record_id: The integer ID of the record to delete.
        db_path: Optional database path override.

    Returns:
        bool: True if a record was deleted, False if it didn't exist.
    """
    conn = init_db(db_path=db_path)

    cursor = conn.execute(
        "DELETE FROM reconciliations WHERE id = ?",
        (record_id,),
    )
    conn.commit()
    deleted = cursor.rowcount > 0
    conn.close()
    return deleted


# ═══════════════════════════════════════════════════════════════════════
# search_history
# ═══════════════════════════════════════════════════════════════════════

def search_history(keyword, db_path=None):
    """Search reconciliation history by keyword in bank_file or ledger_file.

    Search is case-insensitive and uses LIKE for substring matching.

    Args:
        keyword: The search keyword.
        db_path: Optional database path override.

    Returns:
        list[dict]: Matching records, newest first.
    """
    conn = init_db(db_path=db_path)

    pattern = f'%{keyword}%'
    cursor = conn.execute(
        """SELECT * FROM reconciliations
           WHERE bank_file LIKE ? COLLATE NOCASE
              OR ledger_file LIKE ? COLLATE NOCASE
           ORDER BY id DESC""",
        (pattern, pattern),
    )
    rows = cursor.fetchall()
    results = [_dict_from_row(r, cursor.description) for r in rows]
    conn.close()
    return results
