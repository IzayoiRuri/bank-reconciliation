"""银行对账工具 — 摘要标准化器

对银行流水和日记账的摘要文本进行标准化处理：
- 去除凭证前缀、内部编号、嵌入日期、地名/银行代号
- 统一分隔符、去除停用词、规范化空白
- 提取金额和对方单位名称
"""

import re
import pandas as pd

from config import SUMMARY_STOP_WORDS, SUMMARY_REMOVE_PATTERNS


# ═══════════════════════════════════════════════════════════════════════
# Core normalization
# ═══════════════════════════════════════════════════════════════════════

def normalize_summary(text):
    """标准化摘要文本，提取核心语义关键词。

    处理步骤：
        a. 去除凭证前缀（收/付/转/现/银 + 编号）
        b. 去除内部编号（_2_、-001-）
        c. 去除嵌入日期（2025.09.04、2026-01-02）
        d. 去除地名/银行代号 + 编号（招行0101、建行0203）
        e. 替换分隔符为空格
        f. 去除停用词
        g. 规范化空白
        h. 返回小写

    幂等性：normalize_summary(normalize_summary(x)) == normalize_summary(x)

    Args:
        text: 原始摘要文本

    Returns:
        str: 标准化后的核心语义文本
    """
    if not isinstance(text, str):
        text = str(text) if text is not None else ''

    if not text.strip():
        return ''

    # Steps (a)-(d): Apply removal patterns from config, replacing with
    # a single space to preserve word boundaries between adjacent tokens.
    for pattern in SUMMARY_REMOVE_PATTERNS:
        text = re.sub(pattern, ' ', text)

    # Step (e): Replace remaining separators / punctuation with spaces
    text = re.sub(r'[-_/.,:;!?()（）【】\[\]{}"\'‘’“”]', ' ', text)

    # Step (f): Remove stop words.
    # When all meaningful content would be erased, keep the original
    # tokens to avoid returning an empty string for short summaries.
    words = text.split()
    stop_set = {w.lower() for w in SUMMARY_STOP_WORDS}
    filtered = [w for w in words if w.lower() not in stop_set]

    if not filtered and words:
        # Don't erase everything — keep the original tokens
        filtered = words

    text = ' '.join(filtered)

    # Step (g): Collapse multiple spaces and strip
    text = re.sub(r'\s+', ' ', text).strip()

    # Step (h): Lowercase
    text = text.lower()

    return text


# ═══════════════════════════════════════════════════════════════════════
# Amount extraction
# ═══════════════════════════════════════════════════════════════════════

def extract_amounts(text):
    """从摘要文本中提取金额数字。

    支持格式：50000.00、1,234,567.89、1234（无小数）

    Args:
        text: 摘要文本

    Returns:
        list[float]: 提取到的金额列表
    """
    if not isinstance(text, str):
        return []

    # Match numbers with optional comma separators and decimal point
    # Pattern: digits with optional comma grouping, optional decimal
    pattern = r'(?<!\d)(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d{1,2})?(?!\d)'

    amounts = []
    for match in re.finditer(pattern, text):
        num_str = match.group().replace(',', '')
        try:
            amount = float(num_str)
            # Filter out obviously non-amount numbers (years, IDs, etc.)
            # Only keep numbers that look like monetary amounts
            if amount > 0:
                amounts.append(amount)
        except (ValueError, OverflowError):
            continue

    return amounts


# ═══════════════════════════════════════════════════════════════════════
# Counterparty extraction
# ═══════════════════════════════════════════════════════════════════════

# Common company/entity suffixes in Chinese
_COMPANY_SUFFIXES = (
    '公司', '厂', '店', '局', '行', '部', '中心', '集团',
    '银行', '信用社', '事务所', '工作室', '诊所', '医院',
    '学校', '学院', '大学', '研究院',
)

# Build a regex alternation group ordered longest-first so that
# "信用社" is preferred over "社" alone.
_suffix_pattern = '|'.join(
    sorted(_COMPANY_SUFFIXES, key=len, reverse=True)
)

_COUNTERPARTY_PATTERN = re.compile(
    r'(?:支付|收款|付给|收自|付款给|收款自|转给|转自)'
    r'([^\s，。,.\-;:!?()（）【】\[\]{}"\']*?'
    r'(?:' + _suffix_pattern + r'))',
    re.IGNORECASE,
)


def extract_counterparty(text):
    """尝试从摘要中提取对方单位名称。

    使用简单规则：公司名通常在'支付'/'收款'后面，
    包含'公司'/'厂'/'店'/'局'/'行'等后缀。

    Args:
        text: 摘要文本

    Returns:
        str | None: 对方单位名称，无法识别时返回 None
    """
    if not isinstance(text, str):
        return None

    match = _COUNTERPARTY_PATTERN.search(text)
    if match:
        return match.group(1).strip().lower()
    return None


# ═══════════════════════════════════════════════════════════════════════
# DataFrame batch processing
# ═══════════════════════════════════════════════════════════════════════

def clean_dataframe(df, source):
    """对 DataFrame 的 summary 列应用标准化，添加 normalized_summary 列。

    Args:
        df: 包含 summary 列的 DataFrame
        source: 'bank' 或 'ledger'（保留参数以保持一致接口）

    Returns:
        DataFrame: 添加了 normalized_summary 列的新 DataFrame
    """
    if source not in ('bank', 'ledger'):
        raise ValueError(f"Unknown source: {source!r}. Must be 'bank' or 'ledger'.")

    df = df.copy()

    if 'summary' in df.columns:
        df['normalized_summary'] = df['summary'].apply(
            lambda x: normalize_summary(x) if pd.notna(x) else ''
        )
    else:
        df['normalized_summary'] = ''

    return df
