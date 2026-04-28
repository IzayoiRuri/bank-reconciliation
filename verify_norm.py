#!/usr/bin/env python3
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from parser import load_and_parse
from normalizer import clean_dataframe

b, l = load_and_parse(
    os.path.join(os.path.dirname(__file__), 'data/银行流水.xlsx'),
    os.path.join(os.path.dirname(__file__), 'data/公司日记账.xls')
)
b = clean_dataframe(b, 'bank')
l = clean_dataframe(l, 'ledger')

print(f'Bank: {len(b)} rows, Ledger: {len(l)} rows')
print(f'Bank normalized non-empty: {(b["normalized_summary"] != "").mean():.2%}')
print(f'Ledger normalized non-empty: {(l["normalized_summary"] != "").mean():.2%}')
print()
print('=== Bank Sample ===')
for _, row in b[['summary', 'normalized_summary']].head(8).iterrows():
    print(f"  [{row['normalized_summary'][:60]}]")
print()
print('=== Ledger Sample ===')
for _, row in l[['summary', 'normalized_summary']].head(8).iterrows():
    print(f"  [{row['normalized_summary'][:60]}]")
