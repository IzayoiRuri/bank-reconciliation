from src.parser import load_and_parse
from src.normalizer import clean_dataframe

b, l = load_and_parse('data/银行流水.xlsx', 'data/公司日记账.xls')
b = clean_dataframe(b, 'bank')
l = clean_dataframe(l, 'ledger')

print(f'Bank: {len(b)} rows, Ledger: {len(l)} rows')
print(f'Bank normalized_summary non-empty: {(b["normalized_summary"] != "").mean():.2%}')
print(f'Ledger normalized_summary non-empty: {(l["normalized_summary"] != "").mean():.2%}')
print()
print('=== Bank Sample ===')
print(b[['summary', 'normalized_summary']].head(5).to_string())
print()
print('=== Ledger Sample ===')
print(l[['summary', 'normalized_summary']].head(5).to_string())
