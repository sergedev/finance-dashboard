#!/usr/bin/env python3
"""
Convert CSV files to a single Excel file with multiple sheets
"""
import pandas as pd

# Read the CSV files
transactions_df = pd.read_csv('data/transactions.2025.csv')
categories_df = pd.read_csv('data/categories.2025.csv')
snapshots_df = pd.read_csv('data/snapshots.csv')

# Create Excel file with three sheets
with pd.ExcelWriter('data/finance.2025.xlsx', engine='openpyxl') as writer:
    transactions_df.to_excel(writer, sheet_name='Transactions', index=False)
    categories_df.to_excel(writer, sheet_name='Categories', index=False)
    snapshots_df.to_excel(writer, sheet_name='Snapshots', index=False)

print("✓ Created data/finance.2025.xlsx with 3 sheets:")
print(f"  - Transactions: {len(transactions_df)} rows")
print(f"  - Categories: {len(categories_df)} rows")
print(f"  - Snapshots: {len(snapshots_df)} rows")
