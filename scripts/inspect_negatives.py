import pandas as pd

input_path = "/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/output/input file/Input Apr'26.xlsx"
df = pd.read_excel(input_path)

recharge_cols = [
    'Recharge  General Cost - Payroll',
    'Recharge  General Cost - Manager',
    'Recharge  General Cost - Leadership cost',
    'Recharge  General Cost - Desk Cost',
    'Recharge  General Cost - Retirals',
    'Mark up'
]

print("Scanning for negative values in input file...")
negative_counts = {}
for col in recharge_cols:
    if col in df.columns:
        nums = pd.to_numeric(df[col], errors='coerce').fillna(0)
        neg = nums[nums < 0]
        negative_counts[col] = len(neg)
        if len(neg) > 0:
            print(f"Col '{col}' has {len(neg)} negative values. First 5:")
            print(neg.head(5))

print("Negative counts by column:", negative_counts)
