import time
import pandas as pd
import numpy as np
from tqdm import tqdm

# Start timing the entire execution
start_time = time.time()
print("runnintransform_betas.py")

# Load raw beta data from Excel
df = pd.read_excel('raw_betas.xlsx', sheet_name='RawBetas')

# Separate the metadata (first 2 columns)
meta_data = df.iloc[:, :2]
# Remaining columns are the beta values
beta_data = df.iloc[:, 2:]

# Standardize exactly like Excel's (x - mean) / STDEV.S
print("Standardizing data using Excel's formula: (x - mean) / stdev.s...")
standardized_data = pd.DataFrame().reindex_like(beta_data)
for col in tqdm(beta_data.columns, desc="Standardizing columns"):
    vals = beta_data[col].to_numpy()
    mu = np.mean(vals)
    sigma = np.std(vals, ddof=1)   # ddof=1 matches Excel's STDEV.S
    standardized_data[col] = (vals - mu) / sigma

# Save standardized data
standardized_output = pd.concat([meta_data, standardized_data], axis=1)
standardized_output.to_excel(
    'standardized_betas.xlsx',
    index=False,
    sheet_name='Standardized Betas'
)
print("Standardized data saved to 'standardized_betas.xlsx'")

# Flatten and clean NaNs for percentile calculations
flat_vals = standardized_data.values.flatten()
flat_vals = flat_vals[~np.isnan(flat_vals)]

# Pre-sort for fast rank lookups
sorted_vals = np.sort(flat_vals)
n = len(sorted_vals)

# Transform using Excel's PERCENTRANK.EXC replication
print("Transforming standardized data to percentiles (Excel PERCENTRANK.EXC)...")
transformed_data = pd.DataFrame().reindex_like(standardized_data)

for col in tqdm(standardized_data.columns, desc="Transforming columns"):
    mask = ~standardized_data[col].isna()
    if mask.any():
        vals = standardized_data.loc[mask, col].to_numpy()
        # rank = count of values ≤ x
        ranks = np.searchsorted(sorted_vals, vals, side='right')
        # exclusive percent rank: round(rank/(n+1) - 0.0005, 3)
        pr_exc = np.round(ranks / (n + 1) - 0.0005, 3)
        # convert to percent (0–100) then shift/scale
        transformed_data.loc[mask, col] = (pr_exc * 100 - 50.5) / 34
        transformed_data = transformed_data.round(9)

# Combine and save final output
final_data = pd.concat([meta_data, transformed_data], axis=1)
final_data.to_excel(
    'transformed_betas.xlsx',
    index=False,
    sheet_name='Transformed Betas'
)

# Report timing
end_time = time.time()
mins, secs = divmod(end_time - start_time, 60)
print(f"Total execution time: {int(mins)} minutes and {secs:.2f} seconds")
