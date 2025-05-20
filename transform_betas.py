import time
import pandas as pd
import numpy as np
from tqdm import tqdm
from scipy.stats import percentileofscore

# Start timing the entire execution
start_time = time.time()

print("runnintransform_betas.py")
# Load raw beta data from excel

df = pd.read_excel('raw_betas.xlsx', sheet_name='RawBetas')

# Seperate the metadata (first 2 columns)
meta_data = df.iloc[:, :2]

# Remaining columns are the beta values
beta_data = df.iloc[:, 2:]

# Replace zscore with Excel's exact standardization formula
print("Standardizing data using Excel's formula: (x - mean) / stdev.s...")
standardized_data = pd.DataFrame().reindex_like(beta_data)
for col in tqdm(beta_data.columns, desc="Standardizing columns"):
    # Get column values
    values = beta_data[col].values
    # Calculate mean and sample standard deviation (ddof=1 for sample)
    mean = np.mean(values) # type: ignore
    std = np.std(values, ddof=1)  # type: ignore # ddof=1 matches Excel's STDEV.S
    # Apply standardization formula
    standardized_data[col] = (values - mean) / std

# Save the standardized data to a separate file
standardized_output = pd.concat([meta_data, standardized_data], axis=1)
standardized_output.to_excel('standardized_betas.xlsx', index=False, sheet_name='Standardized Betas')
print("Standardized data saved to 'standardized_betas.xlsx'")

# Create flat list of standardized betas
flat_standardized_data = standardized_data.values.flatten()
flat_standardized_data = flat_standardized_data[~np.isnan(flat_standardized_data)]
#print first 100 vals in flat_standardized_data
print(flat_standardized_data[:100])
# Transform standardize betas to scaled percentiles
def transform_to_percentiles(value):
    # If values are NaN, return NaN to preserve structure
    if pd.isna(value):
        return np.nan
    # Calculate the percentile: where does value fall in all values
    percentile = percentileofscore(flat_standardized_data, value, kind='mean')
    # Shift and scale
    return (percentile - 50.5) / 34

# Apply the transformation to the standardized matrix and give current updates on the progress fo teh process
print("Transforming standardized data to percentiles...")

# Pre-sort the flat standardized data once for faster percentile calculation
sorted_data = np.sort(flat_standardized_data)
n = len(sorted_data)

# Vectorized column-by-column approach with optimized percentile calculation
transformed_data = pd.DataFrame().reindex_like(standardized_data)
for col in tqdm(standardized_data.columns, desc="Transforming to percentiles"):
    mask = ~standardized_data[col].isna()
    if mask.any():
        # Get non-NaN values for this column
        values = standardized_data.loc[mask, col].to_numpy()
        
        # Fast percentile calculation using binary search (mimics percentileofscore with kind='mean')
        indices_right = np.searchsorted(sorted_data, values, side='right')
        indices_left = np.searchsorted(sorted_data, values, side='left')
        avg_indices = (indices_left + indices_right) / 2.0
        percentiles = avg_indices * 100.0 / n
        
        # Apply transformation formula
        transformed_data.loc[mask, col] = (percentiles - 50.5) / 34

# Combine the metadata with the transformed data
final_data = pd.concat([meta_data, transformed_data], axis=1)

# Save the transformed data to a new excel file
final_data.to_excel('transformed_betas2.xlsx', index=False, sheet_name='Transformed Betas')

# Calculate and print the total execution time
end_time = time.time()
total_time = end_time - start_time
minutes, seconds = divmod(total_time, 60)
print(f"Total execution time: {int(minutes)} minutes and {seconds:.2f} seconds")
