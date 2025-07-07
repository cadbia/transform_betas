import pandas as pd
import numpy as np

def excel_percentrank_exc(sorted_arr, x):
    """
    Replicates Excel's PERCENTRANK.EXC behavior:
      - returns #N/A if x <= min or x >= max
      - otherwise finds k such that sorted_arr[k] <= x <= sorted_arr[k+1]
        and interpolates:
          P = (k + (x - sorted_arr[k]) / (sorted_arr[k+1] - sorted_arr[k])) 
              / (n + 1)
    """
    n = len(sorted_arr)
    if n < 2:
        return np.nan
    # Exclusive: values strictly beyond endpoints are invalid
    if x < sorted_arr[0] or x > sorted_arr[-1]:
        return np.nan
    
    # Handle exact boundary cases
    if x == sorted_arr[0]:
        return 1 / (n + 1)
    if x == sorted_arr[-1]:
        return n / (n + 1)

    # find insertion index so that sorted_arr[idx-1] < x <= sorted_arr[idx]
    idx = np.searchsorted(sorted_arr, x, side="right")
    k = idx - 1
    xk, xk1 = sorted_arr[k], sorted_arr[k+1]
    # interpolate between positions k and k+1
    fraction = (x - xk) / (xk1 - xk)
    # divide by (n+1) for exclusive rank
    return (k + fraction) / (n + 1)

def main():
    # 1) Load data
    df = pd.read_excel("raw_betas.xlsx", sheet_name="Sheet1")
    meta = df.iloc[:, :2]            # ticker & name
    betas = df.iloc[:, 2:].astype(float)

    # 2) Standardize each column (ddof=1)
    z = betas.copy()
    for col in z.columns:
        vals = betas[col].values
        mu = np.nanmean(vals)
        sigma = np.nanstd(vals, ddof=1)
        z[col] = (vals - mu) / sigma

    # 3) Flatten & sort all z-scores
    flat = z.values.flatten()
    flat = flat[~np.isnan(flat)]
    sorted_flat = np.sort(flat)

    # 4) Compute EXCLUSIVE percentiles and transform
    def transform_exc(val):
        if np.isnan(val):
            return np.nan
        p_exc = excel_percentrank_exc(sorted_flat, val) * 100.0   # 0–100 scale
        return (p_exc - 50.5) / 34.0

    transformed = z.applymap(transform_exc)

    # 5) Write out both sheets
    with pd.ExcelWriter("transformed_factor_betas_07_07_2025.xlsx", engine="openpyxl") as w:
        pd.concat([meta, z], axis=1).to_excel(
            w, sheet_name="StandardizedBetas", index=False
        )
        pd.concat([meta, transformed], axis=1).to_excel(
            w, sheet_name="TransformedBetas", index=False
        )

    # 6) Verify no blank cells in transformed data
    print("Data validation:")
    print(f"Original betas shape: {betas.shape}")
    print(f"Standardized betas - blank cells: {z.isna().sum().sum()}")
    print(f"Transformed betas - blank cells: {transformed.isna().sum().sum()}")
    
    if transformed.isna().sum().sum() > 0:
        print("WARNING: Found blank cells in transformed data!")
        blank_locations = transformed.isna()
        for col in transformed.columns:
            if blank_locations[col].any():
                blank_rows = blank_locations[col][blank_locations[col]].index.tolist()
                print(f"  Column '{col}': rows {blank_rows}")
    else:
        print("✓ All transformed beta cells are filled")
    
    print("Processing complete.")

if __name__ == "__main__":
    main()
