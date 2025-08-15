import pandas as pd
import numpy as np
import re
from pathlib import Path
from datetime import datetime

# -----------------------------
# File / naming configuration
# -----------------------------
# Input Excel file (first two cols = metadata, rest = betas)
INPUT_FILE = "raw_betas.xlsx"
# Sheet containing the raw betas
INPUT_SHEET = "Sheet1"
# Output filename prefix; a date tag (derived from input) will be appended
OUTPUT_PREFIX = "transformed_factor_betas"

def _extract_date_tag(path: Path) -> str:
    """
    Extract a date from the filename (preferred) or fallback to file mtime / today.
    Recognized patterns inside the filename stem:
      YYYYMMDD
      YYYY-MM-DD
      YYYY_MM_DD
      MM-DD-YYYY
      MM_DD_YYYY
    Returns YYYYMMDD.
    """
    stem = path.stem
    # Try YYYYMMDD
    m = re.search(r"(20\d{2}[01]\d[0-3]\d)", stem)
    if m:
        raw = m.group(1)
        try:
            return datetime.strptime(raw, "%Y%m%d").strftime("%Y%m%d")
        except ValueError:
            pass
    # Try YYYY[-_]MM[-_]DD
    m = re.search(r"(20\d{2})[-_](\d{2})[-_](\d{2})", stem)
    if m:
        y, mo, d = m.groups()
        try:
            return datetime(int(y), int(mo), int(d)).strftime("%Y%m%d")
        except ValueError:
            pass
    # Try MM[-_]DD[-_]YYYY
    m = re.search(r"(\d{2})[-_](\d{2})[-_](20\d{2})", stem)
    if m:
        mo, d, y = m.groups()
        try:
            return datetime(int(y), int(mo), int(d)).strftime("%Y%m%d")
        except ValueError:
            pass
    # Fallback: file modified time or today
    try:
        return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y%m%d")
    except Exception:
        return datetime.today().strftime("%Y%m%d")

def _build_output_filename(input_file: str) -> str:
    date_tag = _extract_date_tag(Path(input_file))
    return f"{OUTPUT_PREFIX}_{date_tag}.xlsx"

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
    input_path = Path(INPUT_FILE)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    output_file = _build_output_filename(INPUT_FILE)
    print(f"Reading input: {input_path.name} (sheet='{INPUT_SHEET}')")
    print(f"Output will be: {output_file}")
    df = pd.read_excel(input_path, sheet_name=INPUT_SHEET)
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
    with pd.ExcelWriter(output_file, engine="openpyxl") as w:
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
    
    print(f"Processing complete. Output file: {output_file}")

if __name__ == "__main__":
    main()
