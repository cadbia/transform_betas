import pandas as pd
import numpy as np
import re
from pathlib import Path
from datetime import datetime
import sys

# -----------------------------
# File / naming configuration
# -----------------------------
# Input file can be .xlsx/.xls or .csv/.txt
INPUT_FILE = "raw_betas.xlsx"
# Sheet containing the raw betas (used only for Excel inputs)
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


def _build_output_stem(input_file: str) -> str:
    date_tag = _extract_date_tag(Path(input_file))
    return f"{OUTPUT_PREFIX}_{date_tag}"


def _read_input(input_file: str, sheet_name: str) -> pd.DataFrame:
    """Read input file as DataFrame. Supports Excel (.xlsx/.xls/.xlsm) and CSV/TXT.
    For Excel, uses the provided sheet name; for CSV/TXT, reads the whole file.
    """
    path = Path(input_file)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    ext = path.suffix.lower()
    if ext in {".csv", ".txt"}:
        return pd.read_csv(path, low_memory=False)
    if ext in {".xlsx", ".xls", ".xlsm"}:
        return pd.read_excel(path, sheet_name=sheet_name)
    raise ValueError(f"Unsupported input file type: {ext}")


def _candidate_inputs(base_stem: str):
    return [
        f"{base_stem}.xlsx",
        f"{base_stem}.xls",
        f"{base_stem}.xlsm",
        f"{base_stem}.csv",
        f"{base_stem}.txt",
    ]


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
    # Allow passing the input file via CLI: python transform_betas.py <input_file>
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        preferred = Path(INPUT_FILE)
        if preferred.exists():
            input_file = str(preferred)
        else:
            base_stem = preferred.stem if preferred.suffix else preferred.name
            for cand in _candidate_inputs(base_stem):
                if Path(cand).exists():
                    input_file = cand
                    break
            else:
                tried = ", ".join(_candidate_inputs(base_stem))
                raise FileNotFoundError(
                    f"Input file not found. Looked for: {tried}. "
                    "Pass a path explicitly: python transform_betas.py <file>"
                )

    # 1) Load data
    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    output_stem = _build_output_stem(input_file)

    # Print a context-aware message
    ext = input_path.suffix.lower()
    if ext in {".xlsx", ".xls", ".xlsm"}:
        print(f"Reading input: {input_path.name} (sheet='{INPUT_SHEET}')")
        print(f"Output will be: {output_stem}.xlsx (two sheets)")
    else:
        print(f"Reading input: {input_path.name}")
        print(f"Output will be: {output_stem}_StandardizedBetas.csv and {output_stem}_TransformedBetas.csv")

    df = _read_input(input_file, INPUT_SHEET)

    meta = df.iloc[:, :2].astype(str)            # ticker & name
    betas_raw = df.iloc[:, 2:]

    # Clean numeric text for CSV/TXT inputs: remove thousands separators, NBSPs, and unicode minus
    if ext in {".csv", ".txt"}:
        betas_raw = betas_raw.replace({
            "\u00A0": "",
            ",": "",
            "\u2212": "-",
        }, regex=True)

    # Convert to numeric
    betas = betas_raw.apply(pd.to_numeric, errors="coerce")

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

    # 5) Write out results
    if ext in {".xlsx", ".xls", ".xlsm"}:
        with pd.ExcelWriter(f"{output_stem}.xlsx", engine="openpyxl") as w:
            pd.concat([meta, z], axis=1).to_excel(
                w, sheet_name="StandardizedBetas", index=False
            )
            pd.concat([meta, transformed], axis=1).to_excel(
                w, sheet_name="TransformedBetas", index=False
            )
    else:
        pd.concat([meta, z], axis=1).to_csv(
            f"{output_stem}_StandardizedBetas.csv", index=False
        )
        pd.concat([meta, transformed], axis=1).to_csv(
            f"{output_stem}_TransformedBetas.csv", index=False
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
    
    if ext in {".xlsx", ".xls", ".xlsm"}:
        print(f"Processing complete. Output file: {output_stem}.xlsx")
    else:
        print(
            "Processing complete. Output files: "
            f"{output_stem}_StandardizedBetas.csv, {output_stem}_TransformedBetas.csv"
        )


if __name__ == "__main__":
    main()
