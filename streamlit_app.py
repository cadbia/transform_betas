import streamlit as st
import pandas as pd
import numpy as np
import re
from pathlib import Path
from datetime import datetime
import io

st.set_page_config(page_title="Transform Betas", layout="wide")

# --- Utility functions (adapted from transform_betas.py) ---

def _extract_date_tag_from_filename(filename: str) -> str:
    stem = Path(filename).stem
    # Try YYYYMMDD or similar patterns, return with underscores YYYY_MM_DD
    m = re.search(r"(20\d{2}[01]\d[0-3]\d)", stem)
    if m:
        raw = m.group(1)
        try:
            return datetime.strptime(raw, "%Y%m%d").strftime("%Y_%m_%d")
        except ValueError:
            pass
    m = re.search(r"(20\d{2})[-_](\d{2})[-_](\d{2})", stem)
    if m:
        y, mo, d = m.groups()
        try:
            return datetime(int(y), int(mo), int(d)).strftime("%Y_%m_%d")
        except ValueError:
            pass
    m = re.search(r"(\d{2})[-_](\d{2})[-_](20\d{2})", stem)
    if m:
        mo, d, y = m.groups()
        try:
            return datetime(int(y), int(mo), int(d)).strftime("%Y_%m_%d")
        except ValueError:
            pass
    return datetime.today().strftime("%Y_%m_%d")


def excel_percentrank_exc(sorted_arr, x):
    """
    Replicates Excel's PERCENTRANK.EXC behavior.
    """
    n = len(sorted_arr)
    if n < 2:
        return np.nan
    if x < sorted_arr[0] or x > sorted_arr[-1]:
        return np.nan
    if x == sorted_arr[0]:
        return 1 / (n + 1)
    if x == sorted_arr[-1]:
        return n / (n + 1)
    idx = np.searchsorted(sorted_arr, x, side="right")
    k = idx - 1
    xk, xk1 = sorted_arr[k], sorted_arr[k+1]
    fraction = (x - xk) / (xk1 - xk)
    return (k + fraction) / (n + 1)


@st.cache_data
def process_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Given a raw dataframe where first two columns are meta and the rest are numeric betas,
    returns (standardized_df, transformed_df) where both include the meta cols.
    """
    meta = df.iloc[:, :2].astype(str)
    betas_raw = df.iloc[:, 2:]

    # Clean numeric text for CSV/TXT inputs
    betas_raw = betas_raw.replace({"\u00A0": "", ",": "", "\u2212": "-"}, regex=True)
    betas = betas_raw.apply(pd.to_numeric, errors="coerce")

    # Standardize each column (ddof=1)
    z = betas.copy()
    for col in z.columns:
        vals = betas[col].values
        mu = np.nanmean(vals)
        sigma = np.nanstd(vals, ddof=1)
        if sigma == 0 or np.isnan(sigma):
            z[col] = np.nan
        else:
            z[col] = (vals - mu) / sigma

    # Flatten & sort all z-scores
    flat = z.values.flatten()
    flat = flat[~np.isnan(flat)]
    if len(flat) == 0:
        sorted_flat = np.array([])
    else:
        sorted_flat = np.sort(flat)

    # Compute EXCLUSIVE percentiles and transform
    def transform_exc(val):
        if np.isnan(val) or len(sorted_flat) < 2:
            return np.nan
        p_exc = excel_percentrank_exc(sorted_flat, val) * 100.0
        return (p_exc - 50.5) / 34.0

    transformed = z.applymap(transform_exc)

    standardized_out = pd.concat([meta, z], axis=1)
    transformed_out = pd.concat([meta, transformed], axis=1)
    return standardized_out, transformed_out


def make_excel_bytes(standardized: pd.DataFrame, transformed: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as w:
        # Only include the transformed values in the Excel output (single sheet)
        transformed.to_excel(w, sheet_name="TransformedBetas", index=False)
    return output.getvalue()


def make_csv_bytes(df: pd.DataFrame) -> bytes:
    """Return CSV file bytes for a DataFrame."""
    return df.to_csv(index=False).encode("utf-8")


# --- Streamlit UI ---
st.title("Transform Factor Betas")

uploaded = st.file_uploader("Upload a CSV or Excel file where Column 1 = Symbol, Column 2 = Company Name, Columns 3-90 = numeric beta values.", type=["csv", "txt", "xlsx", "xls", "xlsm"], accept_multiple_files=False)

if uploaded is not None:
    filename = uploaded.name
    ext = Path(filename).suffix.lower()

    try:
        if ext in {".csv", ".txt"}:
            # Read CSV; try to decode as utf-8, fallback to latin-1
            try:
                content = uploaded.getvalue().decode("utf-8")
            except UnicodeDecodeError:
                try:
                    content = uploaded.getvalue().decode("latin-1")
                    st.warning("File decoded using latin-1 encoding; verify characters are correct.")
                except Exception as e:
                    raise ValueError("Unable to decode CSV file; please ensure it's UTF-8 or latin-1 encoded.")
            df = pd.read_csv(io.StringIO(content), low_memory=False)
        else:
            # Excel
            sheet = st.text_input("Excel sheet name", value="Sheet1")
            df = pd.read_excel(uploaded, sheet_name=sheet)

        # Early validation: need at least 3 columns (symbol, name, at least one beta)
        if df.shape[1] < 3:
            st.error("Uploaded file must have at least three columns: Symbol, Company Name, and one or more beta columns.")
            st.stop()

        st.subheader("Preview of input data")
        st.dataframe(df.head(10))

        if st.button("Process file"):
            with st.spinner("Processing..."):
                standardized, transformed = process_dataframe(df)

            st.success("Processing complete")
            # Standardized betas will not be displayed or offered as a separate download
            st.subheader("Transformed Betas")
            st.dataframe(transformed.head(20))

            date_tag = _extract_date_tag_from_filename(filename)
            base_stem = f"transformed_factor_betas_{date_tag}"

            # Offer downloads: CSV inputs -> one CSV; Excel inputs -> one-sheet Excel
            if ext in {".csv", ".txt"}:
                # Input was CSV/TXT -> provide a single CSV with transformed betas
                csv_tr = make_csv_bytes(transformed)
                st.download_button(
                    label="Download Transformed Betas (CSV)",
                    data=csv_tr,
                    file_name=f"{base_stem}_TransformedBetas.csv",
                    mime="text/csv",
                )
            else:
                # Input was Excel -> provide a single-sheet Excel with transformed betas
                excel_bytes = make_excel_bytes(standardized, transformed)
                st.download_button(
                    label="Download Transformed Betas (Excel)",
                    data=excel_bytes,
                    file_name=f"{base_stem}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            # Validation summary
            st.write("Data validation:")
            st.write(f"Original betas shape: { (df.shape[0], df.shape[1]-2) }")
            st.write(f"Transformed - blank cells: { transformed.isna().sum().sum() }")

    except Exception as e:
        st.error(f"Error reading or processing file: {e}")
else:
    st.info("Waiting for file upload — drag and drop a CSV or Excel file.")

st.markdown("---")
