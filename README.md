# Transform Betas Script

## 📂 Repository Structure

```
transform_betas/
├── transform_betas.py # Main Python script for data processing
├── README.md          # Project documentation (this file)
└── raw_betas.xlsx     # Input Excel file (user-provided)
```

- `transform_betas.py`: Executes the full transformation pipeline.
- `raw_betas.xlsx`: Must be provided by the user in the correct format.
- `README.md`: Contains all setup instructions, usage details, and output explanations.

## ✅ Features

- **Excel-Compatible Standardization**:
  Uses the formula `(x - mean) / STDEV.S`, matching Excel’s `STDEV.S` behavior exactly by using sample standard deviation (`ddof=1` in NumPy).
- **Replicates `PERCENTRANK.EXC`**:
  Calculates percentiles using the logic:
  `ROUND(rank / (n + 1) - 0.0005, 3)` to mirror Excel’s `PERCENTRANK.EXC` exactly.
- **Dual Output Excel Files**:
  Saves both:
  - `standardized_betas.xlsx` — standardized scores
  - `transformed_betas2.xlsx` — percentile-transformed values
- **Progress Feedback**:
  Uses `tqdm` to show real-time progress as each column is processed.
- **Lightweight & Efficient**:
  Minimal dependencies and optimized logic for fast performance on large spreadsheets.

## 🛠 Requirements

This project requires:
- **Python 3.7 or higher**
- **pip** (Python's package installer)

### Python Libraries

Required packages:
- `pandas` – for reading/writing Excel files and DataFrame operations.
- `numpy` – for efficient numerical computation.
- `tqdm` – for displaying progress bars during processing.
- `openpyxl` – enables pandas to write `.xlsx` files.

---

### 🔒 (Optional but Recommended) Use a Virtual Environment

A virtual environment keeps dependencies isolated to this project.

#### Create and Activate a Virtual Environment

**macOS/Linux:**
```bash
python3 -m venv venv
source venv/bin/activate
```

**Windows:**
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```
You’ll see your shell prompt change to show the virtual environment is active (e.g., `(venv)`).

#### 📦 Install Required Packages

Once your virtual environment is active, install the necessary libraries:
```bash
pip install pandas numpy tqdm openpyxl
```

Optionally, generate a `requirements.txt` for future reuse:
```bash
pip freeze > requirements.txt
```

To install from `requirements.txt`:
```bash
pip install -r requirements.txt
```

## 🚀 Setup Instructions

Follow these steps to get up and running:

1.  **Clone the Repository (or Download Files)**
    -   If you're using Git:
        ```bash
        git clone https://github.com/cadbia/transform_betas.git
        cd /transform_betas
        ```
    -   Or simply download `transform_betas.py` and `README.md` manually and place them in a new folder.

2.  **Add Your Input File** <a name="-input-file"></a>
    Place your `raw_betas.xlsx` file in the same directory as `transform_betas.py`.

    The script expects:
    -   **File name:** `raw_betas.xlsx`
    -   **Sheet name:** `Sheet1`

    The Excel file must follow this format:

    | ID  | Label | Beta1 | Beta2 | Beta3 | ... |
    |-----|-------|-------|-------|-------|-----|
    | 001 | A     | 0.42  | -0.16 | 0.78  | ... |
    | 002 | B     | 0.35  | 0.10  | 0.65  | ... |

    -   **Columns A–B:** metadata (e.g. IDs, names, dates)
    -   **Columns C onward:** numeric beta values

    > ⚠️ **Important:** Ensure the Excel file is **not open** when running the script, and confirm that all beta columns contain **numeric values only**.

## ▶️ Usage

Once everything is set up and your `raw_betas.xlsx` file is in place, run the script from your terminal or command prompt:

```bash
python transform_betas.py
```

## What it does
1. Reads an Excel file of raw betas (default: `raw_betas.xlsx`, sheet `Sheet1`).
2. Treats the first two columns as metadata (e.g., Ticker, Name); remaining columns are numeric betas.
3. Standardizes each beta column using Excel-style z-scores: (x - mean) / STDEV.S (ddof=1).
4. Computes Excel-like PERCENTRANK.EXC percentiles across the global distribution of all standardized values (interpolated, with proper min/max handling).
5. Applies linear transform: (percentile*100 - 50.5) / 34.0.
6. Writes one output workbook with two sheets:
   - StandardizedBetas
   - TransformedBetas
7. Performs a validation report (counts blank cells).

## File naming
Output file name: `{OUTPUT_PREFIX}_{DATE}.xlsx`  
`OUTPUT_PREFIX` is defined near the top of `transform_betas.py` (default: `transformed_factor_betas`).  
`DATE` is auto-extracted from the input filename if it contains one of:
- `YYYYMMDD` (e.g. `raw_betas_20250707.xlsx`)
- `YYYY-MM-DD` or `YYYY_MM_DD`
- `MM-DD-YYYY` or `MM_DD_YYYY`

If no date pattern is found, the script uses the input file's modified date; if that fails, today’s date.

## Quick start
1. Place `raw_betas.xlsx` in the project root (or adjust `INPUT_FILE`).
2. Ensure first two columns are metadata; all remaining columns must be numeric (NaN allowed).
3. Install dependencies:
   ```bash
   pip install pandas numpy openpyxl
   ```
4. Run:
   ```bash
   python transform_betas.py
   ```
5. Check created file: `transformed_factor_betas_YYYYMMDD.xlsx`.

## Adjusting inputs
Edit these constants at the top of `transform_betas.py`:
```python
INPUT_FILE = "raw_betas.xlsx"
INPUT_SHEET = "Sheet1"
OUTPUT_PREFIX = "transformed_factor_betas"
```

## Validation output
Example console summary:
```
Reading input: raw_betas_2025-07-07.xlsx (sheet='Sheet1')
Output will be: transformed_factor_betas_20250707.xlsx
Data validation:
Original betas shape: (N, M)
Standardized betas - blank cells: 0
Transformed betas - blank cells: 0
✓ All transformed beta cells are filled
Processing complete. Output file: transformed_factor_betas_20250707.xlsx
```
If blanks appear in transformed data, their row indexes are listed.

## Common reasons for blanks
- Original column has all identical values (std = 0) → results become NaN.
- Original cell was NaN (propagates through).

## Troubleshooting
- Wrong sheet name: update `INPUT_SHEET`.
- No output file: verify script executed without exceptions.
- Date tag not what you expect: ensure filename contains a supported date format.

## License / reuse
Internal use; adapt constants as needed.