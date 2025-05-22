# Beta Transformation Script

A lightweight Python script that processes large Excel datasets of raw beta values, standardizes them using Excel’s `STDEV.S` formula, and transforms them into scaled percentiles that match Excel’s `PERCENTRANK.EXC` logic. Outputs are rounded to **nine decimal places** by default.

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
    -   **Sheet name:** `RawBetas`

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

### 🔄 What the Script Does

1.  Reads the `raw_betas.xlsx` Excel file from the `RawBetas` sheet.
2.  Standardizes each beta column using Excel’s `STDEV.S` logic:
    ```
    standardized = (x – mean) / STDEV.S
    ```
3.  Flattens all standardized values and sorts them for percentile ranking.
4.  Calculates percentiles using a replication of Excel’s `PERCENTRANK.EXC` function:
    ```
    percentile = ROUND(rank / (N + 1) – 0.0005, 9)
    ```
5.  Transforms the percentile to a scaled score using:
    ```
    transformed = (percentile * 100 – 50.5) / 34
    ```
6.  Outputs two Excel files:
    -   `standardized_betas.xlsx` — contains metadata + standardized Z-scores
    -   `transformed_betas2.xlsx` — contains metadata + percentile-transformed scores
7.  Prints total execution time in minutes and seconds to the console.

✅ After successful execution, both output Excel files will appear in the same directory as the script.