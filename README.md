# Data Cleaning & Reporting Automation

## Project Overview

This project automates the cleaning, validation, and reporting of the **Online Retail II** dataset. It uses Python and pandas to prepare transaction data, summarize data-quality issues, and generate reports and visualizations.

Place the source Excel workbook in the project folder, run the script, and the cleaned dataset and reports are generated in `output/`.

## Objectives

- Identify and remove duplicate transaction rows.
- Handle missing and inconsistent values.
- Exclude invalid sales records from the cleaned sales dataset.
- Create data-quality and cleaning summaries.
- Generate an Excel report and visual summaries automatically.

## Dataset

This project uses the **Online Retail II** dataset from the UCI Machine Learning Repository.

- **Source:** [UCI Machine Learning Repository — Online Retail II](https://archive.ics.uci.edu/dataset/502/online+retail+ii)
- **Input filename expected by the script:** `online_retail_II.xlsx`

Download the workbook from the UCI page and place it in the same folder as `data_cleaning_automation.py`. The script reads and combines the workbook's sheets.

## Project Structure

```text
task4-data-cleaning-reporting/
├── data_cleaning_automation.py
├── online_retail_II.xlsx          # Download separately from UCI
├── requirements.txt
├── README.md
└── output/
    ├── cleaned_online_retail.csv
    ├── cleaning_summary.csv
    ├── data_quality_summary.csv
    ├── automated_report.xlsx
    ├── monthly_revenue.png
    └── top_countries.png
```

> The source dataset is not included unless you add it yourself.

## Tools & Libraries

- **Python** — automation and data processing
- **pandas** — data cleaning and analysis
- **Matplotlib** — chart generation
- **openpyxl** — Excel workbook support

## Setup and Run

### 1. Install dependencies

Open a terminal in the project folder and run:

```bash
pip install -r requirements.txt
```

If needed, install directly:

```bash
pip install pandas matplotlib openpyxl
```

### 2. Add the dataset

Download the Online Retail II Excel workbook from the UCI page and save it as `online_retail_II.xlsx` beside `data_cleaning_automation.py`.

### 3. Run the automation

```bash
python data_cleaning_automation.py
```

After successful execution, check the `output/` folder.

## Data Cleaning Workflow

The script:

1. Reads all sheets from the Excel workbook and combines them.
2. Standardizes column names and prepares required fields.
3. Counts and removes duplicate rows.
4. Cleans text fields and converts quantity, price, and date fields to suitable data types.
5. Removes rows missing essential transaction fields.
6. Fills missing descriptions and country values with placeholders.
7. Retains missing `CustomerID` values and reports them rather than discarding those transactions.
8. Excludes cancellation invoices (invoice numbers beginning with `C`) and rows with non-positive quantity or price from the cleaned sales data.
9. Calculates `Revenue` as `Quantity × Price`.
10. Exports cleaned data, quality summaries, an Excel report, and charts.

## Generated Outputs

| File | Description |
|---|---|
| `cleaned_online_retail.csv` | Cleaned transaction-level sales data with calculated revenue |
| `cleaning_summary.csv` | Original rows, removed duplicates, excluded records, and final row counts |
| `data_quality_summary.csv` | Missing-value counts before and after cleaning |
| `automated_report.xlsx` | Excel report containing cleaning, missing-value, monthly revenue, and top-country summaries |
| `monthly_revenue.png` | Monthly revenue visualization |
| `top_countries.png` | Top-country visualization |

## Validation Results

The run checked during development produced these results:

| Measure | Result |
|---|---:|
| Original rows | 1,067,371 |
| Duplicate rows removed | 34,335 |
| Rows removed due to missing essential values | 0 |
| Rows with non-positive quantity or price excluded | 25,122 |
| Final cleaned sales rows | 1,007,913 |
| Missing `CustomerID` values retained | 228,488 |
| Duplicate rows in cleaned CSV | 0 |

The cleaned CSV check also showed no missing values in its exported columns. Missing `CustomerID` values are documented separately in the quality summary.

**Note:** These figures describe the run performed during development. Results may differ if the source workbook changes.

## Important Notes

- `CustomerID` is not treated as an essential field; rows with missing customer IDs are retained.
- Cancellation invoices and non-positive quantity/price rows are excluded from the cleaned sales output. Consider this rule when interpreting totals.
- A pandas `DtypeWarning` may appear for the invoice field because invoice identifiers can contain both numeric-looking and text values. Confirm that output files were generated successfully.

## Reproducibility

To reproduce the results:

1. Use the UCI Online Retail II workbook.
2. Keep the expected filename and folder location.
3. Install the dependencies.
4. Run the script.
5. Review the generated files in `output/`.

## Author

**Md. Israk Ahmmed**  


---
