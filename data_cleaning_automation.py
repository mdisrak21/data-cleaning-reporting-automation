"""
Data Cleaning & Reporting Automation

Dataset: UCI Online Retail II
Source: https://archive.ics.uci.edu/dataset/502/online+retail+ii

Before running:
1. Download online_retail_II.xlsx from the official UCI website.
2. Keep it in the same folder as this Python file.
3. Install required libraries:
   pip install pandas matplotlib openpyxl
4. Run:
   python data_cleaning_automation.py
"""

from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt


# --------------------------------------------------
# 1. Project paths
# --------------------------------------------------

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output"

# Create output folder automatically
OUT.mkdir(exist_ok=True)


# --------------------------------------------------
# 2. Load authentic dataset
# --------------------------------------------------

def load_data():

    file_path = ROOT / "online_retail_II.xlsx"

    if not file_path.exists():
        raise FileNotFoundError(
            "\nDataset file not found!\n"
            "Please download online_retail_II.xlsx from:\n"
            "https://archive.ics.uci.edu/dataset/502/online+retail+ii\n"
            "Then place it in the same folder as this Python file.\n"
        )

    print("Reading Excel dataset...")

    # Read all Excel sheets
    sheets = pd.read_excel(
        file_path,
        sheet_name=None
    )

    # Combine all sheets into one DataFrame
    df = pd.concat(
        sheets.values(),
        ignore_index=True
    )

    return df


# --------------------------------------------------
# 3. Main automation process
# --------------------------------------------------

def main():

    print("\nStarting Data Cleaning & Reporting Automation...\n")

    # Load original dataset
    raw = load_data()

    raw_rows = len(raw)

    print("Dataset loaded successfully!")
    print(f"Original rows: {raw_rows:,}")

    # --------------------------------------------------
    # 4. Standardize column names
    # --------------------------------------------------

    raw.columns = [
        str(column).strip()
        for column in raw.columns
    ]

    # Standardize known column names
    raw = raw.rename(
        columns={
            "InvoiceNo": "Invoice",
            "Customer ID": "CustomerID",
            "UnitPrice": "Price"
        }
    )

    # Check required columns
    required_columns = [
        "Invoice",
        "StockCode",
        "Description",
        "Quantity",
        "InvoiceDate",
        "Price",
        "CustomerID",
        "Country"
    ]

    missing_columns = [
        col for col in required_columns
        if col not in raw.columns
    ]

    if missing_columns:
        raise ValueError(
            f"Required columns are missing: {missing_columns}"
        )

    # --------------------------------------------------
    # 5. Check missing values and duplicates
    # --------------------------------------------------

    print("\nChecking data quality...")

    missing_before = raw.isna().sum()

    duplicate_count = int(
        raw.duplicated().sum()
    )

    print(f"Duplicate rows found: {duplicate_count:,}")

    # Remove exact duplicate rows
    df = raw.drop_duplicates().copy()

    # --------------------------------------------------
    # 6. Clean text columns
    # --------------------------------------------------

    text_columns = [
        "Invoice",
        "StockCode",
        "Description",
        "CustomerID",
        "Country"
    ]

    for column in text_columns:

        df[column] = (
            df[column]
            .astype("string")
            .str.strip()
        )

        df[column] = df[column].replace(
            "",
            pd.NA
        )

    # --------------------------------------------------
    # 7. Convert data types
    # --------------------------------------------------

    df["Quantity"] = pd.to_numeric(
        df["Quantity"],
        errors="coerce"
    )

    df["Price"] = pd.to_numeric(
        df["Price"],
        errors="coerce"
    )

    df["InvoiceDate"] = pd.to_datetime(
        df["InvoiceDate"],
        errors="coerce"
    )

    # --------------------------------------------------
    # 8. Handle missing essential values
    # --------------------------------------------------

    rows_before_missing_removal = len(df)

    essential_columns = [
        "Invoice",
        "StockCode",
        "InvoiceDate",
        "Quantity",
        "Price"
    ]

    df = df.dropna(
        subset=essential_columns
    ).copy()

    removed_missing_rows = (
        rows_before_missing_removal - len(df)
    )

    # Fill missing descriptive values
    df["Description"] = df["Description"].fillna(
        "Unknown product"
    )

    df["Country"] = df["Country"].fillna(
        "Unknown"
    )

    # CustomerID is retained when missing because
    # it is not essential for total sales calculation.

    # --------------------------------------------------
    # 9. Handle cancellation invoices
    # --------------------------------------------------

    cancellation_mask = (
        df["Invoice"]
        .astype("string")
        .str.upper()
        .str.startswith("C", na=False)
    )

    cancellation_count = int(
        cancellation_mask.sum()
    )

    print(
        f"Cancellation rows found: {cancellation_count:,}"
    )

    # --------------------------------------------------
    # 10. Handle non-positive quantity and price
    # --------------------------------------------------

    nonpositive_mask = (
        (df["Quantity"] <= 0) |
        (df["Price"] <= 0)
    )

    nonpositive_count = int(
        nonpositive_mask.sum()
    )

    # Remove cancellation and non-positive rows
    clean = df.loc[
        ~cancellation_mask & ~nonpositive_mask
    ].copy()

    # --------------------------------------------------
    # 11. Create revenue column
    # --------------------------------------------------

    clean["Revenue"] = (
        clean["Quantity"] * clean["Price"]
    )

    print("\nData cleaning completed!")
    print(f"Cleaned rows: {len(clean):,}")

    # --------------------------------------------------
    # 12. Generate data quality report
    # --------------------------------------------------

    missing_after = clean.isna().sum()

    quality_report = pd.concat(
        [
            missing_before.rename("Missing Before"),
            missing_after.rename("Missing After")
        ],
        axis=1
    ).fillna(0).astype(int)

    quality_report.index.name = "Column"

    quality_report.to_csv(
        OUT / "data_quality_summary.csv"
    )

    # --------------------------------------------------
    # 13. Generate cleaning summary
    # --------------------------------------------------

    cleaning_summary = pd.DataFrame({

        "Measure": [
            "Original rows",
            "Duplicate rows removed",
            "Rows removed due to missing essential values",
            "Cancellation rows excluded",
            "Rows with non-positive quantity or price excluded",
            "Final cleaned sales rows",
            "Missing CustomerID values retained"
        ],

        "Value": [
            raw_rows,
            duplicate_count,
            removed_missing_rows,
            cancellation_count,
            nonpositive_count,
            len(clean),
            int(clean["CustomerID"].isna().sum())
        ]

    })

    cleaning_summary.to_csv(
        OUT / "cleaning_summary.csv",
        index=False
    )

    # --------------------------------------------------
    # 14. Export cleaned dataset
    # --------------------------------------------------

    clean.to_csv(
        OUT / "cleaned_online_retail.csv",
        index=False
    )

    print("\nCleaned dataset exported.")

    # --------------------------------------------------
    # 15. Monthly revenue report
    # --------------------------------------------------

    monthly_revenue = (
        clean
        .set_index("InvoiceDate")
        .resample("MS")["Revenue"]
        .sum()
        .reset_index()
    )

    # --------------------------------------------------
    # 16. Country-wise revenue report
    # --------------------------------------------------

    country_revenue = (
        clean
        .groupby("Country")["Revenue"]
        .sum()
        .sort_values(ascending=False)
        .head(15)
        .reset_index()
    )

    # --------------------------------------------------
    # 17. Automated Excel report
    # --------------------------------------------------

    excel_path = OUT / "automated_report.xlsx"

    with pd.ExcelWriter(
        excel_path,
        engine="openpyxl"
    ) as writer:

        cleaning_summary.to_excel(
            writer,
            sheet_name="Cleaning Summary",
            index=False
        )

        quality_report.to_excel(
            writer,
            sheet_name="Missing Values"
        )

        monthly_revenue.to_excel(
            writer,
            sheet_name="Monthly Revenue",
            index=False
        )

        country_revenue.to_excel(
            writer,
            sheet_name="Top Countries",
            index=False
        )

    print("Automated Excel report created.")

    # --------------------------------------------------
    # 18. Create monthly revenue chart
    # --------------------------------------------------

    plt.figure(figsize=(10, 5))

    plt.plot(
        monthly_revenue["InvoiceDate"],
        monthly_revenue["Revenue"],
        marker="o"
    )

    plt.title(
        "Monthly Revenue - Cleaned Online Retail II"
    )

    plt.xlabel("Month")
    plt.ylabel("Revenue (£)")

    plt.xticks(rotation=45)
    plt.grid(True)
    plt.tight_layout()

    plt.savefig(
        OUT / "monthly_revenue.png",
        dpi=150
    )

    plt.close()

    # --------------------------------------------------
    # 19. Create top countries chart
    # --------------------------------------------------

    top_countries = (
        clean
        .groupby("Country")["Revenue"]
        .sum()
        .nlargest(10)
        .sort_values()
    )

    plt.figure(figsize=(9, 5))

    top_countries.plot(
        kind="barh"
    )

    plt.title(
        "Top 10 Countries by Revenue"
    )

    plt.xlabel("Revenue (£)")
    plt.ylabel("Country")

    plt.tight_layout()

    plt.savefig(
        OUT / "top_countries.png",
        dpi=150
    )

    plt.close()

    # --------------------------------------------------
    # 20. Final output
    # --------------------------------------------------

    print("\n===================================")
    print("PROJECT COMPLETED SUCCESSFULLY!")
    print("===================================")

    print(f"Original rows: {raw_rows:,}")
    print(f"Cleaned rows: {len(clean):,}")
    print(f"Duplicates removed: {duplicate_count:,}")
    print(f"Missing essential rows removed: {removed_missing_rows:,}")

    print("\nGenerated files are saved in the output folder:")

    print("1. cleaned_online_retail.csv")
    print("2. cleaning_summary.csv")
    print("3. data_quality_summary.csv")
    print("4. automated_report.xlsx")
    print("5. monthly_revenue.png")
    print("6. top_countries.png")

    print("\nOutput folder:")
    print(OUT)


# --------------------------------------------------
# Run project
# --------------------------------------------------

if __name__ == "__main__":
    main()