from pathlib import Path
import pandas as pd
from generate_report import generate_dashboard


BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def load_data():
    files = list(INPUT_DIR.glob("*.xlsx"))

    if not files:
        raise FileNotFoundError(
            f"Excel files not found in: {INPUT_DIR}"
        )

    print(f"Found {len(files)} files")

    data = []

    for file in files:
        df = pd.read_excel(file)

        df.columns = (
            df.columns
            .str.strip()
            .str.lower()
            .str.replace(" ", "_")
        )

        df["month"] = file.stem

        data.append(df)

    return pd.concat(data, ignore_index=True)


def clean_data(df):
    df["quantity"] = pd.to_numeric(
        df["quantity"],
        errors="coerce"
    )

    df["price"] = pd.to_numeric(
        df["price"],
        errors="coerce"
    )

    df["sales"] = df["quantity"] * df["price"]

    return df


def create_reports(df):
    region_report = (
        df.groupby("region", as_index=False)
        .agg(
            sales=("sales", "sum"),
            quantity=("quantity", "sum"),
            orders=("order_id", "nunique")
        )
        .sort_values("sales", ascending=False)
    )

    product_report = (
        df.groupby(
            ["product", "category"],
            as_index=False
        )
        .agg(
            sales=("sales", "sum"),
            quantity=("quantity", "sum")
        )
        .sort_values("sales", ascending=False)
    )

    monthly_report = (
        df.groupby("month", as_index=False)
        .agg(
            sales=("sales", "sum"),
            quantity=("quantity", "sum"),
            orders=("order_id", "nunique")
        )
    )

    return region_report, product_report, monthly_report


def save_report(
    df,
    region_report,
    product_report,
    monthly_report
):
    
    output_file = OUTPUT_DIR / "Sales_Report.xlsx"

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Raw_Data",
            index=False
        )

        region_report.to_excel(
            writer,
            sheet_name="By_Region",
            index=False
        )

        product_report.to_excel(
            writer,
            sheet_name="By_Product",
            index=False
        )

        monthly_report.to_excel(
            writer,
            sheet_name="By_Month",
            index=False
        )

    return output_file


def main():
    df = load_data()

    df = clean_data(df)

    region_report, product_report, monthly_report = (
        create_reports(df)
    )

    output_file = save_report(
        df,
        region_report,
        product_report,
        monthly_report
    )
    generate_dashboard()

    total_sales = df["sales"].sum()
    total_quantity = df["quantity"].sum()
    orders = df["order_id"].nunique()

    print()
    print("========== SALES REPORT ==========")
    print(f"Total sales:    {total_sales:,.0f}")
    print(f"Quantity:       {total_quantity:,.0f}")
    print(f"Orders:         {orders:,}")
    print()
    print(f"Report created: {output_file}")


if __name__ == "__main__":
    main()