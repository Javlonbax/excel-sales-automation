# Excel Sales Automation

A Python-based automation project that combines multiple Excel sales files, cleans the data, generates summary reports, and automatically creates an Excel dashboard.

## Dashboard Preview

![Sales Dashboard](images/dashboard.png)

## Features

- Automatically reads multiple Excel files
- Combines monthly sales data into one dataset
- Cleans and standardizes input data
- Calculates sales automatically
- Generates sales summaries by region
- Generates sales summaries by product
- Generates monthly sales reports
- Creates KPI metrics
- Automatically builds an Excel dashboard
- Reduces repetitive manual Excel reporting

## Dashboard Metrics

The generated dashboard includes:

- Total Sales
- Total Orders
- Total Quantity
- Monthly Sales Trend
- Sales by Region
- Top Products

## Technologies

- Python
- Pandas
- OpenPyXL
- Microsoft Excel
- Git / GitHub

## Project Structure

```text
excel-sales-automation/
│
├── data/
│   ├── input/
│   └── output/
│
├── images/
│   └── dashboard.png
│
├── src/
│   ├── create_data.py
│   ├── generate_report.py
│   └── main.py
│
├── .gitignore
├── README.md
└── requirements.txt
```

## Installation

Install the required Python packages:

```bash
pip install -r requirements.txt
```

## Usage

Place the source Excel files inside:

```text
data/input/
```

Then run:

```bash
python src/main.py
```

The application automatically processes the files and generates:

```text
data/output/Sales_Report.xlsx
```

## Business Use Cases

This solution can be adapted for:

- Sales reporting
- Monthly business reports
- Inventory reporting
- Financial data processing
- Excel report automation
- Data cleaning and consolidation
- Repetitive Excel workflows

## Goal

The goal of this project is to demonstrate how Python can automate repetitive Excel reporting tasks and turn multiple raw files into a structured business report and dashboard.