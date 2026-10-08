# Financial Intelligence Dashboard

**An interactive financial analytics platform built with Python, Streamlit, Pandas, and Excel.**

[**Launch Live Dashboard**](https://financial-analysis-tool-tdvsxgr6kubmqmwnxkps9f.streamlit.app/)

## Project Overview

The Financial Intelligence Dashboard is a Python-based application designed to simplify the process of analyzing and comparing publicly traded companies.

It automates financial statement retrieval, calculates key financial performance metrics, visualizes historical trends, and generates downloadable Excel reports.

The project combines fundamental financial analysis with programming and data analytics to make company research more accessible, organized, and reproducible.

## Key Features

### 1. Company Financial Analysis

Enter a publicly traded company's ticker symbol to retrieve and analyze:

- Revenue and year-over-year revenue growth
- Operating income and operating margin
- Net income and net profit margin
- Operating cash flow and free cash flow
- Free cash flow margin
- Market capitalization, P/E ratio, and EV/EBITDA

### 2. Historical Financial Performance

Visualize historical annual performance through interactive charts covering:

- Revenue, operating income, and net income
- Operating cash flow and free cash flow
- Operating margin, net margin, and free cash flow margin

### 3. Company Comparison

Compare two publicly traded companies side by side.

The comparison engine evaluates financial performance, profitability, growth, cash flow, and valuation metrics. It also identifies differences in fiscal year-end dates and highlights potential limitations when comparing companies across different sectors.

### 4. Financial Health Indicators

Review simplified indicators for:

- Revenue growth
- Profitability
- Free cash flow generation

These indicators provide a descriptive snapshot of financial performance rather than an investment recommendation.

### 5. Excel Report Generation

Download Excel workbooks containing financial metrics, historical financial data, comparison tables, data-quality information, and methodology notes.

## Financial Methodology

The application uses annual financial statement data retrieved through the `yfinance` Python library.

**Revenue Growth**

(Current Year Revenue / Previous Year Revenue − 1) × 100

**Operating Margin**

(Operating Income / Revenue) × 100

**Net Profit Margin**

(Net Income / Revenue) × 100

**Free Cash Flow**

Operating Cash Flow − Capital Expenditures

*Implementation note: The application adds the signed capital expenditure value when the data source reports capital expenditures as negative.*

**Free Cash Flow Margin**

(Free Cash Flow / Revenue) × 100

## Technology Stack

| Technology | Purpose |
|---|---|
| Python | Core application logic and financial calculations |
| Streamlit | Interactive web dashboard |
| Pandas | Financial data processing and analysis |
| yfinance | Financial statement and market data retrieval |
| openpyxl | Excel workbook generation and formatting |
| GitHub | Version control and project hosting |

## Project Structure

```text
financial-analysis-tool/
├── app.py
├── company_comparison.py
├── financial_analysis_tool.py
├── requirements.txt
├── .gitignore
└── README.md
```

## Running the Project Locally

Clone the repository:

```bash
git clone https://github.com/ashwinrahi06-oss/financial-analysis-tool.git
cd financial-analysis-tool
```

Create and activate a Python virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Install the required dependencies:

```bash
python -m pip install -r requirements.txt
```

Launch the dashboard:

```bash
python -m streamlit run app.py
```

The application should open in your browser at `http://localhost:8501`.

## Data Quality and Limitations

- Financial data is obtained through Yahoo Finance via `yfinance` and may contain missing, delayed, or inconsistent values.
- Financial statement metrics use annual reporting periods, while market-based valuation metrics may reflect different dates.
- Companies can have different fiscal year-end dates, which may limit direct comparisons.
- Certain conventional operating-company metrics are not assessed for banking companies.
- Historical data availability varies by company.
- Automated analysis should be independently verified against company filings before use in investment decisions.

## Future Development

Potential enhancements include discounted cash flow valuation, expanded financial ratio analysis, industry-specific benchmarks, and more sophisticated peer-group comparisons.

## Disclaimer

This project was developed for educational and portfolio purposes. It is not intended to provide investment advice or replace professional financial research.
