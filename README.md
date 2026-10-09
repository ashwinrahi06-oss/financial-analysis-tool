# Financial Intelligence Dashboard

### Financial Analytics | Equity Research | DCF Valuation | Python Automation

**An interactive financial research platform combining automated company analysis, comparative financial analytics, and fundamental equity valuation using Python, Streamlit, Pandas, and Excel.**

[**Launch Live Dashboard**](https://financial-analysis-tool-tdvsxgr6kubmqmwnxkps9f.streamlit.app/) | [**View GitHub Repository**](https://github.com/ashwinrahi06-oss/financial-analysis-tool)

---

## 1. Project Overview

The **Financial Intelligence Dashboard** is an end-to-end financial analytics and equity research platform designed to streamline the process of evaluating publicly traded companies.

The application combines automated financial statement retrieval, financial ratio analysis, historical performance visualization, company comparisons, and downloadable Excel reports within a single interactive interface.

In addition to general-purpose company analysis, the platform includes a dedicated **Microsoft Equity Research & DCF Valuation** case study examining whether Microsoft's substantial investment in artificial intelligence infrastructure can generate sufficient long-term free cash flow to support its equity valuation.

The project demonstrates the integration of financial modeling, fundamental analysis, data engineering, and interactive application development.

### Core Capabilities

- Automated financial statement retrieval and analysis
- Historical revenue, profitability, and cash flow visualization
- Side-by-side public company comparisons
- Financial health and valuation indicators
- Five-year discounted cash flow modeling
- Bear, Base, and Bull valuation scenarios
- Weighted average cost of capital analysis
- Excel-based financial modeling and reporting
- SEC-sourced financial data integration for the Microsoft case study
- Interactive equity research dashboards

---

## 2. Application Modules

### Module 1: Company Financial Analysis

Analyze the financial performance of publicly traded companies by entering a ticker symbol.

The application retrieves available financial information and calculates key operating, profitability, cash flow, and market valuation metrics.

**Financial Performance**
- Total revenue
- Year-over-year revenue growth
- Operating income
- Net income
- Operating cash flow
- Capital expenditures
- Free cash flow

**Profitability**
- Operating margin
- Net profit margin
- Free cash flow margin

**Market Valuation**
- Market capitalization
- Price-to-earnings ratio
- Enterprise value to EBITDA

This module provides a structured overview of company performance without requiring users to manually download and organize financial statements.

### Module 2: Historical Financial Performance

Interactive charts allow users to evaluate changes in financial performance across available annual reporting periods.

**Operating Performance**
- Revenue trends
- Operating income trends
- Net income trends

**Cash Flow**
- Operating cash flow
- Free cash flow

**Profitability**
- Operating margin
- Net profit margin
- Free cash flow margin

The visualizations help identify historical growth patterns, margin expansion or compression, and changes in cash flow generation.

### Module 3: Company Comparison

Compare two publicly traded companies through a side-by-side financial analysis.

The comparison engine evaluates:

- Revenue and revenue growth
- Operating profitability
- Net profitability
- Cash flow generation
- Financial margins
- Market valuation metrics

The application also identifies differences in fiscal year-end dates and highlights limitations when comparing businesses across different industries.

This module is designed to support preliminary peer analysis, competitive benchmarking, and fundamental company research.

### Module 4: Financial Health Indicators

The platform provides simplified financial performance indicators covering:

- Revenue growth
- Profitability
- Free cash flow generation

These indicators summarize financial performance using available historical information.

They are intended as descriptive analytical tools rather than investment recommendations or predictive credit ratings.

### Module 5: Excel Report Generation

Users can export financial analysis results into structured Excel workbooks.

Depending on the selected analysis, reports may include:

- Company financial metrics
- Historical financial statements
- Financial ratios
- Company comparison tables
- Data quality information
- Financial methodology notes

This functionality supports further analysis, recordkeeping, and independent review outside the Streamlit application.

---

## 3. Featured Equity Research Case Study

# Microsoft | The AI Reinvestment Test

**A Five-Year DCF Valuation of Microsoft's AI Infrastructure Investment Strategy**

### Research Objective

Microsoft is investing heavily in artificial intelligence infrastructure, cloud computing capacity, and supporting data center operations.

While these investments may strengthen the company's long-term competitive position, they also require substantial capital expenditure.

The central research question is:

**Can Microsoft's future revenue growth, operating profitability, and cash flow generation justify the capital intensity associated with its AI infrastructure expansion?**

To investigate this question, I developed a five-year discounted cash flow valuation model integrating historical financial information, forward-looking operating assumptions, capital expenditure projections, and multiple investment scenarios.

### Valuation Framework

The analysis incorporates:

1. Historical financial data sourced from SEC filings
2. Revenue and operating income forecasts
3. Operating margin assumptions
4. Tax and net operating profit after tax calculations
5. Depreciation and amortization forecasts
6. Capital expenditure forecasts
7. Net working capital adjustments
8. Unlevered free cash flow projections
9. Weighted average cost of capital
10. Terminal value and enterprise-to-equity valuation
11. Bear, Base, and Bull scenario analysis
12. Interactive Streamlit valuation presentation

### Base-Case Valuation Results

| Metric | Base-Case Result |
|---|---:|
| Reference Share Price | $522.61 |
| DCF Implied Share Price | $261.09 |
| Implied Upside / Downside | -50.04% |
| Weighted Average Cost of Capital | 8.80% |
| FY2031 Revenue Growth | 10.9% |
| FY2031 EBIT Margin | 46.0% |
| FY2031 CapEx / Revenue | 25.2% |

*Results reflect the model's reference inputs and assumptions, not continuously updated market data.*

### Investment Thesis

The base-case valuation indicates a substantial difference between Microsoft's modeled intrinsic value and its reference market price.

Although the forecast assumes continued revenue growth and strong operating profitability, the valuation remains sensitive to the level of capital investment required to support Microsoft's AI and cloud infrastructure expansion.

In the FY2031 base case, revenue is projected to grow 10.9%, operating margins reach 46.0%, and capital expenditures represent 25.2% of revenue.

These assumptions illustrate a central valuation challenge: **strong accounting profitability does not necessarily translate into equally strong free cash flow when capital intensity remains elevated.**

The model estimates intrinsic value of approximately $261.09 per share compared with a reference market price of $522.61.

This represents approximately 50.04% implied downside under the base-case assumptions.

This result should not be interpreted as a forecast of Microsoft's future trading price. Instead, it illustrates the difference between the cash flows supported by the modeled assumptions and the valuation represented by the reference share price.

### Scenario Analysis

The model evaluates three distinct operating and valuation environments.

| Scenario | Description |
|---|---|
| Bear Case | Less favorable operating and cash flow assumptions |
| Base Case | Central forecast of operating performance and capital investment |
| Bull Case | More favorable growth, profitability, and cash flow assumptions |

Each scenario estimates future free cash flow and the resulting implied equity value.

Scenario analysis helps evaluate the sensitivity of intrinsic value to uncertainty in Microsoft's long-term financial performance.

### Why AI Capital Expenditure Matters

Capital expenditure is a particularly important variable in Microsoft's valuation.

Investment in data centers, servers, networking infrastructure, and other long-lived assets can support future revenue growth, but requires cash outflows before the associated economic benefits are fully realized.

The model therefore distinguishes between operating profitability and free cash flow generation.

This approach helps examine whether future operating performance can support sustained infrastructure reinvestment while still producing attractive cash flows for shareholders.

### Equity Research Deliverables

The Microsoft case study includes:

- Historical financial dataset
- Five-year financial forecasts
- Forecast assumption workbook
- Excel DCF valuation workbook
- Bear, Base, and Bull valuations
- WACC and terminal value calculations
- Scenario and sensitivity analysis
- SEC source audit information
- Interactive Streamlit equity research page

### Research Limitations

The DCF valuation is sensitive to assumptions regarding revenue growth, operating margins, capital expenditures, working capital, discount rates, and terminal value.

Historical inputs may include provisional figures or modeling adjustments, which should be distinguished from directly reported financial data.

The model is intended to evaluate the consequences of defined financial assumptions rather than predict future stock prices with certainty.

---

## 4. Financial Methodology

### General Financial Analysis

The general company analysis and comparison modules use annual financial statement data retrieved through the `yfinance` Python library.

**Revenue Growth**

Revenue Growth = (Current Year Revenue / Previous Year Revenue − 1) × 100

**Operating Margin**

Operating Margin = (Operating Income / Revenue) × 100

**Net Profit Margin**

Net Profit Margin = (Net Income / Revenue) × 100

**Free Cash Flow**

Free Cash Flow = Operating Cash Flow − Capital Expenditures

*When capital expenditures are reported as negative cash flows, the application adds the signed value to operating cash flow.*

**Free Cash Flow Margin**

Free Cash Flow Margin = (Free Cash Flow / Revenue) × 100

### Discounted Cash Flow Valuation

The Microsoft case study uses a free cash flow to the firm (FCFF) approach.

**Net Operating Profit After Tax**

NOPAT = EBIT × (1 − Tax Rate)

**Unlevered Free Cash Flow**

FCFF = NOPAT + Depreciation & Amortization − Capital Expenditures − Change in Net Working Capital

**Weighted Average Cost of Capital**

WACC = (E / (D + E)) × Cost of Equity + (D / (D + E)) × Cost of Debt × (1 − Tax Rate)

Where:

- E = Market value of equity
- D = Market value of debt

**Terminal Value**

Terminal Value = Final Forecast Year FCFF × (1 + Terminal Growth Rate) / (WACC − Terminal Growth Rate)

**Enterprise Value**

Enterprise Value = Present Value of Forecast FCFF + Present Value of Terminal Value

**Equity Value**

Equity Value = Enterprise Value − Debt + Cash and applicable non-operating assets

**Implied Share Price**

Implied Share Price = Equity Value / Shares Outstanding

The model uses scenario-specific financial forecasts to evaluate the range of potential intrinsic values.

---

## 5. Data Sources

### Yahoo Finance

The general-purpose company analysis modules use Yahoo Finance data accessed through `yfinance`.

Available information includes:

- Annual income statements
- Annual cash flow statements
- Financial performance metrics
- Market-based valuation information

### U.S. Securities and Exchange Commission

The Microsoft valuation case study incorporates historical financial information obtained from SEC filings and company-reported financial disclosures.

SEC financial information is used to support historical analysis and forward-looking valuation assumptions.

**Important distinction:** General company analytics and the Microsoft case study do not necessarily use the same financial data retrieval pipeline.

---

## 6. Technology Stack

| Technology | Application |
|---|---|
| Python | Financial calculations, modeling, and application logic |
| Streamlit | Interactive financial analytics and research dashboard |
| Pandas | Data cleaning, processing, and financial analysis |
| yfinance | Public company financial statement and market data |
| SEC Financial Data | Historical financial inputs for equity research |
| openpyxl | Excel financial modeling, workbook generation, and formatting |
| Microsoft Excel | Financial model review and independent analysis |
| GitHub | Version control and source code hosting |

### Technical Skills Demonstrated

**Financial Analysis**
- Financial statement analysis
- Financial ratio analysis
- Company comparison
- Cash flow analysis
- Discounted cash flow valuation
- Financial forecasting
- Scenario analysis
- WACC and terminal value estimation

**Programming and Data Analytics**
- Python application development
- Financial data extraction
- Data transformation and validation
- Automated financial calculations
- Interactive dashboard development
- Excel automation
- Modular code organization

---

## 7. Project Structure

```text
financial-analysis-tool/
├── app.py
├── company_comparison.py
├── financial_analysis_tool.py
├── requirements.txt
├── .gitignore
├── README.md
│
└── case_studies/
    ├── __init__.py
    │
    └── microsoft_valuation/
        ├── __init__.py
        ├── model_config.py
        ├── financial_data.py
        ├── forecast_assumptions.py
        ├── valuation_engine.py
        ├── build_case_study.py
        ├── equity_research_page.py
        │
        └── data/
            ├── Microsoft_Equity_Valuation.xlsx
            ├── Microsoft_Forecast_Assumptions.xlsx
            ├── Microsoft_Historical.csv
            ├── Microsoft_Forecast.csv
            └── valuation_snapshot.json
```

The application is organized into independent modules for general company analysis, comparative analysis, and dedicated equity research.

The Microsoft valuation module separates financial data retrieval, forecasting assumptions, valuation calculations, report generation, and dashboard presentation.

This structure supports future expansion into additional company research case studies.

---

## 8. Running the Project Locally

### Step 1: Clone the Repository

```bash
git clone https://github.com/ashwinrahi06-oss/financial-analysis-tool.git
cd financial-analysis-tool
```

### Step 2: Create a Virtual Environment

```bash
python3 -m venv .venv
```

Activate the environment on macOS or Linux:

```bash
source .venv/bin/activate
```

For Windows:

```powershell
.venv\Scripts\activate
```

### Step 3: Install Dependencies

```bash
python -m pip install -r requirements.txt
```

### Step 4: Launch the Dashboard

```bash
python -m streamlit run app.py
```

The application should open in your browser at:

`http://localhost:8501`

### Step 5: Access Microsoft Equity Research

Within the dashboard, navigate to:

**Microsoft Equity Research**

The page presents the Microsoft valuation case study and its financial modeling results.

### Optional: Rebuild the Microsoft Case Study

From the repository root, run:

```bash
python -m case_studies.microsoft_valuation.build_case_study
```

This runs the Microsoft case study generation module.

Rebuilding may require network access and appropriate source data availability. Outputs may change if source information, assumptions, or market inputs are updated.

---

## 9. Data Quality and Model Validation

Financial analysis depends on the reliability of both the source data and the assumptions used in financial calculations.

The platform incorporates data quality considerations across its analytical modules.

### General Company Analysis

- Financial data may contain missing or inconsistent values.
- Historical financial periods may vary by company.
- Market valuation metrics may reflect dates different from annual financial statements.
- Companies may have different fiscal year-end dates.
- Cross-sector comparisons may be economically inappropriate.
- Banking and other financial institutions may require specialized financial metrics.

### Microsoft Equity Research

- Historical inputs should be reconciled to company disclosures and SEC filings.
- Forecast assumptions represent modeled estimates rather than reported results.
- Depreciation and capital expenditures must be interpreted consistently.
- Enterprise-to-equity adjustments must reflect the treatment of cash, investments, and debt.
- Terminal value can represent a substantial proportion of enterprise value.
- Discount rate and terminal growth assumptions can materially affect implied share price.

The Microsoft valuation outputs have been cross-checked for consistency between the generated Excel workbook and the Streamlit dashboard.

This consistency check does not constitute an independent audit of the underlying financial data or economic assumptions.

---

## 10. Key Project Outcomes

This project demonstrates the development of a financial research platform that combines automated data analysis with traditional financial modeling.

**Automated Financial Analysis**

Reduced the need for manual financial statement collection and ratio calculations through reusable Python workflows.

**Interactive Financial Research**

Developed a Streamlit interface for exploring company fundamentals, historical trends, and comparative financial performance.

**Integrated Equity Valuation**

Built a Microsoft DCF case study with five-year forecasts, scenario analysis, and Excel-based financial modeling.

**Reproducible Analysis**

Organized financial data processing, modeling assumptions, valuation logic, and reporting into separate Python modules.

**Finance and Technology Integration**

Applied corporate finance concepts through practical programming, data analytics, and application development.

---

## 11. Future Development

Potential future enhancements include:

- Additional equity research case studies
- Reverse DCF analysis to estimate market-implied expectations
- Expanded peer-group benchmarking
- Industry-specific financial models
- Automated model validation and reconciliation tests
- Additional financial risk and return metrics
- More advanced scenario and sensitivity analysis
- Historical valuation comparisons

---

## 12. Disclaimer

This project was developed for educational, research, and portfolio demonstration purposes.

Financial information and model outputs may contain errors, estimates, or incomplete data. Valuation results depend on assumptions that may differ materially from actual future financial performance.

Nothing in this repository constitutes personalized investment advice, a recommendation to buy or sell securities, or a guarantee of future investment returns.

Users should independently verify financial information against primary company disclosures before relying on any analysis for investment decisions.

---

**Developed as an independent financial analytics and equity research project combining corporate finance, Python programming, financial modeling, and interactive data visualization.**
