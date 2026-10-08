
# ============================================================
# DISCOUNTED CASH FLOW VALUATION LAB
# Financial Intelligence Dashboard
# ============================================================

import io
import math

import numpy as np
import pandas as pd
import streamlit as st
import yfinance as yf


# ============================================================
# DATA HELPERS
# ============================================================

def clean_number(value):
    try:
        result = float(value)
        return result if math.isfinite(result) else None
    except (TypeError, ValueError):
        return None


def statement_value(statement, names, period=None):
    if statement is None or statement.empty:
        return None

    if period is None:
        period = sorted(statement.columns, reverse=True)[0]

    if period not in statement.columns:
        return None

    for name in names:
        if name in statement.index:
            value = statement.loc[name, period]

            if isinstance(value, pd.Series):
                value = value.iloc[0]

            return clean_number(value)

    return None


@st.cache_data(ttl=3600, show_spinner=False)
def load_dcf_data(ticker):
    stock = yf.Ticker(ticker)
    info = stock.info or {}

    income = stock.financials
    balance = stock.balance_sheet

    if income.empty or balance.empty:
        raise ValueError(
            "Annual income statement or balance sheet unavailable."
        )

    common_periods = sorted(
        set(income.columns) & set(balance.columns),
        reverse=True
    )

    if not common_periods:
        raise ValueError(
            "No matching annual income statement and "
            "balance sheet dates were found."
        )

    period = common_periods[0]

    revenue = statement_value(
        income, ["Total Revenue"], period
    )

    ebit = statement_value(
        income, ["EBIT", "Operating Income"], period
    )

    cash = statement_value(
        balance,
        [
            "Cash And Cash Equivalents",
            "Cash Cash Equivalents And Short Term Investments"
        ],
        period
    )

    debt = statement_value(
        balance,
        [
            "Total Debt",
            "Long Term Debt And Capital Lease Obligation"
        ],
        period
    )

    shares = clean_number(
        info.get("sharesOutstanding")
    )

    market_price = clean_number(
        info.get("currentPrice")
        or info.get("regularMarketPrice")
    )

    if revenue is None or revenue <= 0:
        raise ValueError("Valid annual revenue unavailable.")

    if ebit is None:
        raise ValueError("Annual EBIT unavailable.")

    if shares is None or shares <= 0:
        raise ValueError("Shares outstanding unavailable.")

    return {
        "name": info.get("longName", ticker),
        "sector": info.get("sector", "N/A"),
        "industry": info.get("industry", "N/A"),
        "period": period.strftime("%Y-%m-%d"),
        "revenue": revenue,
        "ebit": ebit,
        "cash": cash,
        "debt": debt,
        "shares": shares,
        "price": market_price
    }


# ============================================================
# DCF CALCULATION ENGINE
# ============================================================

def calculate_dcf(
    revenue,
    ebit_margin,
    growth_rates,
    tax_rate,
    da_ratio,
    capex_ratio,
    nwc_ratio,
    wacc,
    terminal_growth,
    cash,
    debt,
    shares
):

    if wacc <= terminal_growth:
        raise ValueError(
            "WACC must be greater than terminal growth."
        )

    if shares <= 0:
        raise ValueError(
            "Shares outstanding must be positive."
        )

    forecasts = []
    previous_revenue = revenue

    for year, growth in enumerate(growth_rates, start=1):

        projected_revenue = previous_revenue * (1 + growth)

        ebit = projected_revenue * ebit_margin

        nopat = ebit * (1 - tax_rate)

        depreciation = projected_revenue * da_ratio

        capex = projected_revenue * capex_ratio

        change_nwc = (
            projected_revenue - previous_revenue
        ) * nwc_ratio

        fcff = (
            nopat
            + depreciation
            - capex
            - change_nwc
        )

        discount_factor = (1 + wacc) ** year

        pv_fcff = fcff / discount_factor

        forecasts.append({
            "Year": year,
            "Revenue": projected_revenue,
            "Revenue Growth": growth,
            "EBIT": ebit,
            "NOPAT": nopat,
            "D&A": depreciation,
            "CapEx": capex,
            "Change in NWC": change_nwc,
            "Unlevered FCF": fcff,
            "PV of FCF": pv_fcff
        })

        previous_revenue = projected_revenue

    forecast_df = pd.DataFrame(forecasts)

    final_fcff = forecast_df.iloc[-1]["Unlevered FCF"]

    terminal_value = (
        final_fcff * (1 + terminal_growth)
        / (wacc - terminal_growth)
    )

    pv_terminal_value = (
        terminal_value / (1 + wacc) ** len(growth_rates)
    )

    pv_forecast = forecast_df["PV of FCF"].sum()

    enterprise_value = pv_forecast + pv_terminal_value

    equity_value = enterprise_value + cash - debt

    implied_share_price = equity_value / shares

    return {
        "forecast": forecast_df,
        "enterprise_value": enterprise_value,
        "equity_value": equity_value,
        "implied_price": implied_share_price,
        "pv_terminal_value": pv_terminal_value,
        "pv_forecast": pv_forecast,
        "terminal_weight": (
            pv_terminal_value / enterprise_value * 100
            if enterprise_value != 0 else None
        )
    }


# ============================================================
# EXCEL EXPORT
# ============================================================

def create_dcf_excel(data, assumptions, result, sensitivity):

    output = io.BytesIO()

    summary = pd.DataFrame([
        ["Company", data["name"]],
        ["Fiscal Year", data["period"]],
        ["Enterprise Value", result["enterprise_value"]],
        ["Equity Value", result["equity_value"]],
        ["Implied Share Price", result["implied_price"]],
        ["Market Price", data["price"]],
        ["PV of Forecast FCF", result["pv_forecast"]],
        ["PV of Terminal Value", result["pv_terminal_value"]],
    ], columns=["Metric", "Value"])

    assumptions_df = pd.DataFrame(
        list(assumptions.items()),
        columns=["Assumption", "Value"]
    )

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        summary.to_excel(
            writer, sheet_name="Valuation Summary", index=False
        )

        assumptions_df.to_excel(
            writer, sheet_name="Assumptions", index=False
        )

        result["forecast"].to_excel(
            writer, sheet_name="Five Year Forecast", index=False
        )

        sensitivity.to_excel(
            writer, sheet_name="Sensitivity Analysis"
        )

        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "B2"
            sheet.column_dimensions["A"].width = 28

            for column in "BCDEFGHIJK":
                sheet.column_dimensions[column].width = 20

            for cell in sheet[1]:
                cell.font = __import__(
                    "openpyxl"
                ).styles.Font(
                    bold=True, color="FFFFFF"
                )
                cell.fill = __import__(
                    "openpyxl"
                ).styles.PatternFill(
                    "solid", fgColor="17365D"
                )

    return output.getvalue()


# ============================================================
# STREAMLIT DCF PAGE
# ============================================================

def show_dcf_page():

    st.header("DCF Valuation Lab")

    st.write(
        "Estimate a company's intrinsic equity value using "
        "a five-year unlevered free cash flow forecast "
        "and the perpetuity growth method."
    )

    ticker = st.text_input(
        "Company ticker",
        value="AAPL",
        key="dcf_ticker"
    ).upper().strip()

    if not ticker:
        st.info("Enter a stock ticker to begin.")
        return

    try:
        with st.spinner("Loading annual financial statements..."):
            data = load_dcf_data(ticker)

    except Exception as error:
        st.error(f"Unable to load company data: {error}")
        return

    st.subheader(data["name"])

    st.caption(
        f"Financial statement date: {data['period']} "
        f"| Sector: {data['sector']}"
    )

    if data["sector"] == "Financial Services":
        st.warning(
            "This operating-company FCFF model may not be "
            "appropriate for banks, insurers, and certain "
            "other financial institutions. Use an "
            "industry-specific valuation methodology."
        )
        return

    if data["cash"] is None or data["debt"] is None:
        st.error(
            "Cash or total debt is unavailable for the matched "
            "fiscal period. The enterprise-to-equity bridge "
            "cannot be calculated reliably."
        )
        return

    historical_margin = (
        data["ebit"] / data["revenue"] * 100
    )

    st.divider()
    st.subheader("Valuation Assumptions")

    st.caption(
        "These are analyst-controlled scenario inputs, "
        "not independently estimated forecasts."
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        revenue_growth = st.slider(
            "Annual revenue growth (%)",
            min_value=-20.0,
            max_value=35.0,
            value=7.0,
            step=0.5
        )

        operating_margin = st.slider(
            "Forecast EBIT margin (%)",
            min_value=-20.0,
            max_value=60.0,
            value=float(
                max(-20, min(60, round(historical_margin, 1)))
            ),
            step=0.5
        )

        tax_rate = st.slider(
            "Effective cash tax assumption (%)",
            0.0, 40.0, 21.0, 0.5
        )

    with col2:
        da_ratio = st.slider(
            "D&A / Revenue (%)",
            0.0, 20.0, 4.0, 0.5
        )

        capex_ratio = st.slider(
            "CapEx / Revenue (%)",
            0.0, 30.0, 5.0, 0.5
        )

        nwc_ratio = st.slider(
            "Incremental NWC / Revenue change (%)",
            -20.0, 30.0, 2.0, 0.5
        )

    with col3:
        wacc = st.slider(
            "WACC (%)",
            4.0, 20.0, 9.0, 0.25
        )

        terminal_growth = st.slider(
            "Terminal growth (%)",
            0.0, 5.0, 2.5, 0.25
        )

    st.caption(
        "WACC, tax rate, D&A, CapEx, and working-capital "
        "assumptions should ultimately be supported by "
        "historical filings and a documented forecast thesis."
    )

    growth_rates = [revenue_growth / 100] * 5

    assumptions = {
        "Annual Revenue Growth (%)": revenue_growth,
        "EBIT Margin (%)": operating_margin,
        "Tax Rate (%)": tax_rate,
        "D&A / Revenue (%)": da_ratio,
        "CapEx / Revenue (%)": capex_ratio,
        "Incremental NWC Ratio (%)": nwc_ratio,
        "WACC (%)": wacc,
        "Terminal Growth (%)": terminal_growth,
    }

    inputs = dict(
        revenue=data["revenue"],
        ebit_margin=operating_margin / 100,
        growth_rates=growth_rates,
        tax_rate=tax_rate / 100,
        da_ratio=da_ratio / 100,
        capex_ratio=capex_ratio / 100,
        nwc_ratio=nwc_ratio / 100,
        wacc=wacc / 100,
        terminal_growth=terminal_growth / 100,
        cash=data["cash"],
        debt=data["debt"],
        shares=data["shares"],
    )

    try:
        result = calculate_dcf(**inputs)

    except ValueError as error:
        st.error(str(error))
        return

    st.divider()
    st.subheader("DCF Valuation Results")

    implied_price = result["implied_price"]
    market_price = data["price"]

    upside = (
        (implied_price / market_price - 1) * 100
        if market_price is not None and market_price > 0
        else None
    )

    cols = st.columns(4)

    cols[0].metric(
        "Implied Share Value",
        f"${implied_price:,.2f}"
    )

    cols[1].metric(
        "Market Price",
        f"${market_price:,.2f}"
        if market_price is not None else "N/A"
    )

    cols[2].metric(
        "Implied Upside / Downside",
        f"{upside:+.2f}%"
        if upside is not None else "N/A"
    )

    cols[3].metric(
        "Enterprise Value",
        f"${result['enterprise_value'] / 1e9:,.2f}B"
    )

    st.caption(
        "The implied value is a scenario estimate, not "
        "a verified fair value or buy/sell recommendation. "
        "The current market price may have a different "
        "as-of date from the annual financial statements."
    )

    st.divider()
    st.subheader("Five-Year Financial Forecast")

    forecast = result["forecast"].copy()

    chart_data = forecast.set_index("Year")[
        ["Revenue", "EBIT", "Unlevered FCF"]
    ] / 1e9

    st.line_chart(
        chart_data,
        x_label="Forecast Year",
        y_label="USD (Billions)"
    )

    with st.expander("View Full Forecast Model"):
        st.dataframe(
            forecast.style.format({
                col: "${:,.0f}"
                for col in forecast.columns
                if col not in ("Year", "Revenue Growth")
            }).format({
                "Revenue Growth": "{:.2%}"
            }),
            use_container_width=True,
            hide_index=True
        )

    st.divider()
    st.subheader("Valuation Bridge")

    bridge = pd.DataFrame({
        "Component": [
            "PV of Forecast Cash Flows",
            "PV of Terminal Value",
            "Enterprise Value",
            "Add: Cash",
            "Less: Debt",
            "Equity Value"
        ],
        "Value (USD Billions)": [
            result["pv_forecast"] / 1e9,
            result["pv_terminal_value"] / 1e9,
            result["enterprise_value"] / 1e9,
            data["cash"] / 1e9,
            -data["debt"] / 1e9,
            result["equity_value"] / 1e9
        ]
    })

    st.dataframe(
        bridge,
        use_container_width=True,
        hide_index=True
    )

    if (
        result["terminal_weight"] is not None
        and result["terminal_weight"] > 75
    ):
        st.warning(
            "More than 75% of enterprise value comes from "
            "terminal value. The valuation is highly "
            "dependent on long-term assumptions."
        )

    st.divider()
    st.subheader("Sensitivity Analysis")

    st.write(
        "How does the implied share price change "
        "when WACC and terminal growth change?"
    )

    wacc_values = [
        wacc + offset
        for offset in [-2, -1, 0, 1, 2]
    ]

    growth_values = [
        terminal_growth + offset
        for offset in [-1, -0.5, 0, 0.5, 1]
    ]

    sensitivity_rows = []

    for test_wacc in wacc_values:

        row = {}

        for test_growth in growth_values:

            label = f"{test_growth:.2f}%"

            if test_wacc <= test_growth:
                row[label] = None
                continue

            scenario_inputs = inputs.copy()

            scenario_inputs["wacc"] = test_wacc / 100
            scenario_inputs["terminal_growth"] = (
                test_growth / 100
            )

            scenario = calculate_dcf(
                **scenario_inputs
            )

            row[label] = scenario["implied_price"]

        sensitivity_rows.append(row)

    sensitivity = pd.DataFrame(
        sensitivity_rows,
        index=[f"{x:.2f}%" for x in wacc_values]
    )

    sensitivity.index.name = "WACC"

    st.dataframe(
        sensitivity.style.format("${:,.2f}").background_gradient(
            cmap="RdYlGn",
            axis=None
        ),
        use_container_width=True
    )

    st.caption(
        "Sensitivity values represent implied dollars per "
        "share. Blank cells indicate invalid WACC and "
        "terminal growth combinations."
    )

    st.divider()

    excel_bytes = create_dcf_excel(
        data,
        assumptions,
        result,
        sensitivity
    )

    st.download_button(
        "Download DCF Valuation Model",
        data=excel_bytes,
        file_name=f"{ticker}_DCF_valuation.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary"
    )
