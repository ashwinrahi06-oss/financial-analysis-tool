
"""Interactive Microsoft equity research dashboard."""

import json

import pandas as pd
import plotly.express as px
import streamlit as st

from .model_config import DATA, SCENARIOS
from .valuation_engine import wacc, value_case


def render():
    st.title("Microsoft | The AI Reinvestment Test")

    st.caption(
        "SEC-grounded historical analysis | "
        "Five-year FCFF DCF | "
        "Interactive valuation scenarios"
    )

    snapshot_path = (
        DATA / "valuation_snapshot.json"
    )

    if not snapshot_path.exists():
        st.error(
            "Missing valuation snapshot. "
            "Run build_case_study.py first."
        )
        st.stop()

    snapshot = json.loads(
        snapshot_path.read_text()
    )

    history = pd.read_csv(
        DATA / "Microsoft_Historical.csv"
    )

    forecast = pd.read_csv(
        DATA / "Microsoft_Forecast.csv"
    )

    st.warning(
        "Accounting limitation: D&A begins with a "
        "depreciation-only input rather than fully "
        "reconciled total D&A. Lease treatment and "
        "market assumptions are simplified."
    )

    # --------------------------------------------------------
    # SIDEBAR ASSUMPTIONS
    # --------------------------------------------------------

    with st.sidebar:
        st.header("Valuation Assumptions")

        rf = st.slider(
            "Risk-free rate",
            0.01, 0.08, 0.041, 0.001,
            format="%.3f",
        )

        erp = st.slider(
            "Equity risk premium",
            0.03, 0.08, 0.050, 0.001,
            format="%.3f",
        )

        beta = st.slider(
            "Levered equity beta",
            0.50, 1.80, 0.95, 0.05,
        )

        debt_cost = st.slider(
            "Pre-tax cost of debt",
            0.02, 0.09, 0.048, 0.001,
            format="%.3f",
        )

        terminal = st.slider(
            "Terminal growth",
            0.0, 0.05,
            float(snapshot["terminal_growth"]),
            0.0025,
            format="%.4f",
        )

        capex_delta = st.slider(
            "Additional cash CapEx intensity",
            -0.10, 0.10, 0.0, 0.005,
            format="%.3f",
        )

        da_multiplier = st.slider(
            "D&A forecast multiplier",
            0.50, 1.50, 1.00, 0.05,
        )

    debt = snapshot["debt_m"]
    shares = snapshot["shares_m"]
    price = snapshot["price"]

    rate = wacc(
        rf,
        erp,
        beta,
        debt_cost,
        0.21,
        price * shares,
        debt,
    )

    if rate <= terminal:
        st.error(
            "WACC must exceed terminal growth."
        )
        st.stop()

    forecast["CapEx / Revenue"] += capex_delta

    forecast["D&A / Revenue"] *= da_multiplier

    # --------------------------------------------------------
    # SCENARIO VALUATION
    # --------------------------------------------------------

    revenue_2026 = float(
        history.loc[
            history["Fiscal Year"] == 2026,
            "Revenue",
        ].iloc[0]
    )

    try:
        results = {
            name: value_case(
                revenue_2026,
                forecast.to_dict("records"),
                adjustments,
                rate,
                terminal,
                snapshot["cash_m"],
                snapshot["short_term_investments_m"],
                debt,
                shares,
                price,
            )
            for name, adjustments in SCENARIOS.items()
        }

    except ValueError as error:
        st.error(str(error))
        st.stop()

    base = results["Base"]

    # --------------------------------------------------------
    # KPI CARDS
    # --------------------------------------------------------

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Base Implied Value",
        f"${base['Implied Price']:,.2f}",
    )

    c2.metric(
        "Reference Share Price",
        f"${price:,.2f}",
    )

    c3.metric(
        "Base Upside / Downside",
        f"{base['Upside']:+.1%}",
    )

    c4.metric(
        "Calculated WACC",
        f"{rate:.2%}",
    )

    # --------------------------------------------------------
    # SCENARIO COMPARISON
    # --------------------------------------------------------

    st.subheader("Scenario Valuation")

    scenario_table = pd.DataFrame([
        {
            "Scenario": name,
            "Implied Price ($)": round(
                result["Implied Price"], 2
            ),
            "Upside / Downside (%)": round(
                100 * result["Upside"], 1
            ),
            "PV Terminal Weight (%)": round(
                100 * result["PV Terminal Weight"], 1
            ),
        }
        for name, result in results.items()
    ])

    st.dataframe(
        scenario_table,
        hide_index=True,
        use_container_width=True,
    )

    st.plotly_chart(
        px.bar(
            scenario_table,
            x="Scenario",
            y="Implied Price ($)",
            title="DCF Implied Price by Scenario",
        ),
        use_container_width=True,
    )

    # --------------------------------------------------------
    # CASH FLOW AND REINVESTMENT
    # --------------------------------------------------------

    st.subheader(
        "AI Infrastructure Reinvestment"
    )

    cashflow_chart = history[
        [
            "Fiscal Year",
            "Operating Cash Flow",
            "Cash CapEx",
            "Free Cash Flow",
        ]
    ].melt(
        id_vars="Fiscal Year",
        var_name="Metric",
        value_name="USD millions",
    )

    st.plotly_chart(
        px.line(
            cashflow_chart,
            x="Fiscal Year",
            y="USD millions",
            color="Metric",
            markers=True,
            title="Operating Cash Flow vs. Cash CapEx",
        ),
        use_container_width=True,
    )

    # --------------------------------------------------------
    # FORECAST DETAIL
    # --------------------------------------------------------

    st.subheader(
        "Base-Case Forecast Bridge"
    )

    base_forecast = pd.DataFrame(
        base["Forecast"]
    )

    st.dataframe(
        base_forecast.style.format(
            precision=1
        ),
        use_container_width=True,
    )

    st.caption(
        "PV terminal value as a percentage of "
        f"enterprise value: "
        f"{base['PV Terminal Weight']:.1%}"
    )

    # --------------------------------------------------------
    # RESEARCH THESIS
    # --------------------------------------------------------

    st.subheader("Investment Thesis")

    st.markdown(
        """
        **Core research question:** Can Microsoft
        generate enough incremental cash flow from
        cloud and AI investments to justify its
        elevated infrastructure spending?

        The model examines three major drivers:

        - Revenue growth and operating leverage
        - Cash capital expenditure intensity
        - Discount rates and terminal economics
        """
    )

    st.markdown(
        "[Microsoft FY2026 Form 10-K]"
        "(https://www.sec.gov/Archives/edgar/"
        "data/789019/000119312526323660/"
        "msft-20260630.htm)"
    )

    st.caption(
        f"Quote snapshot: {snapshot['as_of_utc']} "
        f"| {snapshot['quote_source']} "
        "| Educational research only"
    )


if __name__ == "__main__":
    render()
