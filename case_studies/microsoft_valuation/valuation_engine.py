
"""Independent Python FCFF valuation engine."""

from math import isfinite


def wacc(
    rf,
    erp,
    beta,
    cost_debt,
    tax,
    equity_market_value,
    debt,
):
    inputs = [
        rf,
        erp,
        beta,
        cost_debt,
        tax,
        equity_market_value,
        debt,
    ]

    if not all(isfinite(x) for x in inputs):
        raise ValueError("Nonfinite WACC input.")

    if equity_market_value <= 0 or debt < 0:
        raise ValueError("Invalid capital structure.")

    if not 0 <= tax < 1:
        raise ValueError("Invalid tax rate.")

    total = equity_market_value + debt

    equity_weight = equity_market_value / total
    debt_weight = debt / total

    cost_equity = rf + beta * erp
    after_tax_debt = cost_debt * (1 - tax)

    return (
        equity_weight * cost_equity
        + debt_weight * after_tax_debt
    )


def value_case(
    revenue0,
    forecast,
    scenario,
    wacc_rate,
    terminal_growth,
    cash,
    investments,
    debt,
    shares,
    price,
):
    if wacc_rate <= terminal_growth:
        raise ValueError(
            "WACC must exceed terminal growth."
        )

    if shares <= 0 or revenue0 <= 0 or price <= 0:
        raise ValueError(
            "Invalid market or revenue input."
        )

    growth_adj, margin_adj, capex_adj = scenario

    revenue = float(revenue0)
    pv = 0.0
    rows = []

    for i, row in enumerate(forecast, start=1):
        previous_revenue = revenue

        growth = (
            float(row["Revenue Growth"])
            + growth_adj
        )

        margin = (
            float(row["EBIT Margin"])
            + margin_adj
        )

        capex_ratio = (
            float(row["CapEx / Revenue"])
            + capex_adj
        )

        if (
            growth <= -1
            or capex_ratio < 0
            or not 0 <= margin <= 1
        ):
            raise ValueError(
                "Scenario produces invalid inputs."
            )

        revenue = previous_revenue * (1 + growth)

        ebit = revenue * margin

        nopat = ebit * (
            1 - float(row["Tax Rate"])
        )

        da = (
            revenue
            * float(row["D&A / Revenue"])
        )

        capex = revenue * capex_ratio

        nwc = (
            (revenue - previous_revenue)
            * float(row["Incremental NWC Ratio"])
        )

        fcff = nopat + da - capex - nwc

        pv_fcff = fcff / (
            (1 + wacc_rate) ** i
        )

        pv += pv_fcff

        rows.append({
            "Year": int(row["Fiscal Year"]),
            "Revenue": revenue,
            "Growth": growth,
            "EBIT Margin": margin,
            "EBIT": ebit,
            "NOPAT": nopat,
            "D&A": da,
            "Cash CapEx": capex,
            "Change NWC": nwc,
            "FCFF": fcff,
            "PV FCFF": pv_fcff,
        })

    terminal_fcff = (
        rows[-1]["FCFF"]
        * (1 + terminal_growth)
    )

    terminal_value = (
        terminal_fcff
        / (wacc_rate - terminal_growth)
    )

    pv_terminal = (
        terminal_value
        / ((1 + wacc_rate) ** len(rows))
    )

    enterprise_value = pv + pv_terminal

    equity_value = (
        enterprise_value
        + cash
        + investments
        - debt
    )

    implied_price = equity_value / shares

    return {
        "Implied Price": implied_price,
        "Upside": implied_price / price - 1,
        "Enterprise Value": enterprise_value,
        "Equity Value": equity_value,
        "PV Terminal Weight": (
            pv_terminal / enterprise_value
            if enterprise_value
            else float("nan")
        ),
        "PV Forecast FCFF": pv,
        "PV Terminal": pv_terminal,
        "Forecast": rows,
        "WACC": wacc_rate,
        "Terminal Growth": terminal_growth,
    }
