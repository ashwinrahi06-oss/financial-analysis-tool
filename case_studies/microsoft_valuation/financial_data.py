
"""SEC Company Facts extraction and historical source audit."""

import os
from datetime import date

import pandas as pd
import requests

from .model_config import SEC_URL, YEARS


FLOW = {
    "Revenue": [
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "Revenues",
        "SalesRevenueNet",
    ],
    "Operating Income": ["OperatingIncomeLoss"],
    "Net Income": ["NetIncomeLoss"],
    "Operating Cash Flow": [
        "NetCashProvidedByUsedInOperatingActivities"
    ],
    "Cash CapEx": [
        "PaymentsToAcquirePropertyPlantAndEquipment"
    ],
}

BALANCE = {
    "Cash": [
        "CashAndCashEquivalentsAtCarryingValue"
    ],
    "Short-Term Investments": [
        "ShortTermInvestments"
    ],
    "Current Debt": [
        "LongTermDebtCurrent",
        "LongTermDebtAndCapitalLeaseObligationsCurrent",
    ],
    "Long-Term Debt": [
        "LongTermDebtNoncurrent"
    ],
}


def fetch_facts():
    agent = os.getenv("SEC_USER_AGENT", "").strip()

    if "@" not in agent or "example.com" in agent:
        raise ValueError(
            "Set SEC_USER_AGENT to your real name and "
            "contact email before downloading SEC data."
        )

    response = requests.get(
        SEC_URL,
        headers={
            "User-Agent": agent,
            "Accept-Encoding": "gzip, deflate",
        },
        timeout=45,
    )

    response.raise_for_status()
    return response.json()


def pick(facts, tags, year, instant=False):
    end = f"{year}-06-30"
    candidates = []

    for rank, tag in enumerate(tags):
        records = (
            facts.get("facts", {})
            .get("us-gaap", {})
            .get(tag, {})
            .get("units", {})
            .get("USD", [])
        )

        for record in records:
            if (
                record.get("form") != "10-K"
                or record.get("end") != end
                or record.get("val") is None
            ):
                continue

            start = record.get("start")

            if instant and start:
                continue

            if not instant:
                if not start:
                    continue

                try:
                    duration = (
                        date.fromisoformat(end)
                        - date.fromisoformat(start)
                    ).days
                except ValueError:
                    continue

                if not 350 <= duration <= 380:
                    continue

            filed = record.get("filed", "")

            candidates.append(
                (
                    0 if filed.startswith(str(year)) else 1,
                    rank,
                    filed,
                    tag,
                    record,
                )
            )

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: (x[0], x[1], x[2])
    )

    chosen = candidates[0]
    record = chosen[4]

    return {
        "value": float(record["val"]) / 1e6,
        "tag": chosen[3],
        "filed": record.get("filed"),
        "accession": record.get("accn"),
    }


def load_financials(facts):
    rows = []
    audit = []

    for year in YEARS:
        item = {"Fiscal Year": year}

        for metric, tags in FLOW.items():
            fact = pick(facts, tags, year)
            value = None if fact is None else fact["value"]

            if metric == "Cash CapEx" and value is not None:
                value = abs(value)

            item[metric] = value

            audit.append([
                year,
                metric,
                fact["tag"] if fact else "MISSING",
                fact["filed"] if fact else None,
                fact["accession"] if fact else None,
                value,
            ])

        rows.append(item)

    history = pd.DataFrame(rows).set_index(
        "Fiscal Year"
    )

    if history.isna().any().any():
        raise ValueError(
            "Missing required SEC historical values. "
            "Review the source audit and tag mapping."
        )

    history["Free Cash Flow"] = (
        history["Operating Cash Flow"]
        - history["Cash CapEx"]
    )

    balance = {}

    for metric, tags in BALANCE.items():
        fact = pick(
            facts,
            tags,
            2026,
            instant=True,
        )

        balance[metric] = (
            fact["value"] if fact else None
        )

        audit.append([
            2026,
            metric,
            fact["tag"] if fact else "MISSING",
            fact["filed"] if fact else None,
            fact["accession"] if fact else None,
            balance[metric],
        ])

    required_balance = [
        "Cash",
        "Current Debt",
        "Long-Term Debt",
    ]

    if any(
        balance[key] is None
        for key in required_balance
    ):
        raise ValueError(
            "Missing required balance-sheet data. "
            "Refusing to silently assume zero."
        )

    audit_df = pd.DataFrame(
        audit,
        columns=[
            "Fiscal Year",
            "Metric",
            "SEC Tag",
            "Filed",
            "Accession",
            "USD millions",
        ],
    )

    return history, balance, audit_df
