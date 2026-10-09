
"""Tests for the Microsoft DCF valuation engine."""

import sys
import unittest
from pathlib import Path

sys.path.insert(
    0,
    str(
        Path(__file__).resolve().parents[1]
        / "case_studies"
        / "microsoft_valuation"
    ),
)

from valuation_engine import wacc, value_case


class TestValuation(unittest.TestCase):

    def setUp(self):
        self.forecast = [
            {
                "Fiscal Year": 2027 + i,
                "Revenue Growth": 0.10,
                "EBIT Margin": 0.40,
                "CapEx / Revenue": 0.10,
                "D&A / Revenue": 0.05,
                "Tax Rate": 0.21,
                "Incremental NWC Ratio": 0.02,
            }
            for i in range(5)
        ]

    def value(
        self,
        scenario=(0, 0, 0),
        wacc_rate=0.09,
        terminal=0.025,
    ):
        return value_case(
            1000,
            self.forecast,
            scenario,
            wacc_rate,
            terminal,
            200,
            100,
            150,
            100,
            10,
        )

    def test_wacc(self):
        result = wacc(
            0.04,
            0.05,
            1,
            0.05,
            0.21,
            900,
            100,
        )

        self.assertAlmostEqual(
            result,
            0.08495,
            places=6,
        )

    def test_scenario_order(self):
        bear = self.value(
            (-0.03, -0.03, 0.03)
        )["Implied Price"]

        base = self.value()["Implied Price"]

        bull = self.value(
            (0.03, 0.03, -0.02)
        )["Implied Price"]

        self.assertLess(bear, base)
        self.assertLess(base, bull)

    def test_higher_wacc_lowers_value(self):
        high = self.value(
            wacc_rate=0.11
        )["Implied Price"]

        low = self.value(
            wacc_rate=0.08
        )["Implied Price"]

        self.assertLess(high, low)

    def test_higher_capex_lowers_value(self):
        high_capex = self.value(
            (0, 0, 0.02)
        )["Implied Price"]

        base = self.value()["Implied Price"]

        self.assertLess(high_capex, base)

    def test_terminal_guard(self):
        with self.assertRaises(ValueError):
            self.value(
                wacc_rate=0.02,
                terminal=0.025,
            )

    def test_fcff_bridge(self):
        row = self.value()["Forecast"][0]

        expected = (
            row["NOPAT"]
            + row["D&A"]
            - row["Cash CapEx"]
            - row["Change NWC"]
        )

        self.assertAlmostEqual(
            row["FCFF"],
            expected,
        )


if __name__ == "__main__":
    unittest.main()
