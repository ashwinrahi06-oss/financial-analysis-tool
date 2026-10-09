
"""Microsoft equity research: disclosed reference figures and analyst assumptions."""

from pathlib import Path

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
DATA.mkdir(parents=True, exist_ok=True)

SEC_URL = (
    "https://data.sec.gov/api/xbrl/companyfacts/"
    "CIK0000789019.json"
)

FILING_URL = (
    "https://www.sec.gov/Archives/edgar/data/789019/"
    "000119312526323660/msft-20260630.htm"
)

YEARS = list(range(2022, 2027))
FORECAST_YEARS = list(range(2027, 2032))

# FY2026 disclosed reference amounts, USD millions.
# Depreciation is not necessarily complete D&A.
FY26_DEPRECIATION = 34300.0
FY26_DA_AND_OTHER = 38534.0
FY26_CASH_CAPEX_REFERENCE = 115948.0
FY26_FINANCE_LEASE_ADDITIONS = 24608.0
FY26_FINANCE_LEASE_LIABILITIES = 66594.0
FY26_OPERATING_LEASE_LIABILITIES = 21925.0

# Explicitly provisional depreciation-only D&A input.
DA_BASE = FY26_DEPRECIATION

# Illustrative forecasts, not Microsoft guidance.
GROWTH_DECAY = [0.93, 0.85, 0.77, 0.69, 0.61]
MARGIN_DELTAS = [-0.010, -0.012, -0.012, -0.010, -0.008]
CAPEX_RATIO_MULTIPLIERS = [1.06, 1.00, 0.91, 0.81, 0.72]
DA_RATIO_MULTIPLIERS = [1.10, 1.18, 1.23, 1.23, 1.20]

TAX_RATE = 0.21
NWC_RATIO = 0.02

# Editable, illustrative cost-of-capital inputs.
RISK_FREE = 0.041
EQUITY_RISK_PREMIUM = 0.050
LEVERED_BETA = 0.95
PRETAX_DEBT_COST = 0.048
TERMINAL_GROWTH = 0.025

# Revenue growth, EBIT margin, CapEx/revenue adjustments.
SCENARIOS = {
    "Bear": (-0.035, -0.025, 0.025),
    "Base": (0.000, 0.000, 0.000),
    "Bull": (0.025, 0.020, -0.025),
}
