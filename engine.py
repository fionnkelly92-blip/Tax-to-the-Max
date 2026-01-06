import os
import json
import random

import numpy as np
import pandas as pd

COMPANIES_CSV = "companies.csv"

# -------------------------------------
# 1. LOAD COMPANIES
# -------------------------------------

def load_companies() -> pd.DataFrame:
    """
    Load company data from companies.csv.
    Fallback: built-in small example if file is missing.
    """
    if os.path.exists(COMPANIES_CSV):
        df = pd.read_csv(COMPANIES_CSV)
    else:
        # Fallback example (used only if CSV is missing)
        data = [
            {"Company": "Alpha Tech", "Sector": "Tech", "TaxContribution": 1_000_000_000, "BaseProb": 0.05},
            {"Company": "Beta Pharma", "Sector": "Pharma", "TaxContribution": 800_000_000, "BaseProb": 0.04},
            {"Company": "Gamma Finance", "Sector": "Finance", "TaxContribution": 600_000_000, "BaseProb": 0.03},
        ]
        df = pd.DataFrame(data)

    # Sanity checks & types
    required_cols = {"Company", "Sector", "TaxContribution", "BaseProb"}
    if not required_cols.issubset(df.columns):
        raise ValueError(
            f"companies.csv must contain columns {required_cols}, "
            f"but has {set(df.columns)}"
        )

    df["TaxContribution"] = df["TaxContribution"].astype(float)
    df["BaseProb"] = df["BaseProb"].astype(float)

    return df


COMPANIES = load_companies()
TOTAL_REVENUE = COMPANIES["TaxContribution"].sum()

# -------------------------------------
# 2. SCENARIOS
# -------------------------------------

SCENARIOS = {
    # scenario_name: probability multiplier
    "Mild": 0.5,    # probabilities halved
    "Medium": 1.0,  # base case
    "Severe": 1.5,  # probabilities increased by 50%
}


def _get_multiplier(scenario: str) -> float:
    if scenario not in SCENARIOS:
        raise ValueError(f"Unknown scenario '{scenario}'. Must be one of {list(SCENARIOS.keys())}")
    return SCENARIOS[scenario]


# -------------------------------------
# 3. EXPECTED LOSS TABLE
# -------------------------------------

def expected_loss_table(scenario: str = "Medium") -> pd.DataFrame:
    """
    Return a DataFrame with:
    Company, Sector, Tax Contribution, Base Prob, Adj Prob, Expected Loss
    for the given scenario.
    """
    mult = _get_multiplier(scenario)

    df = COMPANIES.copy()
    df["Tax Contribution"] = df["TaxContribution"]
    df["Base Prob"] = df["BaseProb"]
    df["Adj Prob"] = df["BaseProb"] * mult

    # Cap probabilities at 1.0 (100%), just in case
    df["Adj Prob"] = df["Adj Prob"].clip(upper=1.0)

    df["Expected Loss"] = df["TaxContribution"] * df["Adj Prob"]

    return df[[
        "Company",
        "Sector",
        "Tax Contribution",
        "Base Prob",
        "Adj Prob",
        "Expected Loss",
    ]]


# -------------------------------------
# 4. SIMPLE REVENUE PROJECTION
# -------------------------------------

def project_revenue(years: int = 10, growth: float = 0.03) -> pd.DataFrame:
    """
    Simple deterministic projection of total corporation tax revenue
    assuming a constant annual growth rate on TOTAL_REVENUE.

    Returns columns: Year, ProjectedRevenue
    """
    rows = []
    base = TOTAL_REVENUE
    for i in range(years + 1):  # include year 0
        year_num = i  # year 0, 1, 2, ...
        amount = base * ((1 + growth) ** i)
        rows.append({"Year": year_num, "ProjectedRevenue": round(amount)})

    return pd.DataFrame(rows)


# -------------------------------------
# 5. MONTE CARLO SIMULATION
# -------------------------------------

def run_monte_carlo(runs: int = 1000, scenario: str = "Medium") -> pd.DataFrame:
    """
    Monte Carlo simulation of corporation-tax shocks.

    For each run:
      - For each company, randomly decide if it 'leaves' (1) or 'stays' (0)
        based on the Adjusted Probability.
      - If it leaves, we lose its full TaxContribution.
      - Compute PostShockRevenue = TOTAL_REVENUE - sum(lost contributions).

    Returns DataFrame with one column 'PostShockRevenue' and index 0..runs-1.
    """
    df = expected_loss_table(scenario)
    probs = df["Adj Prob"].values
    contribs = df["Tax Contribution"].values

    results = []
    rng = np.random.default_rng()

    for _ in range(runs):
        # Bernoulli (0/1) for each company
        leaves = rng.binomial(n=1, p=probs)
        loss = np.sum(contribs * leaves)
        post_shock = TOTAL_REVENUE - loss
        results.append(post_shock)

    return pd.DataFrame({"PostShockRevenue": results})


# -------------------------------------
# 6. DEMO WHEN RUN DIRECTLY
# -------------------------------------

if __name__ == "__main__":
    print("Loaded companies:")
    print(COMPANIES, "\n")
    print(f"Total revenue: {TOTAL_REVENUE:.0f}\n")

    df_loss = expected_loss_table("Medium")
    print("Expected loss table (first 5 rows):")
    print(df_loss.head(), "\n")

    mc = run_monte_carlo(runs=1000, scenario="Medium")
    print("Monte Carlo summary:")
    print(mc["PostShockRevenue"].describe())
