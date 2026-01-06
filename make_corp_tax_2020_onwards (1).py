import pandas as pd

# Use the "keys" file – it has the month number & tax head code
src = "Open Data Tax Receipts Keys.csv"

df = pd.read_csv(
    src,
    header=None,
    names=[
        "Year",
        "MonthNo",
        "Month",
        "PeriodKey",
        "Type",
        "TaxHeadCode",
        "TaxHead",
        "Amount",
    ],
)

# Keep only Actual Outturn + Corporation Tax
mask = (df["Type"] == "Actual Outturn") & (df["TaxHead"] == "Corporation Tax")
corp = df.loc[mask].copy()

# Filter to 2020 onwards
corp = corp[corp["Year"] >= 2020].reset_index(drop=True)

print(corp.head())
print("Years in file:", sorted(corp["Year"].unique()))

# Save a tidy CSV you can reuse
corp.to_csv("corp_tax_2020_onwards.csv", index=False)
print("Saved corp_tax_2020_onwards.csv")
