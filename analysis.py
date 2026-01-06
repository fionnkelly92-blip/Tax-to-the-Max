import matplotlib.pyplot as plt
from engine import expected_loss_table

df = expected_loss_table("Medium")
print(df)
print("Total expected loss:", df["Expected Loss"].sum())
print("Post-shock revenue:", df["Tax Contribution"].sum() - df["Expected Loss"].sum())

df = expected_loss_table("Severe")
print(df)

from engine import run_monte_carlo, TOTAL_REVENUE

mc = run_monte_carlo(runs=10000, scenario="Medium")

threshold = TOTAL_REVENUE * 0.75  # a 25% drop
prob_drop = (mc["PostShockRevenue"] < threshold).mean()

print("Probability revenue falls ≥25%:", prob_drop)

from excel_generator import generate_excel

generate_excel(
    output_file="Tax_Shock_Model.xlsx",
    scenario="Medium",
    runs=2000,
    growth=0.03
)



plt.hist(mc["PostShockRevenue"], bins=40)
plt.title("Monte Carlo Distribution")
plt.show()