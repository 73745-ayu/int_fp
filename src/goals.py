import numpy as np
import pandas as pd
from scipy.optimize import brentq
from src.tsr import TSRCalculator 

def find_equal_p(df, base, years, tsr_probs, tol=1e-6):
    rev, marg, mult = df["Revenue"], df["EBITDA Margin"], df["EV/EBITDA"]
    tsr_s, D1, S1 = df["TSR"], base["net_debt_2026"], base["shares_2026"]

    def tsr_at(p):
        row = pd.DataFrame({
            "Revenue": [rev.quantile(1 - p)],
            "EBITDA Margin": [marg.quantile(1 - p)],
            "EV/EBITDA": [mult.quantile(1 - p)]
        })
        return compute_tsr(row, base, years)["TSR"].iloc[0]

    results = []
    for p_tsr in tsr_probs:
        target = tsr_s.quantile(1 - p_tsr)
        try:
            p_metric = brentq(lambda p: tsr_at(p) - target, tol, 1 - tol)
        except ValueError:
            p_metric = np.nan

        results.append({
            "p_tsr": p_tsr,
            "Revenue": rev.quantile(1 - p_metric),
            "p_revenue": p_metric,
            "EBITDA Margin": marg.quantile(1 - p_metric),
            "p_margin": p_metric,
            "EV/EBITDA": mult.quantile(1 - p_metric),
            "p_multiple": p_metric,
            "Market Cap": mult.quantile(1 - p_metric) * rev.quantile(1 - p_metric) * marg.quantile(1 - p_metric) - D1,
            "Share price": (mult.quantile(1 - p_metric) * rev.quantile(1 - p_metric) * marg.quantile(1 - p_metric) - D1) / S1,
            "TSR": target,
            "Probability": p_metric
        })

    return pd.DataFrame(results).set_index("p_tsr")
