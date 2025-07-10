import numpy as np
import pandas as pd

def compute_tsr(df, base, years):
    R0, M0, E0 = base["revenue_2024"], base["ebitda_margin_2024"], base["ev_ebitda_2024"]
    EV0, D0, S0 = base["ev_2024"], base["net_debt_2024"], base["shares_2024"]
    Y1, D1, S1 = base["div_yield_2026"], base["net_debt_2026"], base["shares_2026"]

    R1, M1, E1 = df["Revenue"], df["EBITDA Margin"], df["EV/EBITDA"]

    EV1 = R1 * M1 * E1
    cap0, cap1 = EV0 - D0, EV1 - D1
    share_price_0, share_price_1 = cap0 / S0, cap1 / S1

    df["cagr_revenue"] = (R1 / R0)**(1/years) - 1
    df["cagr_ebitda_margin"] = (M1 / M0)**(1/years) - 1
    df["cagr_ev_ebitda"] = (E1 / E0)**(1/years) - 1
    df["cagr_market_cap"] = (cap1 / cap0)**(1/years) - 1
    df["cagr_shares"] = (S1 / S0)**(1/years) - 1
    df["cagr_share_price"] = (share_price_1 / share_price_0)**(1/years) - 1
    df["dividend_return"] = Y1
    df["TSR"] = df["cagr_share_price"] + df["dividend_return"]

    return df
