excel_filepath = "Combined_Forecast_Summary_With_Linking.xlsx"
ticker = "CRDA.L"
poa_input = "CY2026"
years = 2  # Example period (2024-2026)
n_simulations = 10000
tsr_probs = [0.8, 0.5, 0.2]

base = {
    "revenue_2024": 1500.0,
    "ebitda_margin_2024": 0.28,
    "ev_ebitda_2024": 10.5,
    "ev_2024": 4500.0,
    "net_debt_2024": 500.0,
    "shares_2024": 138.0,
    "div_yield_2026": 0.03,
    "net_debt_2026": 500.0,
    "shares_2026": 139.0
}
