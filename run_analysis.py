from src.monte_carlo import simulate
from src.tsr import compute_tsr
from src.goals import find_equal_p
from read_summary import read_summary_from_excel
import config

def main():
    company_data = read_summary_from_excel(config.excel_filepath, config.ticker, config.poa_input)
    df_simulated = simulate(company_data, config.n_simulations)
    df_tsr = compute_tsr(df_simulated, config.base, config.years)
    results_table = find_equal_p(df_tsr, config.base, config.years, config.tsr_probs)

    results_table.to_csv("multi_goalseek_output.csv")
    print(results_table.round(4))

if __name__ == "__main__":
    main()
