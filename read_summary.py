import pandas as pd
import numpy as np

def read_summary_from_excel(excel_filepath, ticker, poa_input):
    df = pd.read_excel(excel_filepath, sheet_name="Forecast Summary", header=None)
    summary_header = f"Summary Statistics - {ticker}"
    start_row = df[df.iloc[:, 0] == summary_header].index[0]

    header_row = df.iloc[start_row:].index[df.iloc[start_row:, 0] == "Statistic"][0]
    summary_data = pd.read_excel(
        excel_filepath, sheet_name="Forecast Summary", header=header_row
    ).dropna(subset=["Statistic"]).set_index("Statistic")

    stats = {
        "Revenue": {
            "median": summary_data.at["Median", f"Revenue {poa_input}"],
            "0th": summary_data.at["10th Percentile", f"Revenue {poa_input}"],
            "100th": summary_data.at["90th Percentile", f"Revenue {poa_input}"],
        },
        "EBITDA_Margin": {
            "median": summary_data.at["Median", f"EBITDA Margin {poa_input}"],
            "0th": summary_data.at["10th Percentile", f"EBITDA Margin {poa_input}"],
            "100th": summary_data.at["90th Percentile", f"EBITDA Margin {poa_input}"],
        },
        "EV_EBITDA": {
            "median": summary_data.at["Median", f"EV/EBITDA {poa_input}"],
            "0th": summary_data.at["10th Percentile", f"EV/EBITDA {poa_input}"],
            "100th": summary_data.at["90th Percentile", f"EV/EBITDA {poa_input}"],
        },
    }
    return stats
