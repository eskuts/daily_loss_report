import os

import gspread
import pandas as pd
from dotenv import load_dotenv
from src.calculate_metrics import metrics
from utils.database import route_distr
from src.prepare_df import prepared_df

from utils.args_parser.parser import parse_args


def create_plan_fact_report(args) -> int:
    load_dotenv()

    path = os.environ["path_to_op3m_report"].replace("month", args.month)
    excel_df = pd.read_excel(path, sheet_name="CSV без 557")

    client = gspread.service_account(os.environ["path_to_serv_acc"])

    colons_sites_df = route_distr()

    dists = pd.read_excel(os.environ["path_to_dists"])
    df_cap = pd.read_excel(os.environ["path_to_cap"], header=None)
    coefs = pd.read_excel(os.environ["path_to_coefs"], sheet_name="Sheet1")

    excel_df.loc[:, "Маршрут"] = excel_df["Маршрут"].astype(str)
    excel_df = excel_df[excel_df["Маршрут"] != "557"]

    df = prepared_df(excel_df, dists, df_cap, coefs, colons_sites_df)
    df_name = f"results/prepared_df_{args.month}.xlsx"
    df.to_excel(df_name)

    # df = pd.read_excel(df_name)

    result = metrics(df)
    result.to_excel(f"results/op3m_data_{args.month}.xlsx")

    result_total = result.values.tolist()

    wks = client.open_by_key(os.environ["daily_loss_report_id"]).worksheet(args.month)
    wks.update("C4", result_total)

    sites = ["Витебский", "Горская", "Зеленогорск"]

    cells_results_dict = {
        "total": ["C38", "C72", "C106"],
    }
    for i, site in enumerate(sites):
        filtered_df = df[df["Площадка"] == site]
        result_df = metrics(filtered_df)
        result_df.to_excel(f"results/op3m_data_{args.month}_{site}.xlsx")
        total = result_df.values.tolist()
        wks.update(cells_results_dict["total"][i], total)

    return 0


if __name__ == "__main__":
    args = parse_args()
    raise SystemExit(create_plan_fact_report(args))

# python plan_fact_report.py --coef_period "01.04.2025-30.06.2025" --month "Апрель"
