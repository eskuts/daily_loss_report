import os

import pandas as pd
from dotenv import load_dotenv
from utils.database import route_distr

from utils.args_parser.parser import parse_args


def plan_fact():
    """план-факт за месяц (Задача № 6817)"""
    args = parse_args()
    month = args.month
    excel_df = pd.read_excel(r"results\prepeared_df_month.xlsx".replace("month", month))
    plan_flights = (
        excel_df.groupby(["Дата", "Маршрут", "Направление", "Дист"])["Дата"]
        .count()
        .reset_index(name="План рейсов")
    )
    fact_flights = (
        excel_df[excel_df["Причина"].isnull()]
        .groupby(["Дата", "Маршрут", "Направление", "Дист"])
        .size()
        .reset_index(name="Факт рейсов")
    )

    merged_df = pd.merge(
        plan_flights,
        fact_flights[
            [
                "Дата",
                "Маршрут",
                "Направление",
                "Факт рейсов",
                "Дист",
            ]
        ],
        on=["Дата", "Маршрут", "Дист", "Направление"],
        how="outer",
    )

    merged_df["План км"] = merged_df["План рейсов"] * merged_df["Дист"]
    merged_df["Факт км"] = merged_df["Факт рейсов"] * merged_df["Дист"]

    merged_df = merged_df.drop(["Направление", "Дист", "Дата"], axis=1)
    merged_df = merged_df.groupby(["Маршрут"]).sum()

    merged_df.to_excel(f"results\План-факт {month}.xlsx")
    return 0


def create_fin_dep_report():
    load_dotenv()

    colon_sites_df = route_distr("2025-03-01", "2025-03-31")
    dists = pd.read_excel(os.environ["path_to_dists"])
    df_cap = pd.read_excel(os.environ["path_to_cap"], header=None)
    coefs = pd.read_excel(os.environ["path_to_coefs"], sheet_name="Sheet1")
    plan_fact()
    return 0


if __name__ == "__main__":
    raise SystemExit(create_fin_dep_report())
