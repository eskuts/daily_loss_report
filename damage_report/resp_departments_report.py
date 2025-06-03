import calendar
import os

import gspread
import numpy as np
import pandas as pd
from dotenv import load_dotenv

from utils.args_parser.parser import parse_args
from utils.consts.consts import *
from utils.database import df_acts, df_coo


def add_responsible_dep(df):

    df["responsible_department"] = np.where(
        df["responsible_department"].isna(),
        df["event"].map(event_to_dep),
        df["responsible_department"],
    )
    return df


def add_potential_damage(df):

    mask = df["responsible_department"].isna() & (df["potential_damage"] == 0)

    # Условия для расчёта штрафа
    conditions = [
        df.loc[mask, "event"].isin(non_penalty_events),  # Событие в non_penalty_events
        ~df.loc[mask, "event"].isin(
            non_penalty_events
        ),  # Событие не в non_penalty_events
    ]

    penalty_values = [
        df.loc[mask, "flight_price"],  # Ущерб равен стоимости рейса
        df.loc[mask, "flight_price"]
        + (df.loc[mask, "flight_price"] * 0.01)
        + 1000,  # Стоимость + 1% + 1000
    ]

    # Применение условий и обновление potential_damage
    df.loc[mask, "potential_damage"] = np.select(conditions, penalty_values)
    return df


def group_by_dep(df, name):
    conditions = [
        df["responsible_department"] == "Ремзона",  # 1) Ремзона
        df["responsible_department"] == "ОЭДО",  # 2) ОЭДО
        (df["responsible_department"] == "Эксплуатация")  # 3) Отсутствие водителя
        & (~df["reason_event"].str.contains("Запланированный срыв", na=False)),
        (
            (
                df["responsible_department"].str.contains(
                    r"^(ОЭ|Эксплуатация|Колонна № \d+)$", na=False, regex=True
                )
            )  # 4) ОЭ
            & (~df["reason_event"].str.contains("Отсутствие водителя", na=False))
        ),
        df["responsible_department"].isin(
            ["ОБД", "Топливо", "Другое", "ОБД/Топливо/Другое", "ЦДС"]
        ),  # 5) Другое
    ]

    categories = ["Ремзона", "ОЭДО", "Отсутствие водителя", "ОЭ", "Другое"]

    df["group"] = np.select(conditions, categories, default="Неизвестно")
    unknown_rows = df[df["group"] == "Неизвестно"]
    for i, condition in enumerate(conditions, start=1):
        unknown_rows[f"condition_{i}"] = condition[unknown_rows.index]

    grouped = df.groupby(["site", "date_event", "group"], as_index=False).agg(  #
        total_damage=("potential_damage", "sum"), unique_count=("id", "nunique")
    )
    grouped["date_event"] = pd.to_datetime(grouped["date_event"])
    return grouped


def get_month_dates(month_year):
    month, year = map(int, month_year.split("_"))
    _, last_day = calendar.monthrange(year, month)
    start_date = f"{year}-{month:02d}-01"
    end_date = f"{year}-{month:02d}-{last_day}"
    return start_date, end_date


def create_result_table(
    month_year, op3m_data, forecast_data, confirmed_data, aggregate_by_site=False
):

    start_date, end_date = get_month_dates(month_year)
    dates = pd.date_range(start_date, end_date)
    responsibilities = ["Отсутствие водителя", "Ремзона", "ОЭ", "Другое", "ОЭДО"]
    metrics_order = [
        "unique_count_forecast",
        "total_damage_forecast",
        "forecast_damage_percent",
        "total_damage_confirmed",
    ]

    if aggregate_by_site:
        forecast_data = forecast_data.groupby(
            ["date_event", "group"], as_index=False
        ).sum()
        confirmed_data = confirmed_data.groupby(
            ["date_event", "group"], as_index=False
        ).sum()

    strict_index = pd.MultiIndex.from_product(
        [dates, responsibilities], names=["date_event", "group"]
    )
    result_df = pd.DataFrame(index=strict_index).reset_index()

    result_df = result_df.merge(forecast_data, on=["date_event", "group"], how="left")

    result_df = result_df.merge(
        confirmed_data,
        on=["date_event", "group"],
        how="left",
        suffixes=("_forecast", "_confirmed"),
    )

    result_df = result_df.merge(
        op3m_data, right_on=["Дата"], left_on=["date_event"], how="left"
    )

    # Рассчитываем % прогнозного ущерба к общему
    result_df["forecast_damage_percent"] = (
        result_df["total_damage_forecast"] / result_df["Пот. макс. ущ"]
    )

    # Заполняем пропущенные значения нулями
    result_df = result_df.fillna(0)
    pivot = result_df.pivot(index="date_event", columns="group").fillna("")
    pivot = pivot.swaplevel(axis=1).sort_index(axis=1)

    pivot = pivot.loc[:, (responsibilities, metrics_order)]
    empty_rows = (pivot == 0).all(axis=1)
    next_row_empty = empty_rows.shift(-1, fill_value=False)
    pivot.loc[empty_rows & next_row_empty] = ""
    pivot.loc[empty_rows & ~next_row_empty] = ""  # чтобы посл строка 0 не заполнялась

    return pivot


def add_site(df):
    df["site"] = df["column_num"].map(dict_site)
    return df


def result_for_site(forecast_df, confirm_df, month_year, month_name, wks, *op3m_data):
    sites = ["Витебский", "Горская", "Зеленогорск"]
    cells = ["O38", "O72", "O106"]

    for i, site in enumerate(sites):
        op3m_data = pd.read_excel(
            f"results/op3m_data_{month_name}_{site}.xlsx"
        )
        filtered_forecast_df = forecast_df[forecast_df["site"] == site]
        filtered_forecast_df.drop(columns="site")
        filtered_confirm_df = confirm_df[confirm_df["site"] == site]
        filtered_confirm_df.drop(columns="site")

        result = create_result_table(
            month_year, op3m_data, filtered_forecast_df, filtered_confirm_df
        )

        result_cleaned = result.replace([np.inf, -np.inf, np.nan], None)
        total = result_cleaned.values.tolist()
        wks.update(cells[i], total)
    return 0


def format_table(sheet):
    requests = [
        # Формат для чисел с разделением тысяч и двумя знаками после запятой
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet.id,
                    "startColumnIndex": col_index,
                    "endColumnIndex": col_index + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        }
        for col_index in [
            15,
            17,
            19,
            21,
            23,
            25,
            27,
            29,
            31,
            33,
        ]  # P, R, T, V, X, Z, AB, AD, AF, AH
    ] + [
        # Формат для процентов с двумя знаками после запятой
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet.id,
                    "startColumnIndex": col_index,
                    "endColumnIndex": col_index + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        }
        for col_index in [16, 20, 24, 28, 32]  # Q, U, Y, AC, AG
    ]

    # Отправляем запрос через batchUpdate
    sheet.spreadsheet.batch_update({"requests": requests})


def create_resp_dep_report(args):
    load_dotenv()

    month_year = args.month_year
    month, year = map(int, month_year.split("_"))
    month_name = months[month].capitalize()
    start_date, end_date = get_month_dates(month_year)

    op3m = os.environ["path_to_op3m_report"].replace('month', month_name)
    op3m_data = pd.read_excel(op3m)

    confirmed = add_site(df_acts(start_date))
    confirmed = add_potential_damage(
        confirmed
    )  # сначала ущерб, потому что его только на необработанные строки
    new = add_responsible_dep(confirmed)
    confirmed_result = group_by_dep(new, "confirmed")

    forecast = add_site(df_coo(start_date))
    forecast = add_potential_damage(
        forecast
    )  # сначала ущерб, потому что его только на необработанные строки
    new = add_responsible_dep(forecast)
    forecast_result = group_by_dep(new, "forecast")

    result = create_result_table(
        month_year, op3m_data, forecast_result, confirmed_result, True
    )

    result_cleaned = result.replace([np.inf, -np.inf, np.nan], None)
    result_total = result_cleaned.values.tolist()

    client = gspread.service_account(os.environ["path_to_serv_acc"])
    wks = client.open_by_key(os.environ["daily_loss_report_id"]).worksheet(month_name)
    wks.update("O4", result_total)

    result_for_site(forecast_result, confirmed_result, month_year, month_name, wks)
    format_table(wks)
    return 0

if __name__ == "__main__":
    args = parse_args()
    raise SystemExit(create_resp_dep_report(args))