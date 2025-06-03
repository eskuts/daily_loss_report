import pandas as pd
from utils.args_parser.parser import parse_args


def add_coefficients(
    acts_data: pd.DataFrame, coefficient_period: str, coefs
) -> pd.DataFrame:
    coefs["Маршрут"] = coefs["Маршрут"].astype(str).str.upper()
    coefs_for_period = coefs[["Маршрут", coefficient_period]]
    coefs_for_period = coefs_for_period.rename(columns={coefficient_period: "Коэф"})
    return acts_data.merge(coefs_for_period, on="Маршрут", how="left")


def assign_sites_and_colons(
    colons_sites_df, excel_df, route_label="Маршрут", date_label="Дата", order="Наряд"
):
    # Приводим даты в единую структуру (datetime.date)
    excel_df[date_label] = pd.to_datetime(
        excel_df[date_label], format="%d.%m.%Y", errors="coerce"
    ).dt.date

    # Группировка данных из Excel по дате, маршруту и наряду
    grouped = excel_df.groupby([date_label, route_label, order])

    # Список для хранения результатов
    results = []

    for (date_event, route, order_value), group in grouped:
        # Фильтруем по маршруту и дате
        matching_rows = colons_sites_df[
            (colons_sites_df["route"] == route)
            & (colons_sites_df["date_distr"] == date_event)
        ]

        # Если для маршрута и даты нет данных, добавляем None
        if matching_rows.empty:
            for idx in group.index:
                results.append({**group.loc[idx], "Колонна": None, "Площадка": None})
            continue

        # Фильтруем по наряду (если есть в данных)
        filtered_rows = matching_rows.dropna(subset=["order_numbers"]).copy()
        filtered_rows["order_numbers"] = filtered_rows["order_numbers"].apply(
            lambda x: list(map(int, x))
        )

        # Проверяем, есть ли строка с конкретным нарядом
        matched_row = filtered_rows[
            filtered_rows["order_numbers"].apply(lambda x: int(order_value) in x)
        ]

        if not matched_row.empty:
            # Нашли точное совпадение по наряду
            row = matched_row.iloc[0]
            column_value = row["column_num"].split()[-1]
            platform_value = row["platform"]
        else:
            # Если совпадение по наряду не найдено, берём первую запись без учета наряда
            row = matching_rows.iloc[0]
            column_value = row["column_num"].split()[-1]
            platform_value = row["platform"]

        # Добавляем найденные значения в результат
        for idx in group.index:
            results.append(
                {**group.loc[idx], "Колонна": column_value, "Площадка": platform_value}
            )

    # Создаем итоговый DataFrame
    results_df = pd.DataFrame(results)
    return results_df


def assign_distances(excel_df, dists):
    # Преобразование дат в формат datetime
    excel_df["Дата"] = pd.to_datetime(
        excel_df["Дата"], format="%d.%m.%Y", errors="coerce"
    )
    dists["Дата"] = pd.to_datetime(dists["Дата"], format="%d.%m.%Y", errors="coerce")

    # Группировка и сортировка
    grouped = excel_df.groupby(["Дата", "Маршрут", "Направление"])
    sorted_dists = dists.sort_values(by=["Маршрут", "Дата"], ascending=[False, False])

    distances_dict = {}

    for group_key, group_data in grouped:
        date_event, route, direction = group_key
        route_dists = sorted_dists[sorted_dists["Маршрут"] == route]

        # Поиск ближайшей даты
        matching_row = route_dists[route_dists["Дата"] <= date_event].head(1)

        if not matching_row.empty:
            distance = (
                matching_row.iloc[0]["от НП"]
                if direction in ["от НП", "Прямое", "НП"]
                else matching_row.iloc[0]["от КП"]
            )
            for idx in group_data.index:
                distances_dict[idx] = distance

    excel_df["Дист"] = excel_df.index.map(distances_dict)
    return excel_df


def prepared_df(excel_df, dists, df_cap, coefs, colons_sites_df):
    args = parse_args()
    excel_df = excel_df[
        ["Маршрут", "Дата", "Причина", "Направление", "План начало", "Наряд"]
    ]

    excel_df.loc[:, "Наряд"] = excel_df["Наряд"].astype(str)
    (
        excel_df.loc[:, "Дист"],
        excel_df.loc[:, "Колонна"],
        excel_df.loc[:, "Площадка"],
        excel_df.loc[:, "Вмест"],
    ) = (None, None, None, None)

    df_cap.iloc[:, 0] = df_cap[0].astype(str)
    excel_df["Причина"] = excel_df["Причина"].replace(r"(?i)\s*р/с.*", None, regex=True)
    cap_dict = dict(zip(df_cap[0], df_cap[1]))
    excel_df["Вмест"] = excel_df["Маршрут"].map(cap_dict)

    excel_df = assign_distances(excel_df, dists)

    excel_df = assign_sites_and_colons(colons_sites_df, excel_df)  # db
    # excel_df = new_assign_sites_and_colons(colons_sites_df, excel_df) #path

    coefficient_period = args.coef_period
    excel_df = add_coefficients(excel_df, coefficient_period, coefs)

    return excel_df


if __name__ == "__main__":

    raise 0
