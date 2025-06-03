import pandas as pd


def metrics(excel_df):
    excel_df.loc[:, "Колонна"] = excel_df["Колонна"].astype(str)

    plan_flights = (
        excel_df.groupby(["Маршрут", "Направление", "Дист", "Коэф", "Вмест", "Дата"])[
            "План начало"
        ]
        .count()
        .reset_index(name="План рейсов")
    )
    fact_flights = (
        excel_df[excel_df["Причина"].isnull()]
        .groupby(["Маршрут", "Направление", "Дист", "Коэф", "Вмест", "Дата"])
        .size()
        .reset_index(name="Факт рейсов")
    )

    merged_df = pd.merge(
        plan_flights,
        fact_flights[
            ["Маршрут", "Направление", "Факт рейсов", "Дист", "Коэф", "Вмест", "Дата"]
        ],
        on=["Маршрут", "Дист", "Вмест", "Коэф", "Направление", "Дата"],
        how="outer",
    )

    merged_df["Дельта рейсов"] = merged_df["План рейсов"] - merged_df["Факт рейсов"]

    merged_df["План км"] = merged_df["План рейсов"] * merged_df["Дист"]
    merged_df["Факт км"] = merged_df["Факт рейсов"] * merged_df["Дист"]
    merged_df["Дельта км"] = merged_df["План км"] - merged_df["Факт км"]

    merged_df["План стоимость"] = (
        merged_df["План рейсов"]
        * merged_df["Дист"]
        * merged_df["Коэф"]
        * merged_df["Вмест"]
    )
    merged_df["Факт стоимость"] = (
        merged_df["Факт рейсов"]
        * merged_df["Дист"]
        * merged_df["Коэф"]
        * merged_df["Вмест"]
    )
    merged_df["Дельта стоимость"] = (
        merged_df["План стоимость"] - merged_df["Факт стоимость"]
    )
    merged_df["Пот. макс. ущ"] = (
        merged_df["Дельта рейсов"] * 1000
        + merged_df["Дельта стоимость"] * 0.01
        + merged_df["Дельта стоимость"]
    )
    merged_df = merged_df.groupby(["Дата"]).sum()

    merged_df = merged_df.drop(
        ["Маршрут", "Направление", "Дист", "Коэф", "Вмест"], axis=1
    )

    return merged_df
