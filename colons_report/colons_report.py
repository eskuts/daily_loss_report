import os

import gspread
import pandas as pd
from dotenv import load_dotenv
from gspread_dataframe import set_with_dataframe
from gspread_formatting import (CellFormat, Color, NumberFormat, TextFormat,
                                format_cell_range)
from utils.args_parser.parser import parse_args
load_dotenv()


def metrics(excel_df):
    plan_flights = (
        excel_df.groupby(["Маршрут", "Колонна", "Дата", "Площадка"])["План начало"]
        .count()
        .reset_index(name="План")
    )
    fact_flights = (
        excel_df[excel_df["Причина"].isnull()]
        .groupby(["Маршрут", "Колонна", "Дата", "Площадка"])
        .size()
        .reset_index(name="Факт")
    )

    merged_df = pd.merge(
        plan_flights,
        fact_flights[["Площадка", "Маршрут", "Колонна", "Дата", "Факт"]],
        on=["Площадка", "Маршрут", "Колонна", "Дата"],
        how="outer",
    )
    merged_df["% выполнения"] = (100 * merged_df["Факт"] / merged_df["План"]).round(2)

    return merged_df


def create_pivot(res):
    # Группируем данные: суммируем факт и план
    grouped = res.groupby(["Колонна", "Маршрут", "Дата"], as_index=False).agg(
        {"Факт": "sum", "План": "sum"}
    )
    grouped["% выполнения"] = (grouped["Факт"] / grouped["План"] * 100).round(2)

    # Создаём сводную таблицу по маршрутам
    pivot = grouped.pivot_table(
        index=["Колонна", "Маршрут"], columns="Дата", values="% выполнения"
    ).round(2)

    # --- Считаем итоги по колонне (по дням) ---
    subtotal_colon = grouped.groupby(["Колонна", "Дата"], as_index=False).sum()
    subtotal_colon["% выполнения"] = (
        subtotal_colon["Факт"] / subtotal_colon["План"] * 100
    ).round(2)
    subtotal_colon["Маршрут"] = "Итог по колонне"
    subtotal_colon = subtotal_colon.pivot_table(
        index=["Колонна", "Маршрут"], columns="Дата", values="% выполнения"
    ).round(2)

    # --- Считаем итоги по колонне (за месяц) ---
    total_monthly_colon = grouped.groupby(["Колонна"], as_index=False).sum()

    total_monthly_colon["% выполнения"] = (
        total_monthly_colon["Факт"] / total_monthly_colon["План"] * 100
    ).round(2)
    total_monthly_colon = total_monthly_colon.set_index("Колонна")["% выполнения"]

    # Объединяем маршруты и итоги по колонне
    final_pivot = pd.concat([pivot, subtotal_colon]).sort_index(level=0)

    # --- Добавляем "Итог за месяц" ---
    # Для маршрутов
    total_monthly_routes = grouped.groupby(["Колонна", "Маршрут"], as_index=False).sum()
    total_monthly_routes["% выполнения"] = (
        total_monthly_routes["Факт"] / total_monthly_routes["План"] * 100
    ).round(2)
    total_monthly_routes = total_monthly_routes.set_index(["Колонна", "Маршрут"])[
        "% выполнения"
    ]

    # Добавляем рассчитанные значения в итоговую таблицу
    final_pivot["Итог за месяц"] = total_monthly_routes
    total_monthly_colon = total_monthly_colon.to_frame(name="Итог за месяц")
    total_monthly_colon["Маршрут"] = "Итог по колонне"
    total_monthly_colon.set_index("Маршрут", append=True, inplace=True)
    final_pivot.update(total_monthly_colon)

    return final_pivot


def create_summary(res):
    # Группируем по колоннам
    total_colon = res.groupby(["Площадка", "Колонна"], as_index=False).sum()
    total_colon["% выполнения"] = (
        total_colon["Факт"] / total_colon["План"] * 100
    ).round(2)
    total_colon = total_colon[["Площадка", "Колонна", "% выполнения"]].set_index(
        ["Колонна"]
    )

    # Группируем по площадкам
    total_site = res.groupby("Площадка", as_index=False).sum()
    total_site["% выполнения"] = (total_site["Факт"] / total_site["План"] * 100).round(
        2
    )
    total_site = total_site[["Площадка", "% выполнения"]].set_index("Площадка")

    return pd.concat([total_colon, total_site])


def download_df_to_ws(worksheet, df):
    # worksheet.clear()
    set_with_dataframe(worksheet, df)


def clear_data_and_format(worksheet):
    request_body = {
        "requests": [
            {
                "unmergeCells": {"range": {"sheetId": worksheet.id}}
            },  # Разъединяем все ячейки
            {
                "updateCells": {
                    "range": {"sheetId": worksheet.id},
                    "fields": "userEnteredValue,userEnteredFormat",  # Очищаем данные и форматы
                }
            },
        ]
    }
    worksheet.spreadsheet.batch_update(request_body)


def format_google_sheet(worksheet, df):
    """Оптимальное форматирование Google Sheets через batch update."""
    requests = [
        {
            "updateCells": {
                "range": {"sheetId": worksheet.id},  # ID листа
                "fields": "userEnteredFormat",
            }
        }
    ]

    worksheet.spreadsheet.batch_update({"requests": requests})

    # Форматирование заголовков (первая строка)
    header_format = CellFormat(
        backgroundColor=Color(1, 1, 0.8),
        textFormat=TextFormat(bold=True),
        numberFormat=NumberFormat(type="DATE", pattern="dd.mm"),
    )
    format_cell_range(worksheet, "1:1", header_format)

    num_rows, num_cols = df.shape  # Количество строк и столбцов
    start_row = 2  # Начинаем с 2-й строки (1-я — заголовки)

    requests = []  # Массив запросов для batch update

    # 1. Определяем диапазоны для объединения ячеек в первом столбце
    merge_requests = []
    values = df.iloc[:, 0].tolist()
    prev_value = None
    merge_start = start_row

    for i, value in enumerate(values, start=start_row):
        if value == prev_value:
            continue
        if prev_value is not None:
            merge_requests.append(
                {
                    "sheetId": worksheet.id,  # Добавляем ID листа
                    "startRowIndex": merge_start - 1,
                    "endRowIndex": i - 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                }
            )
        merge_start = i
        prev_value = value

    merge_requests.append(
        {
            "sheetId": worksheet.id,  # Добавляем ID листа
            "startRowIndex": merge_start - 1,
            "endRowIndex": num_rows,
            "startColumnIndex": 0,
            "endColumnIndex": 1,
        }
    )

    for merge in merge_requests:
        requests.append({"mergeCells": {"range": merge, "mergeType": "MERGE_ALL"}})

    worksheet.spreadsheet.batch_update({"requests": requests})

    even_rows = [
        {
            "sheetId": worksheet.id,  # Добавляем ID нужного листа
            "startRowIndex": i,
            "endRowIndex": i + 1,
            "startColumnIndex": 1,
            "endColumnIndex": num_cols,
        }
        for i in range(start_row, num_rows + 1, 2)
    ]

    for row in even_rows:
        requests.append(
            {
                "repeatCell": {
                    "range": row,
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 0.94,
                                "green": 0.94,
                                "blue": 0.94,
                            }
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor",
                }
            }
        )

    worksheet.spreadsheet.batch_update({"requests": requests})

    # 3. Делаем жирным строки с "Итог"
    routes_indices = df[df.iloc[:, 1].str.contains("Итог", case=False, na=False)].index

    # Определяем диапазоны на основе фактического количества столбцов в DataFrame
    num_columns = df.shape[1]  # Количество столбцов
    col_letter_end = chr(
        66 + num_columns
    )  # Преобразуем в буквенное обозначение (A-Z, затем AA и т. д.)

    routes_ranges = [f"A{i + 2}:{col_letter_end}{i + 2}" for i in routes_indices]

    # Форматирование строк с "Итог"
    total_row_format = {
        "backgroundColor": {"red": 1, "green": 1, "blue": 0.8},
        "textFormat": {"bold": True},
    }

    for range_str in routes_ranges:
        start_row = int(range_str.split(":")[0][1:]) - 1  # Индекс строки (начало)
        end_row = int(range_str.split(":")[1][1:])  # Индекс строки (конец)

        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet.id,  # ID листа
                        "startRowIndex": start_row,
                        "endRowIndex": end_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": num_columns,  # Подстраивается под ширину DataFrame
                    },
                    "cell": {"userEnteredFormat": total_row_format},
                    "fields": "userEnteredFormat",
                }
            }
        )

    worksheet.spreadsheet.batch_update({"requests": requests})
    # # 4. Делаем последний столбец жирным
    last_column_format = {
        "backgroundColor": {"red": 1, "green": 1, "blue": 0.8},  # Жёлтый
        "textFormat": {"bold": True},  # Жирный текст
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE",
    }

    # Генерация запроса для форматирования последнего столбца
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 0,  # Начало с первой строки
                    "endRowIndex": df.shape[0],  # До конца таблицы
                    "startColumnIndex": num_columns - 1,  # Последний столбец
                    "endColumnIndex": num_columns,
                },
                "cell": {"userEnteredFormat": last_column_format},
                "fields": "userEnteredFormat",
            }
        }
    )
    # Отправляем batch update одним запросом
    if requests:
        worksheet.spreadsheet.batch_update({"requests": requests})

    cell_format = CellFormat(
        wrapStrategy="WRAP", horizontalAlignment="CENTER", verticalAlignment="MIDDLE"
    )
    format_cell_range(worksheet, "A:AZ", cell_format)


def create_colon_report(args):
    month = args.month
    path = os.environ[r"path_to_results"] + r"\prepared_df_month.xlsx".replace(
        "month", month
    )
    df = pd.read_excel(path)

    res = metrics(df)
    res["Дата"] = pd.to_datetime(res["Дата"]).dt.strftime("%d.%m")

    final_pivot = create_pivot(res).reset_index()

    sheet_key = os.environ["colon_report_id"]
    client = gspread.service_account(os.environ["path_to_serv_acc"])
    spreadsheet = client.open_by_key(sheet_key).worksheet(month)

    clear_data_and_format(spreadsheet)
    download_df_to_ws(spreadsheet, final_pivot)

    format_google_sheet(spreadsheet, final_pivot)

    sum_pivot = create_summary(res).reset_index()
    download_df_to_ws(client.open_by_key(sheet_key).worksheet("площадки"), sum_pivot)

    return 0


if __name__ == "__main__":
    args = parse_args()
    raise SystemExit(create_colon_report(args))
