import os

import psycopg2
import pandas as pd


def df_from_bd(query, start_date=None):
    try:
        connection = psycopg2.connect(os.getenv("DATABASE_URL"))

        params = (start_date,) if start_date is not None else None
        df = pd.read_sql(query, connection, params=params)

        connection.close()
        return df

    except psycopg2.Error as e:
        print(f"Ошибка при подключении или выполнении запроса: {e}")
        return pd.DataFrame()


def df_coo(start_date):
    query = """
            SELECT *
            FROM public.oper_cds AS coo
            JOIN
            transport_work.flight_price_for_oper_cds AS prices
            ON coo.id = prices.cds_id
            WHERE
                date_trunc('month', date_event) = %s::date
            AND coo.route != '557'
            AND coo.route != '531'
            AND coo.route != '462'
            AND coo.route != '665'
            AND coo.route != '670'
            AND (coo.type_event != 'Нет' OR coo.type_event IS NULL);
            """

    return df_from_bd(query, start_date)


def df_acts(start_date):

    query = """
            SELECT *
            FROM public.oper_acts AS acts
            JOIN
            transport_work.flight_price_for_oper_acts AS prices
            ON acts.id = prices.acts_id
            WHERE
                date_trunc('month', date_event) = %s::date
            AND (acts.add_info != 'Нарушение снято' OR acts.add_info IS NULL);
            """
    return df_from_bd(query, start_date)

def route_distr():
    query = """
        select *
        FROM routes_catalog.merged_routes_distributions
        """
    return df_from_bd(query)