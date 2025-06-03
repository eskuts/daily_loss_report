import argparse


def parse_args() -> argparse.Namespace:
    """
    аргументы, передаваемые через терминал
    """
    parser = argparse.ArgumentParser(
        # description="Get current state of the processing of Acts from OrgP"
    )
    parser.add_argument(
        "--coef_period",
        # nargs="+",
        required=True,
        type=str,
        # action="extend",
        # help="Sheet names to include in update of state",
    )
    parser.add_argument(
        "--month",
        # nargs="+",
        required=True,
        type=str,
        # action="extend",
        # help="Sheet names in act (excel file) to include in update of state",
    )
    parser.add_argument(
        "--month_year",
        # nargs="+",
        required=True,
        type=str,
        # action="extend",
        # help="Sheet names to include in update of state",
    )
    return parser.parse_args()
