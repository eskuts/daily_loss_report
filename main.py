import os

from utils.args_parser.parser import parse_args

from plan_fact_report.plan_fact_report import create_plan_fact_report
from colons_report.colons_report import create_colon_report
from damage_report.resp_departments_report import create_resp_dep_report
from utils.logger import setup_logging
from dotenv import load_dotenv

def main():
    load_dotenv()
    root_logger = setup_logging(log_file=os.getenv("path_to_logs"), module=None)

    try:
        args = parse_args()

        root_logger.info("creating plan-fact report")
        create_plan_fact_report(args)

        root_logger.info("creating colons report")
        create_colon_report(args)

        root_logger.info("creating damage bi resp. dep. report")
        create_resp_dep_report(args)
        return 0

    except Exception as e:
        root_logger.error(f"An error occurred: {str(e)}", exc_info=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
