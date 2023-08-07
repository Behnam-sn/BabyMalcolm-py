from application.utils import Utils
from domain.processor import Processor
from keys import Keys


class Application:
    @staticmethod
    def generate_time_report() -> None:
        records = Utils.extract_records_from_excel(Keys.EXCEL_FILE_PATH)
        summary = Processor.generate_time_summary(records)
        Utils.list_to_json_file(summary, Keys.TIME_REPORT_FILE_NAME)
        Utils.list_to_text_file(summary, Keys.TIME_REPORT_FILE_NAME)

    @staticmethod
    def generate_daily_report() -> None:
        records = Utils.extract_records_from_excel(Keys.EXCEL_FILE_PATH)
        summary = Processor.generate_daily_summary(records)
        Utils.list_to_json_file(summary, Keys.DAILY_REPORT_FILE_NAME)

    @staticmethod
    def generate_weekly_report(start: int, end: int) -> None:
        records = Utils.extract_records_from_excel(Keys.EXCEL_FILE_PATH)
        summary = Processor.generate_weekly_summary(records, start, end)
        Utils.dictionary_to_json_file(summary, Keys.WEEKLY_REPORT_FILE_NAME)
        Utils.dictionary_to_text_file(summary, Keys.WEEKLY_REPORT_FILE_NAME)

    @staticmethod
    def generate_monthly_report() -> None:
        records = Utils.extract_records_from_excel(Keys.EXCEL_FILE_PATH)
        summary = Processor.generate_monthly_summary(records)
        Utils.dictionary_to_json_file(summary, Keys.MONTHLY_REPORT_FILE_NAME)
        Utils.dictionary_to_text_file(summary, Keys.MONTHLY_REPORT_FILE_NAME)
