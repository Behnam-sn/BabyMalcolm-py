from domain.record import Record
from domain.utils import Utils


class Processor:
    @staticmethod
    def generate_time_summary(records: list[Record]) -> list:
        filtered_records = Utils.filter_records_by_category(
            records=records, categories=["ورود", "خروج"]
        )
        logs = Utils.extract_logs(filtered_records)
        time_summary = Utils.log_to_dictionary(logs)
        return time_summary

    @staticmethod
    def generate_daily_summary(records: list[Record]) -> list:
        daily_summary = Utils.record_to_dictionary(records)
        return daily_summary

    @staticmethod
    def generate_weekly_summary(records: list[Record], start: int, end: int) -> dict:
        filtered_records = Utils.filter_records_by_date(records, start, end)
        weekly_summary = Utils.summarize_records(filtered_records)
        sorted_weekly_summary = Utils.sort_summary(weekly_summary)
        return sorted_weekly_summary

    @staticmethod
    def generate_monthly_summary(records: list[Record]) -> dict:
        monthly_report = Utils.summarize_records(records)
        sorted_monthly_report = Utils.sort_summary(monthly_report)
        return sorted_monthly_report
