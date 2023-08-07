from domain.log import Log
from domain.record import Record
from keys import Keys


class Utils:
    @staticmethod
    def log_to_dictionary(logs: list[Log]) -> list:
        temp = []

        for log in logs:
            temp.append(log.convert_to_dictionary())

        return temp

    # @staticmethod
    # def record_to_dictionary(records: list[Record]) -> dict:
    #     dictionary = {}

    #     for record in records:
    #         dictionary[record.id] = record.convert_to_dictionary()

    #     return dictionary

    @staticmethod
    def record_to_dictionary(records: list[Record]) -> list:
        temp = []

        for record in records:
            temp.append(record.convert_to_dictionary())

        return temp

    @staticmethod
    def filter_records_by_category(
        records: list[Record], categories: list[str]
    ) -> list[Record]:
        filtered_records = filter(
            lambda item: item.category in categories,
            records,
        )

        return list(filtered_records)

    @staticmethod
    def filter_records_by_date(
        records: list[Record], start_int: int, end_int: int
    ) -> list[Record]:
        start = number_to_string(start_int)
        end = number_to_string(end_int)

        filtered_records = filter(
            lambda item: (start <= item.get_day() and end >= item.get_day()),
            records,
        )

        return list(filtered_records)

    @staticmethod
    def summarize_records(records: list[Record]) -> dict:
        summary = {}

        for record in records:
            if record.category not in summary:
                summary[record.category] = {}

            if (
                record.subject is not None
                and record.subject not in summary[record.category]
            ):
                summary[record.category][record.subject] = {"hour": 0, "minute": 0}

            if (
                record.subject is not None
                and summary[record.category][record.subject] is not None
                and record.time_spent is not None
            ):
                summary[record.category][record.subject][
                    "hour"
                ] += record.time_spent.hour
                summary[record.category][record.subject][
                    "minute"
                ] += record.time_spent.minute

                if summary[record.category][record.subject]["minute"] >= 60:
                    hour = int(summary[record.category][record.subject]["minute"] / 60)
                    summary[record.category][record.subject]["hour"] += hour
                    summary[record.category][record.subject]["minute"] = (
                        summary[record.category][record.subject]["minute"] % 60
                    )

        return summary

    @staticmethod
    def sort_summary(summary: dict) -> dict:
        sorted_summary = Utils.new_dictionary_with_defualt_categories()

        for category in summary:
            if category not in sorted_summary:
                sorted_summary[category] = {}

            subjects = summary[category]
            sorted_summary[category] = subjects

        return sorted_summary

    @staticmethod
    def new_dictionary_with_defualt_categories() -> dict:
        dictionary = {}

        for category in Keys.CATEGORIES:
            dictionary[category] = {}

        return dictionary

    @staticmethod
    def extract_logs(records: list[Record]):
        logs: dict[str, Log] = {}
        log_last_id = 0

        for record in records:
            if record.date not in logs:
                log_last_id += 1
                logs[record.date] = Log(id=log_last_id, date=record.date)

            if record.category == "ورود":
                logs[record.date].clock_in = record.start_time

            if record.category == "خروج":
                logs[record.date].clock_out = record.start_time

        return list(logs.values())


def number_to_string(number: int) -> str:
    string = str(number)

    if len(string) == 1:
        string = f"0{string}"

    return string
