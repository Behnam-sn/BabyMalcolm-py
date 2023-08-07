import json
from os import mkdir

from openpyxl import load_workbook

from domain.record import Record
from keys import Keys


class Utils:
    @staticmethod
    def extract_records_from_excel(excel_file_path: str) -> list[Record]:
        records = []
        record_last_id = 0

        workbook = load_workbook(
            data_only=True,
            filename=excel_file_path,
        )
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=2, values_only=True):  # type: ignore
            record_last_id += 1

            record = Record(
                id=record_last_id,
                date=row[Keys.ITEM_DATE],
                day_in_week=row[Keys.ITEM_DAY_IN_WEEK],
                start_time=row[Keys.ITEM_START_TIME],
                category=row[Keys.ITEM_CATEGORY],
                subject=row[Keys.ITEM_SUBJECT],
                detail=row[Keys.ITEM_DETAIL],
                time_spent=row[Keys.ITEM_TIME_SPENT],
            )

            records.append(record)

        return records

    @staticmethod
    def create_folder(path: str, name: str) -> None:
        try:
            mkdir(f"{path}\\{name}")

        except Exception:
            pass

    @staticmethod
    def list_to_json_file(summary: list, file_name: str) -> None:
        Utils.create_folder(path=".", name="dist")

        with open(f"dist\\{file_name}.json", "w", encoding="utf-8") as file:
            json.dump(summary, file, ensure_ascii=False, indent=4)

    @staticmethod
    def list_to_text_file(summary: list, file_name: str) -> None:
        Utils.create_folder(path=".\\", name="dist")

        text_file = open(f"dist\\{file_name}.txt", "w", encoding="utf-8")

        for item in summary:
            line = ""

            for key in item:
                line += f"{item[key]}\t"

            text_file.write(f"{line}\n")

    @staticmethod
    def dictionary_to_json_file(summary: dict, file_name: str) -> None:
        Utils.create_folder(path=".", name="dist")

        with open(f"dist\\{file_name}.json", "w", encoding="utf-8") as file:
            json.dump(summary, file, ensure_ascii=False, indent=4)

    @staticmethod
    def dictionary_to_text_file(summary: dict, file_name: str) -> None:
        Utils.create_folder(path=".\\", name="dist")

        text_file = open(f"dist\\{file_name}.txt", "w", encoding="utf-8")

        for category in summary:
            for subject in summary[category]:
                hour = summary[category][subject]["hour"]
                minute = summary[category][subject]["minute"]

                stringified_hour = number_to_string(hour)
                stringified_minute = number_to_string(minute)

                time_spent = f"{stringified_hour}:{stringified_minute}"

                text_file.write(f"{category}\t{subject}\t{time_spent}\n")


def number_to_string(number: int) -> str:
    string = str(number)

    if len(string) == 1:
        string = f"0{string}"

    return string
