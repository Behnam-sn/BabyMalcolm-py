from item import Item
import utils
import keys


def generate_daily_report(items: list[Item]):
    daily_report = utils.item_to_dictionary(items)
    utils.dictionary_to_json_file(daily_report, keys.DAILY_REPORT_FILE_NAME)


def generate_weekly_report(items: list[Item], start: int, end: int):
    filtered_items = utils.filter_items_by_date(items, start, end)
    weekly_report = utils.summarize_items(filtered_items)
    sorted_weekly_report = utils.sort_summary(weekly_report)
    utils.dictionary_to_json_file(sorted_weekly_report, keys.WEEKLY_REPORT_FILE_NAME)
    utils.dictionary_to_text_file(sorted_weekly_report, keys.WEEKLY_REPORT_FILE_NAME)


def generate_monthly_report(items: list[Item]):
    monthly_report = utils.summarize_items(items)
    sorted_monthly_report = utils.sort_summary(monthly_report)
    utils.dictionary_to_json_file(sorted_monthly_report, keys.MONTHLY_REPORT_FILE_NAME)
    utils.dictionary_to_text_file(sorted_monthly_report, keys.MONTHLY_REPORT_FILE_NAME)
