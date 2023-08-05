import processor
import utils
import keys


if __name__ == "__main__":
    items = utils.excel_to_items(keys.EXCEL_FILE_PATH)

    processor.generate_daily_report(items)
    processor.generate_weekly_report(items=items, start=24, end=28)
    processor.generate_monthly_report(items)
