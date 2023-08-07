from application.application import Application


if __name__ == "__main__":
    Application.generate_time_report()
    Application.generate_daily_report()
    Application.generate_weekly_report(start=24, end=28)
    Application.generate_monthly_report()
