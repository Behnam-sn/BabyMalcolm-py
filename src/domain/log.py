import datetime


class Log:
    clock_in: datetime.time
    clock_out: datetime.time

    def __init__(self, id: int, date: str):
        self.id = id
        self.date = date

    def convert_to_dictionary(self):
        return {
            "تاریخ": self.date,
            "ساعت ورود": time_to_string(self.clock_in),
            "ساعت خروج": time_to_string(self.clock_out),
        }


def time_to_string(time: datetime.time | None) -> str | None:
    if time == None:
        return None
    return time.strftime("%H:%M")
