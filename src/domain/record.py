from dataclasses import dataclass
import datetime


@dataclass
class Record:
    id: int
    date: str
    day_in_week: str
    start_time: datetime.time
    category: str
    subject: str | None
    detail: str | None
    time_spent: datetime.time | None

    def convert_to_dictionary(self):
        return {
            "تاریخ": self.date,
            "روز هفته": self.day_in_week,
            "ساعت شروع": time_to_string(self.start_time),
            "دسته بندی": self.category,
            "موضوع": self.subject,
            "جزئیات": self.detail,
            "مدت زمان": time_to_string(self.time_spent),
        }

    def get_day(self):
        return self.date[len(self.date) - 2 :]


def time_to_string(time: datetime.time | None) -> str | None:
    if time == None:
        return None
    return time.strftime("%H:%M")
