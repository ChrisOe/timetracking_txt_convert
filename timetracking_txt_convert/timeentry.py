from datetime import date, datetime, timedelta


class TimeEntry:

    def __init__(
            self,
            time_date: date,
            time_type: str,
            start: datetime,
            end: datetime,
            duration: timedelta,
            comment: str
    ):
        self.time_date = time_date
        self.time_type = time_type
        self.start = start
        self.end = end
        self.duration = duration
        self.comment = comment
