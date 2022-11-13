import timeentry
from datetime import date, datetime, timedelta


class Employee:

    def __init__(self, first_name: str, last_name: str, personal_no: str, department: str, report_month: int):
        self.first_name = first_name
        self.last_name = last_name
        self.full_name = f"{last_name}, {first_name}"
        self.personal_no = personal_no
        self.department = department
        self.report_month = report_month
        self.report_year = datetime.now().year
        self.is_full_month = False
        self.paid_per_hour = False
        self.three_month_mean = timedelta(seconds=0, days=0)
        self.holidays = 0
        self.days_vacation = 0
        self.days_unpaid_vacation = 0
        self.days_paid_sick = 0
        self.days_unpaid_sick = 0
        self.days_else = 0
        self.days = {}

    def add_time(
            self,
            time_date: date,
            time_type: str,
            start: datetime,
            end: datetime,
            duration: timedelta,
            comment: str
    ):
        # adds a new entry to the list of times
        self.paid_per_hour = True
        time_entry = timeentry.TimeEntry(time_date, time_type, start, end, duration, comment)
        seen = set(self.days)
        if time_date not in seen:
            self.days[time_date] = {
                "times": [time_entry],
                "sum": timedelta()
            }
        else:
            self.days[time_date]["times"].append(time_entry)

    def get_monthly_sum(self) -> timedelta:
        # summarizes all day sums to one month sum
        month_sum = timedelta()
        for day in self.days:
            month_sum += self.days[day]['sum']
        return month_sum

    def get_avg_sum(self) -> int:
        # average hours an employee worked, saved in seconds.
        sums = [self.days[key]['sum'].seconds + self.days[key]['sum'].days * 86400 for key in self.days]
        if len(sums) == 0:
            return 0
        else:
            return int(sum(sums) / len(sums))

    def get_sum_paid_extra_days(self):
        return self.holidays + self.days_vacation + self.days_paid_sick
