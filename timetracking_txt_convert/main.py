from datetime import date, datetime, timedelta
import os
import json
from employee import Employee
import workbook
import secrets

WEEKDAYS = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
MONTH_NAMES = {
    1: "Januar",
    2: "Februar",
    3: "März",
    4: "April",
    5: "Mai",
    6: "Juni",
    7: "Juli",
    8: "August",
    9: "September",
    10: "Oktober",
    11: "November",
    12: "Dezember",
}
HOLIDAYS = [
    "Neujahr",
    "Karfreitag",
    "Ostermontag",
    "Maifeiertag",
    "Himmelfahrt",
    "Pfingstmontag",
    "Tag",
    "Reformationstag",
    "Heiligabend",
    "1.",
    "2.",
    "Silvester",
]
VACATION = ["URLAUB", "Unbez."]
SICK = ["KRANK"]
# This is a dictionary with employee names that should be skipped by default.
# Mainly for test accounts that are not relevant.
SKIP_NAMES = secrets.SKIP_NAMES


def get_datetime_from_str(time_str: str, date_obj: date) -> datetime:
    # Get a string in the form of "#0:#0" and form a datetime object.
    hour = int(time_str.split(":")[0])
    minute = int(time_str.split(":")[1])
    datetime_obj = datetime(year=date_obj.year, month=date_obj.month, day=date_obj.day, hour=hour, minute=minute)
    return datetime_obj


def get_timedelta_from_str(time_str: str) -> timedelta:
    # Get a string in the form of "#0:#0" and form a timedelta object.
    hour = int(time_str.split(":")[0])
    minute = int(time_str.split(":")[1])
    return timedelta(hours=hour, minutes=minute)


def format_title_date(source_date: str) -> date:
    # Get a string in the form of "dd.mm.yy" and form a date object. The century has to be provided.
    # Year 2000 bug anyone? -.-
    format_date = [int(j) for j in source_date.split(".")]
    format_year = int(f"20{format_date[2]}")
    end_date = date(day=format_date[0], month=format_date[1], year=format_year)
    return end_date


def format_line_date(source_date: str, report_from_date: date, report_to_date: date) -> date:
    # Get a string in the form of "dd.mm." and form a date object. The year has to come from somewhere else.
    format_date = f"{source_date}".split(".")
    if int(format_date[1]) > report_to_date.month:
        format_year = report_from_date.year
    else:
        format_year = report_to_date.year
    return date(year=format_year, month=int(format_date[1]), day=int(format_date[0]))


def str_from_seconds(delta_seconds: int, delta_days: int) -> str:
    # forms a readable time from seconds
    delta_seconds += delta_days * 86400
    delta_hours = delta_seconds // 3600
    delta_seconds = delta_seconds - (delta_hours * 3600)
    delta_minutes = delta_seconds // 60
    return f"{delta_hours:02d}:{delta_minutes:02d}"


def guess_year(guess_aid_month: int) -> int:
    # Todo: find a way to get the relevant year that is more precise
    current_date = datetime.now()
    if guess_aid_month <= current_date.month:
        guessed_year = current_date.year
    else:
        guessed_year = current_date.year - 1
    return guessed_year


def check_for_dir(directory: str):
    # check if directory exists and if not, create one
    if not os.path.isdir(f"./{directory}"):
        os.mkdir(f"./{directory}")


def last_day_of_month(init_date: date) -> date:
    next_month = init_date.replace(day=28) + timedelta(days=4)
    return next_month - timedelta(days=next_month.day)


def get_three_month_mean(to_month: int, to_year: int, means: dict):
    temp_date = date(year=to_year, month=to_month, day=1) - timedelta(days=40)
    from_day = date(year=temp_date.year, month=temp_date.month, day=1)
    to_day = date(year=to_year, month=to_month, day=20)
    temp_means = []
    for day, time in means.items():
        dict_date = date(year=int(day.split("-")[0]), month=int(day.split("-")[1]), day=15)
        if from_day <= dict_date <= to_day:
            temp_means.append(int(time))
    if len(temp_means) == 0:
        return timedelta(seconds=0)
    else:
        result = int(sum(temp_means) / len(temp_means))
        return timedelta(seconds=result)


check_for_dir("data")
check_for_dir("input")
check_for_dir("output")

try:
    with open("data/employee_mean.json", mode="r") as file:
        employee_mean = json.load(file)
except FileNotFoundError:
    employee_mean = {}

time_data = {}
employee_times = {}
employees = []
companies = []
pre_valid_month = {}
full_name = ""
error_message = ""
process_employee = True
valid_file = True
from_date = date(day=1, month=1, year=2000)
to_date = last_day_of_month(from_date)

input_file_list = os.listdir('input/')

for file_name in input_file_list:
    if file_name[-3:] == "txt":
        valid_month = []
        print(f"Einlesen der Daten von {file_name}.")
        with open(f"input/{file_name}", mode="r", encoding='iso-8859-1') as file:
            lines = file.readlines()

        spaced_lines = [lines[index].replace("-", " ").replace("\t", " ").split() for index in range(len(lines))]

        for i in range(len(spaced_lines)):
            line = spaced_lines[i]
            if i == 3:
                from_date = format_title_date(line[3])
                to_date = format_title_date(line[4])

            if len(line) != 0:
                if line[0] in WEEKDAYS:
                    line_date = format_line_date(
                        source_date=line[1],
                        report_from_date=from_date,
                        report_to_date=to_date,
                    )
                    if line_date.month not in pre_valid_month:
                        pre_valid_month[line_date.month] = {
                            "first_day": False,
                            "last_day": False,
                        }
                    if line_date.day == 1:
                        pre_valid_month[line_date.month]["first_day"] = True
                    if line_date == last_day_of_month(line_date):
                        pre_valid_month[line_date.month]["last_day"] = True

        for key in pre_valid_month:
            if pre_valid_month[key]["first_day"] and pre_valid_month[key]["last_day"]:
                valid_month.append(key)

        for index in range(len(spaced_lines)):
            spaced_line = spaced_lines[index]
            if index == 3:
                from_date = format_title_date(spaced_line[3])
                to_date = format_title_date(spaced_line[4])

            if len(spaced_line) > 0:
                if spaced_line[0] == "Name:":
                    name_line = lines[index].replace(",", " ").replace("\t", " ").split()
                    l_name = name_line[1]
                    process_employee = l_name not in SKIP_NAMES
                    if process_employee:
                        # check for two first names
                        if len(name_line) == 9:
                            f_name = f"{name_line[2]} {name_line[3]}"
                            personal_no = name_line[5]
                            department = name_line[7]
                        else:
                            f_name = name_line[2]
                            personal_no = name_line[4]
                            department = name_line[6]
                        report_date = format_line_date(
                            source_date=spaced_lines[index + 4][1],
                            report_from_date=from_date,
                            report_to_date=to_date,
                        )
                        employee = Employee(
                            first_name=f_name,
                            last_name=l_name,
                            personal_no=personal_no,
                            department=department,
                            report_month=report_date.month,
                        )
                        employee.report_year = report_date.year
                        full_name = employee.full_name
                        month_year = f"{report_date.year}-{report_date.month}"
                        if month_year not in time_data:
                            time_data[month_year] = {
                                department: {full_name: employee}
                            }
                        else:
                            if department not in time_data[month_year]:
                                time_data[month_year][department] = {full_name: employee}
                            else:
                                time_data[month_year][department][full_name] = employee

                elif spaced_line[0] in WEEKDAYS and process_employee:
                    line_date = format_line_date(
                        source_date=spaced_line[1],
                        report_from_date=from_date,
                        report_to_date=to_date,
                    )
                    line_type = spaced_line[2]

                    line_sum = lines[index].split(" ")
                    line_sum_split = line_sum[-1].split()

                    if line_date.month in valid_month:
                        time_data[month_year][department][full_name].is_full_month = True
                    if line_date >= date.today():
                        time_data[month_year][department][full_name].is_full_month = False
                    # If normal workday, search times
                    if line_type == "ARBZ" or line_type == "PAUS":
                        line_times = spaced_line[2:-len(line_sum_split)]
                        block = []
                        for i in range(len(line_times)):
                            if len(block) > 2 and \
                                    (line_times[i] == "ARBZ" or
                                     line_times[i] == "PAUS" or
                                     i == len(line_times)-1):
                                if i == len(line_times)-1:
                                    block.append(line_times[i])
                                if ":" in block[1]:
                                    line_type = block[0]
                                    if len(block) == 5 or len(block) == 6:
                                        start = get_datetime_from_str(block[3], line_date)
                                        end = get_datetime_from_str(block[4], line_date)
                                    else:
                                        start = get_datetime_from_str(block[1], line_date)
                                        end = get_datetime_from_str(block[2], line_date)
                                    duration = end - start
                                    comment = ""
                                else:
                                    duration = get_timedelta_from_str(block[-1])
                                    line_type = "SPEZ"
                                    start = datetime(
                                        day=line_date.day,
                                        month=line_date.month,
                                        year=line_date.month,
                                        hour=7,
                                        minute=0
                                    )
                                    end = start + duration
                                    comment = " ".join(block[1:-1])
                                time_data[month_year][department][full_name].add_time(
                                    time_date=line_date,
                                    time_type=line_type,
                                    start=start,
                                    end=end,
                                    duration=duration,
                                    comment=comment,
                                )
                                block = []
                            block.append(line_times[i])
                        day_sum = get_timedelta_from_str(time_str=spaced_line[len(spaced_line) - len(line_sum_split)])
                        try:
                            time_data[month_year][department][full_name].days[line_date]['sum'] = day_sum
                        except KeyError:
                            print(line_date)
                            print(spaced_line)
                            print(line_sum_split)
                            print(time_data[month_year][department][full_name].days)
                            break
                    # If special day, then increase the relevant value
                    elif spaced_line[0] != "So" and full_name != "":
                        if spaced_line[2] in HOLIDAYS:
                            time_data[month_year][department][full_name].holidays += 1
                        elif spaced_line[2] in VACATION:
                            if spaced_line[2] == "URLAUB":
                                time_data[month_year][department][full_name].days_vacation += 1
                            else:
                                time_data[month_year][department][full_name].days_unpaid_vacation += 1
                        elif spaced_line[2] in SICK:
                            if spaced_line[3] == "MIT":
                                time_data[month_year][department][full_name].days_paid_sick += 1
                            else:
                                time_data[month_year][department][full_name].days_unpaid_vacation += 1
                        else:
                            try:
                                time_data[month_year][department][full_name].days_else += 1
                            except KeyError:
                                print(f"{index}|{spaced_lines[index]}")

# Calculate employee mean workhours per day
for work_month_year in time_data:
    for department in time_data[work_month_year]:
        for employee in time_data[work_month_year][department]:
            employee_obj = time_data[work_month_year][department][employee]
            if employee_obj.is_full_month:
                mean = employee_obj.get_avg_sum()
                if mean > 0:
                    full_name = employee_obj.full_name
                    if full_name not in employee_mean:
                        employee_mean[full_name] = {}
                    month_year = f"{employee_obj.report_year}-{employee_obj.report_month}"
                    employee_mean[full_name][month_year] = employee_obj.get_avg_sum()

with open("data/employee_mean.json", mode="w", encoding="UTF-8") as json_file:
    json.dump(employee_mean, json_file, indent=4)

for work_month_year in time_data:
    for department in time_data[work_month_year]:
        for key, employee in time_data[work_month_year][department].items():
            if key in employee_mean:
                report_month = int(employee.report_month)
                report_year = int(employee.report_year)
                employee.three_month_mean = get_three_month_mean(report_month, report_year, employee_mean[key])

# Export data to excel workbook
for work_month_year in time_data:
    for department in time_data[work_month_year]:
        work_year = work_month_year.split("-")[0]
        work_month = work_month_year.split("-")[1]
        wb = workbook.Workbook(month=work_month, year=work_year, department=department)
        for employee in time_data[work_month_year][department]:
            wb.add_employee_sheet(time_data[work_month_year][department][employee])
        error_message += wb.save_workbook()

if len(error_message) > 0:
    print(error_message)
    # input("Hit Enter to exit.")
