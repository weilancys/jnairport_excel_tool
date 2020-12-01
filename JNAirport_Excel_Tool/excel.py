import openpyxl
import datetime

def str_to_dt(date_str):
    FORMAT = "%Y%m%d"
    dt = datetime.datetime.strptime(date_str, FORMAT)
    return dt


def is_lt_30days_from_now(dt):
    time_span = datetime.timedelta(days=30)
    now = datetime.datetime.now()
    return now <= dt and dt - now <= time_span


def is_expired(dt):
    now = datetime.datetime.now()
    return dt < now


def find_noticeable_rows(excel_filename):
    if not excel_filename.endswith(".xlsx"):
        raise ValueError("invalid excel file")
    
    DEADLINE_COL_INDEX = 19
    source_file = openpyxl.load_workbook(excel_filename, read_only=True)
    sheet = source_file["Sheet1"]

    rows_lt_30_days = []
    rows_expired = []
    selection = sheet.values
    next(selection) # skip title row

    for row in selection:
        if is_lt_30days_from_now(str_to_dt(str(row[DEADLINE_COL_INDEX]))):
            rows_lt_30_days.append(row)
        if is_expired(str_to_dt(str(row[DEADLINE_COL_INDEX]))):
            rows_expired.append(row)
    return rows_lt_30_days, rows_expired