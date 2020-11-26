import openpyxl
import datetime

def str_to_dt(date_str):
    FORMAT = "%Y%m%d"
    dt = datetime.datetime.strptime(date_str, FORMAT)
    return dt


def is_lt_30days_from_now(dt):
    time_span = datetime.timedelta(days=30)
    now = datetime.datetime.now()
    return now - dt <= time_span


def find_rows_30days_from_now(excel_filename):
    if not excel_filename.endswith(".xlsx"):
        raise ValueError("invalid excel file")
    
    DEADLINE_COL_INDEX = 19
    source_file = openpyxl.load_workbook(excel_filename, read_only=True)
    sheet = source_file["Sheet1"]

    rows_to_notice = []
    selection = sheet.values
    next(selection) # skip title row

    for row in selection:
        if is_lt_30days_from_now(str_to_dt(str(row[DEADLINE_COL_INDEX]))):
            rows_to_notice.append(row)
    return rows_to_notice