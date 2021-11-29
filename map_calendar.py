#!/usr/bin/env python3

import logging
from itertools import cycle
from re import match

import smartsheet

from utils import COLOR_INDEX, clear_rows, column_name_to_id_map, get_cell_by_column_name, write_rows

MAP_SHEET = 6446980407814020
CALENDAR_SHEET = 6041912604944260

fmt_str = "%(levelname)s:%(asctime)s:%(name)s: %(message)s"
formatter = logging.Formatter(fmt_str)

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

d_file_handler = logging.FileHandler("api_and_calendar.log")
d_file_handler.setLevel(logging.DEBUG)
d_file_handler.setFormatter(formatter)

i_file_handler = logging.FileHandler("calendar.log")
i_file_handler.setLevel(logging.INFO)
i_file_handler.setFormatter(formatter)

stream_handler = logging.StreamHandler()
stream_handler.setLevel(logging.INFO)
stream_handler.setFormatter(formatter)

logger.addHandler(i_file_handler)
logger.addHandler(d_file_handler)
logger.addHandler(stream_handler)

s_file_handler = logging.FileHandler("api_and_calendar.log")
s_file_handler.setLevel(logging.INFO)
s_file_handler.setFormatter(formatter)

s_logger = logging.getLogger("smartsheet.smartsheet")
s_logger.addHandler(d_file_handler)
s_logger.setLevel(logging.DEBUG)

smart = smartsheet.Smartsheet()  # use 'SMARTSHEET_ACCESS_TOKEN' env variable
smart.errors_as_exceptions(True)
CHANGE_AGENT = "dkarpele_smartsheet_calendar"
smart.with_change_agent(CHANGE_AGENT)


def process_sheet():
    new_cells = []
    color_cycle = cycle(COLOR_INDEX)
    sheet = smart.Sheets.get_sheet(MAP_SHEET)
    rows = sheet.rows
    columns = sheet.columns
    col_map = column_name_to_id_map(columns=columns)
    date_cols = [
        col
        for col in columns
        if col.type == "DATE"
        and not col.hidden
        and col.title != "Event Start Date"
        and col.title != "Event End Date"
    ]
    logger.debug(f"found {len(columns)} total columns")
    logger.debug(f"found {len(date_cols)} date-type columns")
    logger.debug(f"found {len(rows)} rows")
    for row in rows:
        event = get_cell_by_column_name(row, "Show Name", col_map).display_value
        # if the row matches any of the below, it should not be added to the calendar
        if match(r"^Q[1-4]", str(event)) or not row.parent_id or event is None:
            logger.debug(f"{event} was identified as a non-event row")
            continue
        if get_cell_by_column_name(row, "TechX Status", col_map).value != 'Green':
            logger.debug(f"{event} was identified as an unconfirmed event")
            continue
        if not (staff := get_cell_by_column_name(row, "TechX Resource", col_map).value):
            logger.debug(f"{event} was identified as an event without anyone assigned")
            continue

        logger.debug(f"{event} is being processed")
        color = next(color_cycle)  # each event gets its own color
        staff = staff.strip('"')
        event = f'{event} | {staff}'

        start_col = next(col for col in columns if col.title == "Event Start Date")
        end_col = next(col for col in columns if col.title == "Event End Date")
        start_date = row.get_column(start_col.id).value
        end_date = row.get_column(end_col.id).value
        new_cells.append((event, start_date or "", end_date or "", color))

        for date_col in date_cols:
            item = date_col.title
            name = f"{event} | {item}"
            date = row.get_column(date_col.id).value
            new_cells.append((name, date or "", "", color))

    cal_sheet = smart.Sheets.get_sheet(CALENDAR_SHEET)
    clear_rows(smart, cal_sheet)
    write_rows(smart, cal_sheet, new_cells)


if __name__ == "__main__":
    logger.info("starting map calendar program")
    process_sheet()
    logger.info("map calendar program finished\n")
