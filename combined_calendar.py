#!/usr/bin/env python3

import logging
from itertools import cycle
from re import match

import smartsheet

from utils import (COLOR_INDEX, clear_rows, column_name_to_id_map, get_cell_by_column_name, write_rows,
                   replace_event_names)

MAP_SHEET = 6446980407814020
INTAKE_FORM_SHEET = 3901696217769860
CALENDAR_SHEET = 8959263235172228

fmt_str = "%(levelname)s:%(asctime)s:%(name)s: %(message)s"
formatter = logging.Formatter(fmt_str)

logger = logging.getLogger()
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


def process_map_sheet():
    new_cells = []
    color_cycle = cycle(COLOR_INDEX)
    sheet = smart.Sheets.get_sheet(MAP_SHEET)
    rows = sheet.rows
    columns = sheet.columns
    col_map = column_name_to_id_map(columns=columns)
    logger.debug(f"found {len(columns)} total columns")
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

        logger.debug(f"{event} is being processed")
        color = next(color_cycle)  # each event gets its own color

        if staff := get_cell_by_column_name(row, "TechX Resource", col_map).value:
            staff = staff.strip('"')
            event = f'{event} | {staff}'
        else:
            logger.debug(f"{event} was identified as an event without anyone assigned")

        start_col = next(col for col in columns if col.title == "Event Start Date")
        end_col = next(col for col in columns if col.title == "Event End Date")
        start_date = row.get_column(start_col.id).value
        end_date = row.get_column(end_col.id).value
        new_cells.append((event, start_date or "", end_date or "", color))

    return new_cells


def process_intake_sheet():
    new_cells = []
    color_cycle = cycle(COLOR_INDEX)
    sheet = smart.Sheets.get_sheet(INTAKE_FORM_SHEET)
    rows = sheet.rows
    columns = sheet.columns
    col_map = column_name_to_id_map(columns=columns)
    logger.debug(f"found {len(columns)} total columns")
    logger.debug(f"found {len(rows)} rows")
    for row in rows:
        event = get_cell_by_column_name(row, "Event Name", col_map).value
        # if the row matches, it is a label row. Contains no data so it's skipped
        if match(r"^Q[1-4] FY\d{2}", event):
            logger.debug(f"{event} was identified as a separator row")
            continue
        event_state = get_cell_by_column_name(row, "Event State & Type", col_map).value
        if event_state is not None and "Canceled" in event_state:
            logger.debug(f"{event} was identified as a canceled event")
            event = f'(Canceled) {event}'
        logger.debug(f"{event} is being processed")
        color = next(color_cycle)  # each event gets its own color
        event = replace_event_names(event)  # do some filtering to shorten some words

        start_col = next(col for col in columns if col.title == "Event Start Date")
        end_col = next(col for col in columns if col.title == "Event End Date")
        start_date = row.get_column(start_col.id).value
        end_date = row.get_column(end_col.id).value
        new_cells.append((event, start_date or "", end_date or "", color))

    return new_cells


def process_sheet():
    new_cells = process_map_sheet()
    new_cells.extend(process_intake_sheet())
    cal_sheet = smart.Sheets.get_sheet(CALENDAR_SHEET)
    clear_rows(smart, cal_sheet)
    write_rows(smart, cal_sheet, new_cells)


if __name__ == "__main__":
    logger.info("starting combined calendar program")
    process_sheet()
    logger.info("combined calendar program finished\n")