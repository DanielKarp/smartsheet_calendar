#!/usr/bin/env python3

import logging
from itertools import cycle
from re import match

import smartsheet

from utils import clear_rows, column_name_to_id_map, get_cell_by_column_name, replace_event_names, write_rows

INTAKE_FORM_SHEET = 3901696217769860
CALENDAR_SHEET = 4620060233885572

COLORS = [
    "none",
    "#000000",
    "#FFFFFF",
    "transparent",
    "#FFEBEE",
    "#FFF3DF",
    "#FFFEE6",
    "#E7F5E9",
    "#E2F2FE",
    "#F4E4F5",
    "#F2E8DE",
    "#FFCCD2",
    "#FFE1AF",
    "#FEFF85",
    "#C6E7C8",
    "#B9DDFC",
    "#EBC7EF",
    "#EEDCCA",
    "#E5E5E5",
    "#F87E7D",
    "#FFCD7A",
    "#FEFF00",
    "#7ED085",
    "#5FB3F9",
    "#D190DA",
    "#D0AF8F",
    "#BDBDBD",
    "#EA352E",
    "#FF8D00",
    "#FFED00",
    "#40B14B",
    "#1061C3",
    "#9210AD",
    "#974C00",
    "#757575",
    "#991310",
    "#EA5000",
    "#EBC700",
    "#237F2E",
    "#0B347D",
    "#61058B",
    "#592C00",
]
COLOR_INDEX = list([i for i, _ in enumerate(COLORS)][4:])
color_cycle = cycle(COLOR_INDEX)

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


smart = smartsheet.Smartsheet()  # use 'SMARTSHEET_ACCESS_TOKEN' env variable
smart.errors_as_exceptions(True)
CHANGE_AGENT = "dkarpele_smartsheet_calendar"
smart.with_change_agent(CHANGE_AGENT)


def process_sheet(sheet_id):
    new_cells = []
    sheet = smart.Sheets.get_sheet(sheet_id)
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
        event = get_cell_by_column_name(row, "Event Name", col_map).value
        # if the row matches, it is a label row. Contains no data so it's skipped
        if match(r"^Q[1-4] FY\d{2}", event):
            logger.debug(f"{event} was identified as a separator row")
            continue
        if event_state := get_cell_by_column_name(row, "Event State & Type", col_map).value:
            if "Canceled" in event_state:
                logger.debug(f"{event} was identified as a canceled event")
                event = f'(Canceled) {event}'
        logger.debug(f"{event} is being processed")
        color = next(color_cycle)  # each event gets its own color
        event = replace_event_names(event)  # do some filtering to shorten some words

        start_col = next(col for col in columns if col.title == "Event Start Date")
        end_col = next(col for col in columns if col.title == "Event End Date")
        start_date = row.get_column(start_col.id).value
        end_date = row.get_column(end_col.id).value
        new_cells.append((event, start_date, end_date, color))

        for date_col in date_cols:
            item = replace_event_names(date_col.title)
            name = f"{event} | {item}"
            date = row.get_column(date_col.id).value
            new_cells.append((name, date, "", color))

    cal_sheet = smart.Sheets.get_sheet(CALENDAR_SHEET)
    clear_rows(smart, cal_sheet)
    write_rows(smart, cal_sheet, new_cells)


if __name__ == "__main__":
    logger.info("starting program")
    process_sheet(INTAKE_FORM_SHEET)
    logger.info("program finished\n")
