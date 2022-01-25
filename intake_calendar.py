#!/usr/bin/env python3

import logging
from itertools import cycle
from re import match

import smartsheet

from utils import COLOR_INDEX, clear_and_write_sheet, column_name_to_id_map, get_cell_by_column_name, \
    replace_event_names

logger = logging.getLogger('main')

smart = smartsheet.Smartsheet()  # use 'SMARTSHEET_ACCESS_TOKEN' env variable
smart.errors_as_exceptions(True)
CHANGE_AGENT = "dkarpele_smartsheet_calendar"
smart.with_change_agent(CHANGE_AGENT)


def intake_processing(intake_sheet_id: int):
    new_cells = []
    color_cycle = cycle(COLOR_INDEX)
    sheet = smart.Sheets.get_sheet(intake_sheet_id)
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

        for date_col in date_cols:
            item = replace_event_names(date_col.title)
            name = f"{event} | {item}"
            date = row.get_column(date_col.id).value
            new_cells.append((name, date or "", "", color))

    return new_cells


def process_sheet(sheet_ids):
    clear_and_write_sheet(smart, sheet_ids['destination'], intake_processing(sheet_ids['source']))
