#!/usr/bin/env python3

import logging
from itertools import cycle
from re import match

import smartsheet

from utils import COLOR_INDEX, clear_and_write_sheet, column_name_to_id_map, get_cell_by_column_name

logger = logging.getLogger('main')

smart = smartsheet.Smartsheet()  # use 'SMARTSHEET_ACCESS_TOKEN' env variable
smart.errors_as_exceptions(True)
CHANGE_AGENT = "dkarpele_smartsheet_calendar"
smart.with_change_agent(CHANGE_AGENT)


def map_processing(map_sheet_id: int):
    new_cells = []
    color_cycle = cycle(COLOR_INDEX)
    sheet = smart.Sheets.get_sheet(map_sheet_id)
    rows = sheet.rows
    columns = sheet.columns
    col_map = column_name_to_id_map(columns=columns)
    logger.debug(f"found {len(columns)} total columns")
    logger.debug(f"found {len(rows)} rows")
    for row in rows:
        event = get_cell_by_column_name(row, "Event Name", col_map).display_value
        # if the row matches any of the below, it should not be added to the calendar
        if match(r"^Q[1-4]", str(event)) or not row.parent_id or event is None:
            logger.debug(f"{event} was identified as a non-event row")
            continue
        if get_cell_by_column_name(row, "TechX Status", col_map).value != 'Green':
            logger.debug(f"{event} was identified as an unconfirmed event")
            continue

        logger.debug(f"{event} is being processed")
        color = next(color_cycle)  # each event gets its own color

        just_event = event
        if staff := get_cell_by_column_name(row, "TechX Resource", col_map).value:
            staff = staff.strip('"')
            logger.debug(f"    {event} staff identified: {staff}")
            event = f'{event} | {staff}'
        else:
            logger.debug(f"    {event} was identified as an event without anyone assigned")

        start_date = get_cell_by_column_name(row, "Event Start Date", col_map).value
        end_date = get_cell_by_column_name(row, "Event End Date", col_map).value
        new_cells.append((event, start_date or "", end_date or "", color))

        jll_date = get_cell_by_column_name(row, "JLL Hand over date", col_map).value
        new_cells.append((just_event + ' | JLL Hand Over', jll_date or "", "", color))

        setup_date = get_cell_by_column_name(row, "Move In Date", col_map).value
        new_cells.append((just_event + ' | Setup Start', setup_date or "", "", color))

    return new_cells


def process_sheet(sheet_ids):
    clear_and_write_sheet(smart, sheet_ids['destination'], map_processing(sheet_ids['source']))
