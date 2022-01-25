#!/usr/bin/env python3

import logging

import smartsheet

from intake_calendar import intake_processing
from map_calendar import map_processing
from utils import clear_and_write_sheet

logger = logging.getLogger('main')

smart = smartsheet.Smartsheet()  # use 'SMARTSHEET_ACCESS_TOKEN' env variable
smart.errors_as_exceptions(True)
CHANGE_AGENT = "dkarpele_smartsheet_calendar"
smart.with_change_agent(CHANGE_AGENT)


def process_sheet(sheet_ids):
    new_cells = map_processing(sheet_ids['source']['map'])
    new_cells.extend(intake_processing(sheet_ids['source']['intake']))
    clear_and_write_sheet(smart, sheet_ids['destination'], new_cells)


if __name__ == "__main__":
    logger.info("starting combined calendar program")
    import yaml
    with open('sheet_id.yaml') as yaml_file:
        sheet_id = yaml.safe_load(yaml_file)
    process_sheet(sheet_id['combined'])
    logger.info("combined calendar program finished\n")
