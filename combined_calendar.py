#!/usr/bin/env python3

import logging

import smartsheet

from intake_calendar import intake_processing
from map_calendar import map_processing
from utils import clear_and_write_sheet

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
