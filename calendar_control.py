import logging

from combined_calendar import process_sheet as process_combined
from intake_calendar import process_sheet as process_intake
from map_calendar import process_sheet as process_map
from request_to_map.request_to_map_calendar import process_sheet as process_request

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


logger.info("starting control program")

for process_cal in (process_request, process_intake, process_map, process_combined):
    logger.info(f"starting {process_cal.__module__.split('.')[-1].replace('_', ' ')} processing")
    # process_cal()

logger.info("program finished")
