import logging

from map_calendar import MAP_SHEET, process_sheet as process_map
from smartsheet_intake_calendar import INTAKE_FORM_SHEET, process_sheet as process_intake

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


logger.info("starting control program")

logger.info("starting intake calendar processing")
process_intake(INTAKE_FORM_SHEET)

logger.info("starting map calendar processing")
process_map(MAP_SHEET)

logger.info("program finished\n")
