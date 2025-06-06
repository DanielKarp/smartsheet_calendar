import logging

import calendar_control

fmt_str = "%(levelname)s:%(asctime)s::%(module)s:%(funcName)s: %(message)s"
formatter = logging.Formatter(fmt_str)

logger = logging.getLogger('main')
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
s_logger.addHandler(s_file_handler)
s_logger.setLevel(logging.INFO)


def run():
    calendar_control.run()


if __name__ == '__main__':
    run()
