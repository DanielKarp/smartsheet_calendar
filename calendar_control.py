import logging

import yaml

logger = logging.getLogger('main')


def run() -> None:
    logger.info("starting control program")
    from combined_calendar import process_sheet as process_combined
    from intake_calendar import process_sheet as process_intake
    from map_calendar import process_sheet as process_map
    from request_to_map.request_to_map_calendar import process_sheet as process_request

    with open('sheet_id.yaml') as yaml_file:
        sheet_ids = yaml.safe_load(yaml_file)

    for process_cal in (
                    #   process_request,
                        process_intake,
                        process_map,
                        process_combined,
                        ):
        module = process_cal.__module__.split('.')[-1].replace('_', ' ')
        logger.info(f"starting {module} processing")
        process_cal(sheet_ids[' '.join(module.split()[:-1])])

    logger.info("program finished")
