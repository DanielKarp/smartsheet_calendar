import logging

import smartsheet

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


def replace_event_names(event: str) -> str:
    replacements = [
        ("Cisco Live", "CL"),
        ("Cisco ", ""),
        ("Partner Summit", "PS"),
        ("Date", ""),
    ]
    for original, new in replacements:
        if original in event:
            new_event = event.replace(original, new)
            logger.debug(
                f'found "{original}" in {event}, replaced with "{new}", result is {new_event}'
            )
            event = new_event
    return event


def clear_rows(smart: smartsheet.Smartsheet, sheet: smartsheet.models.Sheet) -> None:
    if rows := [row.id for row in sheet.rows]:
        smart.Sheets.delete_rows(sheet_id=sheet.id, ids=rows)
        logger.info(f"cleared {len(rows)} rows")
    else:
        logger.warning("no rows cleared")


def write_rows(smart: smartsheet.Smartsheet, sheet: smartsheet.models.Sheet, rows: list) -> None:
    new_rows = []
    name_col, date_col, end_date_col = [col.id for col in sheet.columns[:3]]
    for name, date, end_date, color in rows:
        if name and date:
            new_row = smartsheet.models.Row()
            new_row.to_bottom = True
            new_row.format = f",,,,,,,,,{color},{color},,,,,,"

            new_cell1 = smartsheet.models.Cell()
            new_cell1.value = name
            new_cell1.column_id = name_col

            new_cell2 = smartsheet.models.Cell()
            new_cell2.strict = False
            new_cell2.value = date
            new_cell2.column_id = date_col

            new_cell3 = smartsheet.models.Cell()
            new_cell3.strict = False
            new_cell3.value = end_date
            new_cell3.column_id = end_date_col

            new_row.cells = [new_cell1, new_cell2, new_cell3]
            new_rows.append(new_row)
    if new_rows:
        smart.Sheets.add_rows(sheet.id, new_rows)
        logger.info(f"wrote {len(new_rows)} rows")
    else:
        logger.warning("no rows written")


def get_cell_by_column_name(
    row: smartsheet.models.Row, column_name: str, col_map: dict
) -> smartsheet.models.Cell:
    return row.get_column(col_map[column_name])  # {NAME: ID}


def column_name_to_id_map(columns: list) -> dict:
    # returns a title:id dict of all columns in sheet
    return {column.title: column.id for column in columns}
