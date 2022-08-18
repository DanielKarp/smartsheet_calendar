import logging

import smartsheet

logger = logging.getLogger('main')

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
                f'    found "{original}" in {event}, replaced with "{new}"; result is: {new_event}'
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


def clear_and_write_sheet(smart: smartsheet.Smartsheet, cal_sheet_id: int, new_cells: list) -> None:
    cal_sheet = smart.Sheets.get_sheet(cal_sheet_id)
    clear_rows(smart, cal_sheet)
    write_rows(smart, cal_sheet, new_cells)
