from re import match
from itertools import cycle

import smartsheet


smart = smartsheet.Smartsheet()  # use 'SMARTSHEET_ACCESS_TOKEN' env variable
smart.errors_as_exceptions(True)

CHANGE_AGENT = 'dkarpele_smartsheet_calendar'
smart.with_change_agent(CHANGE_AGENT)

INTAKE_FORM_SHEET = 7774724450215812
CALENDAR_SHEET = 4620060233885572

COLORS = ["none", "#000000", "#FFFFFF", "transparent",
          "#FFEBEE", "#FFF3DF", "#FFFEE6", "#E7F5E9", "#E2F2FE", "#F4E4F5", "#F2E8DE", "#FFCCD2", "#FFE1AF", "#FEFF85",
          "#C6E7C8", "#B9DDFC", "#EBC7EF", "#EEDCCA", "#E5E5E5", "#F87E7D", "#FFCD7A", "#FEFF00", "#7ED085", "#5FB3F9",
          "#D190DA", "#D0AF8F", "#BDBDBD", "#EA352E", "#FF8D00", "#FFED00", "#40B14B", "#1061C3", "#9210AD", "#974C00",
          "#757575", "#991310", "#EA5000", "#EBC700", "#237F2E", "#0B347D", "#61058B", "#592C00"]
COLOR_INDEX = list([i for i, _ in enumerate(COLORS)][4:])
color_cycle = cycle(COLOR_INDEX)


def process_sheet(sheet_id):
    new_cells = []
    col_map = column_name_to_id_map(sheet_id=sheet_id)
    rows = smart.Sheets.get_sheet(sheet_id).rows
    date_cols = [c for c in smart.Sheets.get_columns(sheet_id=sheet_id, include_all=True).data if c.type == 'DATE']

    print(f'found {len(rows)} rows')
    for row in rows:
        event = get_cell_by_column_name(row, "Event Name", col_map).value
        if match(r'^Q[1-4] FY\d{2}', event):  # if the row matches, it is a label row. Contains no data so it's skipped
            continue
        color = next(color_cycle)  # each event gets its own color
        event = replace_event_names(event)  # do some filtering to shorten some words

        for date_col in date_cols:
            name = f'{event} | {date_col.title}'
            date = row.get_column(date_col.id).value
            new_cells.append((name, date, color))

    clear_rows(CALENDAR_SHEET)
    write_rows(CALENDAR_SHEET, new_cells)


def replace_event_names(event: str) -> str:
    if 'Cisco Live' in event:
        event = event.replace('Cisco Live', 'CL')
    if 'Cisco' in event:
        event = event.replace('Cisco ', '')
    return event


def clear_rows(sheet_id: int):
    rows = [row.id for row in smart.Sheets.get_sheet(sheet_id).rows]
    if rows:
        smart.Sheets.delete_rows(sheet_id=sheet_id, ids=rows)
        print(f'cleared {len(rows)} rows')
    else:
        print('no rows cleared')


def write_rows(sheet_id: int,
               rows: list):
    new_rows = []
    col1, col2 = [col.id for col in smart.Sheets.get_columns(sheet_id=sheet_id, include_all=True).data][:2]
    for name, date, color in rows:
        if name and date:
            new_row = smartsheet.models.Row()
            new_row.to_bottom = True
            new_row.format = f",,,,,,,,,{color},{color},,,,,,"

            new_cell1 = smartsheet.models.Cell()
            new_cell1.value = name
            new_cell1.column_id = col1

            new_cell2 = smartsheet.models.Cell()
            new_cell2.strict = False
            new_cell2.value = date
            new_cell2.column_id = col2

            new_row.cells = [new_cell1, new_cell2]
            new_rows.append(new_row)
    if new_rows:
        smart.Sheets.add_rows(sheet_id, new_rows)
        print(f'wrote {len(new_rows)} rows')
    else:
        print('no rows written')


def get_cell_by_column_name(row: smartsheet.models.Row,
                            column_name: str,
                            col_map: dict) -> smartsheet.models.Cell:
    return row.get_column(col_map[column_name])  # {NAME: ID}


def column_name_to_id_map(sheet_id: int) -> dict:
    # returns a title:id dict of all columns in sheet
    return {column.title: column.id for column in smart.Sheets.get_columns(sheet_id, include_all=True).data}


if __name__ == '__main__':
    process_sheet(INTAKE_FORM_SHEET)
