import argparse
from datetime import datetime

import smartsheet

from FY_Q_sort import calc_fy_q_hardcoded
from colorize import colorize_rows

smart = smartsheet.Smartsheet()  # use 'SMARTSHEET_ACCESS_TOKEN' env variable
smart.errors_as_exceptions(True)

CHANGE_AGENT = 'dkarpele_smartsheet_calendar'
smart.with_change_agent(CHANGE_AGENT)

REQUEST_SHEET_ID = 4026093033285508
MAP_SHEET_ID = 1656024609384324


def v_print(*_, **__) -> None:
    pass


def process_sheet(request_sheet_id: int = REQUEST_SHEET_ID,
                  map_sheet_id: int = MAP_SHEET_ID,
                  simulate: bool = False) -> None:
    """Main loop for processing the sheet
    takes the sheet ids for the request sheet to pull rows from, and
    the map sheet to send rows to. An optional simulate option does not
    alter any sheets, but only shows the rows that would be copied.
    Gets rows from the request sheet, and builds {name: id} column map
    Prints out the column names and ids
    For each row in the request sheet, the TechX status column is checked
    if it is Yellow, the row is printed and if simulation is false, the
    row is sent to the map sheet and then the TechX Status column is
    changed to green.
    Does not return anything.
    """

    rows = smart.Sheets.get_sheet(request_sheet_id).rows
    request_column_mapping = column_name_to_id_map(request_sheet_id)
    map_column_mapping = column_name_to_id_map(map_sheet_id)

    print_col_headings(request_column_mapping)

    for row in rows:
        if check_row(row, request_column_mapping):
            v_print(f'  ^row will be processed')
            print_row(row, request_column_mapping)
            if not simulate:
                send_row(sheet_id=map_sheet_id,
                         row=row,
                         request_column_mapping=request_column_mapping,
                         map_column_mapping=map_column_mapping)
                smart.Sheets.update_rows(request_sheet_id,
                                         update_row_status(row=row,
                                                           column_mapping=request_column_mapping,
                                                           value='Green'))
            else:
                v_print('Simulation! This row would have been updated to green and added to the map sheet.\n')
    if not simulate:
        v_print('Row addition complete, colorizing rows...')
        colorize_rows(smart, map_sheet_id)
    v_print('All operations complete!')


def send_row(sheet_id: int,
             row: smartsheet.models.Row,
             request_column_mapping: dict,
             map_column_mapping: dict) -> None:
    """Main function for sending each row
    Takes the map sheet id, the row to be sent, and the request sheet
    {name: id} column map.
    Calculated the FY/Quarter number, thee map sheet {name: id}
    column map, and the dictionary of fy and quarter rows
    Creates an empty row, and sets the parent_id to the id of the
    quarter to which it belongs, and sets it to be added to the
    bottom of that row's children
    Iterates though all the cells, and if the name of that cell's
    column id is also in the map sheet, creates a new empty cell,
    copies over the value of the cell, sets the column id to be the
    column id that has the same name as the old cell's column, and
    appends that cell to the new row.
    Finally, that row is added to the map sheet
    Does not return anything.
    """
    v_print('  Sending row...')
    fy, q = calc_fy_q_hardcoded(get_cell_by_column_name(row=row,
                                                        column_name='Event Start Date',
                                                        col_map=request_column_mapping).value)
    v_print(f'  Fiscal Year: {fy}, Quarter: {q}')
    fy_q_dict = make_fy_q_dict(sheet_id, map_column_mapping)
    v_print(f'  Found these fiscal years in sheet:', *list(fy_q_dict))

    new_row = smartsheet.models.Row()

    for cell in row.cells:
        column_name = reverse_dict_search(request_column_mapping, cell.column_id)
        if column_name in map_column_mapping.keys():
            new_cell = smartsheet.models.Cell()
            new_cell.value = cell.value
            new_cell.column_id = map_column_mapping[column_name]
            new_row.cells.append(new_cell)

    row_parent_id = get_quarter_parent_id(fy, q, fy_q_dict, map_column_mapping, sheet_id)
    sib_id = sort_quarter_rows(sheet_id,
                               row_parent_id,
                               new_row,
                               map_column_mapping)
    if sib_id:
        new_row.sibling_id = sib_id
        new_row.above = True
        v_print(f'  Found sibling row with ID: {sib_id} (row will be added above its sibling)')
    else:
        new_row.parent_id = row_parent_id
        new_row.to_bottom = True
        v_print(f'  Sibling row not found. Falling back to parent ID of quarter ' +
                f'row (row will be added to bottom of quarter row\'s children)')

    update_row_status(row=new_row,
                      column_mapping=map_column_mapping,
                      value='Green')
    new_row.cells.append(smartsheet
                         .models.Cell(dict(value=True,
                                           column_id=map_column_mapping['TechX Service Request'])))
    v_print('  Checked "TechX Service Request" column')

    smart.Sheets.add_rows(sheet_id, new_row)
    v_print(f'  Row sent to sheet {sheet_id}!')


def update_row_status(row: smartsheet.models.Row,
                      column_mapping: dict,
                      column_name: str = 'TechX Status',
                      value: str = 'Green') -> smartsheet.models.Row:
    """ Updates a row's TechX Status Column to Green

    Takes a row and a column mapping, and optionally the column name
    and value to change it to.

    Creates a new row object with the old row's id and  non-empty cells,
    then changes the cell with the given column name to the specified
    value

    Returns the new row object with the updated color column
    """

    new_row = smartsheet.models.Row()
    new_row.id, new_row.cells = row.id, [cell for cell in row.cells if cell.value]
    get_cell_by_column_name(new_row, column_name, column_mapping).value = value
    v_print(f'  Updated {column_name} column in row {row.id if row.id is not None else "(no ID yet)"} to {value}')
    return new_row


def check_row(row: smartsheet.models.Row,
              column_mapping: dict,
              column_name: str = 'TechX Status',
              val_to_test: str = 'Yellow') -> bool:
    v_print(f'--checking row {get_cell_by_column_name(row, "Event Name", column_mapping).value} ({row.id}) ')
    return get_cell_by_column_name(row, column_name, column_mapping).value == val_to_test


def print_col_headings(cols: dict) -> None:  # prints column name and id for all columns, plus FY/Quarter
    v_print(*(column_format(col_title) for col_title in cols.keys()), 'FY/Quarter')
    v_print(*(str(col_id).ljust(16) for col_id in cols.values()))
    v_print()


def print_row(row: smartsheet.models.Row,
              column_mapping: dict,
              column_name: str = 'Event Start Date') -> None:
    # format, print the columns in the row + FY/Quarter
    v_print(*(column_format(cell.display_value or cell.value) for cell in row.cells), sep=' ', end=' ')
    fy, q = calc_fy_q_hardcoded(get_cell_by_column_name(row, column_name, column_mapping).value)
    v_print(f'FY{fy} Q{q}')


def column_format(item: str, just: int = 16) -> str:
    return (str(item)[:just - 2] + (str(item)[just - 2:] and '..')).ljust(just)


def column_name_to_id_map(sheet_id: int) -> dict:
    # returns a title:id dict of all columns in sheet
    return {column.title: column.id for column in smart.Sheets.get_columns(sheet_id, include_all=True).data}


def get_cell_by_column_name(row: smartsheet.models.Row,
                            column_name: str,
                            col_map: dict) -> smartsheet.models.Cell:
    return row.get_column(col_map[column_name])  # {NAME: ID}


def reverse_dict_search(search_dict: dict, search_value: str) -> str:
    for key, val in search_dict.items():
        if val == search_value:
            return key


def make_fy_q_dict(sheet_id: int,
                   column_mapping: dict,
                   column_name: str = 'Event Name') -> dict:
    fy_q_dict = {str(get_cell_by_column_name(fy,
                                             column_name,
                                             column_mapping).value): [fy, {}] for fy in find_fy_rows(sheet_id)}
    for year, (year_row, _) in fy_q_dict.items():
        quarters = find_child_rows(sheet_id, year_row.id)
        fy_q_dict[year][1] = {get_cell_by_column_name(quarter,
                                                      column_name,
                                                      column_mapping).value: quarter for quarter in quarters}
    return fy_q_dict


def find_fy_rows(sheet_id: int):
    return (
        row
        for row in smart.Sheets.get_sheet(sheet_id, level=2, include=['objectValue']).rows
        if not row.to_dict().get('parentId', False))


def find_child_rows(sheet_id: int, parent_row_id: int):
    return (
        row
        for row in smart.Sheets.get_sheet(sheet_id, level=2, include=['objectValue']).rows
        if row.to_dict().get('parentId', False) == parent_row_id)


def get_quarter_parent_id(fy: int, q: int, fy_q_dict: dict, column_mapping: dict, sheet_id: int) -> int:
    if ('FY' + str(fy)) not in fy_q_dict:
        add_fyq_rows(fy, column_mapping)
        fy_q_dict = make_fy_q_dict(sheet_id, column_mapping)
    return fy_q_dict['FY' + str(fy)][1]['Q' + str(q)].id


def add_fyq_rows(fy: int, column_mapping: dict) -> None:
    main_column_id = column_mapping['Event Name']

    fy_row = smartsheet.models.Row()
    fy_row.cells.append({
        "column_id": main_column_id,
        "value": "FY" + str(fy)
    })
    fy_add_result = smart.Sheets.add_rows(MAP_SHEET_ID, fy_row)  # add the FY row first
    fy_row_id = fy_add_result.result[0].id  # get the id of the row we just added

    quarter_rows = []
    for quarter in range(1, 5):
        new_row = smartsheet.models.Row()
        new_row.cells.append({
            "column_id": main_column_id,
            "value": f"Q{quarter}"
        })
        new_row.parent_id = fy_row_id  # set all the quarters to be children of the FY row
        new_row.to_bottom = True
        quarter_rows.append(new_row)

    smart.Sheets.add_rows(MAP_SHEET_ID, quarter_rows)  # add the quarter rows


def sort_quarter_rows(sheet_id: int,
                      quarter_row_id: int,
                      new_row: smartsheet.models.Row,
                      col_map: dict) -> int:
    rows_in_quarter = find_child_rows(sheet_id, quarter_row_id)
    new_row_start_date = datetime.strptime(get_cell_by_column_name(new_row,
                                                                   'Event Start Date',
                                                                   col_map).value,
                                           '%Y-%m-%d').date()
    for row in rows_in_quarter:
        row_date = datetime.strptime(get_cell_by_column_name(row,
                                                             'Event Start Date',
                                                             col_map).value,
                                     '%Y-%m-%d').date()
        if new_row_start_date < row_date:
            return row.id


def get_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description='Copy rows selected by Yellow ' +
                                                 'value in TechX Status column in ' +
                                                 'Request sheet to Map sheet',
                                     epilog=f'Written by Eugenia Liu and Daniel Karpelevitch')
    parser.add_argument('-V', '--version', action='version',
                        version=f"%(prog)s v1.0.3")
    parser.add_argument('-v', '--verbose', action='store_true', dest='verbose',
                        help='Enable verbose output')
    parser.add_argument('-s', '--simulate', action='store_true', dest='simulate',
                        help="Don't change sheets, just print what rows would be changed")
    return parser.parse_args()


if __name__ == '__main__':
    args = get_args()
    if args.verbose:
        def v_print(*print_args, **print_kwargs) -> None:
            print(*print_args, **print_kwargs)
    process_sheet(REQUEST_SHEET_ID, MAP_SHEET_ID, simulate=args.simulate)
