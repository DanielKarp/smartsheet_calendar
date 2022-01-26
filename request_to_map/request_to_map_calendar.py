import logging
from datetime import date, datetime

import smartsheet

from FY_Q_sort import calc_fy_q_hardcoded
from colorize import colorize_rows

logger = logging.getLogger('main')

smart = smartsheet.Smartsheet()  # use 'SMARTSHEET_ACCESS_TOKEN' env variable
smart.errors_as_exceptions(True)

CHANGE_AGENT = 'dkarpele_smartsheet_calendar'
smart.with_change_agent(CHANGE_AGENT)


def process_sheet(sheet_ids):
    _process_sheet(sheet_ids['source'], sheet_ids['destination'], simulate=False)


def _process_sheet(request_id: int,
                   map_id: int,
                   simulate: bool = False) -> None:
    """Main loop for processing the sheet
    takes the sheet ids for the request sheet to pull rows from, and
    the map sheet to send rows to. An optional simulate option does not
    alter any sheets, but only shows the rows that would be copied.
    Get rows from the request sheet, and builds {name: id} column map
    Prints out the column names and ids
    For each row in the request sheet, the TechX status column is checked
    if it is Yellow, the row is printed and if simulation is false, the
    row is sent to the map sheet and then the TechX Status column is
    changed to green.
    Does not return anything.
    """

    rows = smart.Sheets.get_sheet(request_id, level=2, include=['objectValue']).rows
    request_column_mapping = column_name_to_id_map(request_id)
    map_column_mapping = column_name_to_id_map(map_id)

    print_col_headings(request_column_mapping)
    rows_moved = 0
    for row in rows:
        if check_row(row, request_column_mapping):
            logger.debug(f'  ^row will be processed')
            print_row(row, request_column_mapping)
            if not simulate:
                rows_moved += 1
                send_row(sheet_id=map_id,
                         row=row,
                         request_column_mapping=request_column_mapping,
                         map_column_mapping=map_column_mapping)
                smart.Sheets.update_rows(request_id,
                                         update_row_status(row=row,
                                                           column_mapping=request_column_mapping,
                                                           value='Green'))
            else:
                logger.debug('Simulation! This row would have been updated to green and added to the map sheet.\n')
    logger.info(f'{rows_moved} rows moved')
    if not simulate:
        logger.info('colorizing rows...')
        colorize_rows(smart, map_id)
    logger.info('all operations complete!')


def send_row(sheet_id: int,
             row: smartsheet.models.Row,
             request_column_mapping: dict,
             map_column_mapping: dict) -> None:
    """Main function for sending each row
    Takes the map sheet id, the row to be sent, and the request sheet
    {name: id} column map.
    Calculated the FY/Quarter number, the map sheet {name: id}
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
    logger.debug('  Sending row...')
    fy, q = calc_fy_q_hardcoded(get_cell_by_column_name(row=row,
                                                        column_name='Event Start Date',
                                                        col_map=request_column_mapping).value)
    logger.debug(f'  Fiscal Year: {fy}, Quarter: {q}')
    fy_q_dict = make_fy_q_dict(sheet_id, map_column_mapping)
    logger.debug(f'  Found these fiscal years in sheet: {list(fy_q_dict)}')

    new_row = smartsheet.models.Row()

    for cell in row.cells:
        column_name = reverse_dict_search(request_column_mapping, cell.column_id)
        if column_name in map_column_mapping.keys():
            cell_contents = dict(column_id=map_column_mapping[column_name],
                                 strict=False,
                                 override_validation=True,
                                 )
            try:
                if cell.object_value.object_type == 8:
                    cell_value = smartsheet.models.MultiPicklistObjectValue()
                    cell_value.values = cell.object_value.values
                    cell_contents['object_value'] = cell_value
                else:
                    raise AttributeError
            except AttributeError:
                cell_contents['value'] = cell.value or ' '
            new_row.cells.append(smartsheet.models.Cell(cell_contents))

    row_parent_id = get_quarter_parent_id(fy, q, fy_q_dict, map_column_mapping, sheet_id)
    sib_id = sort_quarter_rows(sheet_id,
                               row_parent_id,
                               new_row,
                               map_column_mapping)
    if sib_id:
        new_row.sibling_id = sib_id
        new_row.above = True
        logger.debug(f'  Found sibling row with ID: {sib_id} (row will be added above its sibling)')
    else:
        new_row.parent_id = row_parent_id
        new_row.to_bottom = True
        logger.debug(f'  Sibling row not found. Falling back to parent ID of quarter ' +
                     f'row (row will be added to bottom of quarter row\'s children)')

    new_row.cells.append(smartsheet
                         .models.Cell(dict(value=True,
                                           column_id=map_column_mapping['TechX Service Request'])))
    logger.debug('  Checked "TechX Service Request" column')
    smart.Sheets.add_rows(sheet_id, new_row)
    logger.debug(f'  Row sent to sheet {sheet_id}!')


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
    new_cells = []
    new_row = smartsheet.models.Row(dict(id=row.id,
                                         cells=row.cells,
                                         expanded=row.expanded,
                                         ))
    for cell in row.cells:
        cell_contents = dict(column_id=cell.column_id,
                             strict=False,
                             override_validation=True,
                             )
        try:
            if cell.object_value.object_type == 8:
                cell_value = smartsheet.models.MultiPicklistObjectValue()
                cell_value.values = cell.object_value.values
                cell_contents['object_value'] = cell_value
            else:
                raise AttributeError
        except AttributeError:
            cell_contents['value'] = cell.value or ' '
        new_cells.append(smartsheet.models.Cell(cell_contents))

    new_row.cells = new_cells
    get_cell_by_column_name(new_row, column_name, column_mapping).value = value
    logger.debug(
        f'  Updated {column_name} column in row {new_row.id if new_row.id is not None else "(no ID yet)"} to {value}')
    return new_row


def check_row(row: smartsheet.models.Row,
              column_mapping: dict,
              column_name: str = 'TechX Status',
              val_to_test: str = 'Yellow') -> bool:
    logger.debug(f'  checking row {get_cell_by_column_name(row, "Event Name", column_mapping).value} ({row.id}) ')
    return get_cell_by_column_name(row, column_name, column_mapping).value == val_to_test


def print_col_headings(cols: dict) -> None:  # prints column name and id for all columns, plus FY/Quarter
    logger.debug(' '.join([column_format(col_title) for col_title in cols.keys()] + ['FY/Quarter']))
    logger.debug(' '.join(str(col_id).ljust(24) for col_id in cols.values()))


def print_row(row: smartsheet.models.Row,
              column_mapping: dict,
              column_name: str = 'Event Start Date') -> None:
    # format, print the columns in the row + FY/Quarter
    fy, q = calc_fy_q_hardcoded(get_cell_by_column_name(row, column_name, column_mapping).value)
    logger.debug(' '.join([column_format(c.display_value or c.value) for c in row.cells] + [f'FY{fy} Q{q}']))


def column_format(item: str, just: int = 24) -> str:
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
        if not row.to_dict().get('parentId', False) and str(row.cells[0].value).startswith('FY'))


def find_child_rows(sheet_id: int, parent_row_id: int):
    return (
        row
        for row in smart.Sheets.get_sheet(sheet_id, level=2, include=['objectValue']).rows
        if row.to_dict().get('parentId', False) == parent_row_id)


def get_quarter_parent_id(fy: int, q: int, fy_q_dict: dict, column_mapping: dict, sheet_id: int) -> int:
    if ('FY' + str(fy)) not in fy_q_dict:
        add_fyq_rows(fy, column_mapping, sheet_id)
        fy_q_dict = make_fy_q_dict(sheet_id, column_mapping)
    return fy_q_dict['FY' + str(fy)][1]['Q' + str(q)].id


def add_fyq_rows(fy: int, column_mapping: dict, sheet_id: int) -> None:
    main_column_id = column_mapping['Event Name']

    fy_row = smartsheet.models.Row()
    fy_row.cells.append({
        "column_id": main_column_id,
        "value": "FY" + str(fy)
    })
    fy_add_result = smart.Sheets.add_rows(sheet_id, fy_row)  # add the FY row first
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

    smart.Sheets.add_rows(sheet_id, quarter_rows)  # add the quarter rows


def sort_quarter_rows(sheet_id: int,
                      quarter_row_id: int,
                      new_row: smartsheet.models.Row,
                      col_map: dict) -> int:
    rows_in_quarter = find_child_rows(sheet_id, quarter_row_id)
    new_row_start_date = get_start_date(new_row, col_map)
    for row in rows_in_quarter:
        row_date = get_start_date(row, col_map)
        if new_row_start_date < row_date:
            return row.id


def get_start_date(row: smartsheet.models.row, col_map: dict, column: str = 'Event Start Date') -> date:
    if event_date := get_cell_by_column_name(row, column, col_map).value:
        try:
            return datetime.strptime(event_date, '%Y-%m-%d').date()
        except ValueError:
            return date(1, 1, 1)  # if string is not a date, put it in the past
    else:
        return date(3000, 1, 1)  # if the row has no date, put it in the future


if __name__ == '__main__':
    import yaml
    logging.basicConfig(level=logging.DEBUG)
    with open('../sheet_id.yaml') as yaml_file:
        config_sheet_id = yaml.safe_load(yaml_file)
    request_sheet_id = config_sheet_id['request to map']['source']
    map_sheet_id = config_sheet_id['request to map']['destination']
    _process_sheet(request_sheet_id, map_sheet_id, simulate=False)
