import smartsheet

color_white = 2
color_for_year_row = 21
color_for_quarter = [{'quarter row': 12, 'event row': 5},
                     {'quarter row': 14, 'event row': 7},
                     {'quarter row': 15, 'event row': 8},
                     {'quarter row': 16, 'event row': 9}]


def colorize_rows(smart: smartsheet.Smartsheet, sheet_id: int) -> None:
    from request_to_map.request_to_map_calendar import find_child_rows, find_fy_rows
    rows_to_update = []
    for year_row in find_fy_rows(sheet_id):
        year_color = color_white
        if any(row.parent_id == year_row.id for row in smart.Sheets.get_sheet(sheet_id).rows):
            year_color = color_for_year_row
        new_row = copy_row(year_row)
        new_row = process_row(new_row, year_color)
        rows_to_update.append(new_row)

        for quarter_num, quarter_row in enumerate(find_child_rows(sheet_id, year_row.id)):
            new_row = copy_row(quarter_row)
            new_row = process_row(new_row, color_for_quarter[quarter_num]["quarter row"])
            rows_to_update.append(new_row)

            for event_row in find_child_rows(sheet_id, quarter_row.id):
                new_row = copy_row(event_row)
                new_row = process_row(new_row, color_for_quarter[quarter_num]["event row"])
                rows_to_update.append(new_row)

    smart.Sheets.update_rows(sheet_id, rows_to_update)


def copy_row(row: smartsheet.models.row) -> smartsheet.models.row:
    return smartsheet.models.Row(dict(id=row.id,
                                      cells=row.cells,
                                      expanded=row.expanded,
                                      ))


def process_row(row: smartsheet.models.row, color: int) -> smartsheet.models.row:
    new_cells = []
    for cell in row.cells:
        cell_contents = dict(column_id=cell.column_id,
                             strict=False,
                             override_validation=True,
                             format=f',,,,,,,,,{color},,,,,,',
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

    row.cells = new_cells
    row.format = f',,,,,,,,,{color},,,,,,'
    return row
