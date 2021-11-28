from map_calendar import MAP_SHEET, process_sheet as process_map
from smartsheet_intake_calendar import INTAKE_FORM_SHEET, process_sheet as process_intake

process_intake(INTAKE_FORM_SHEET)

process_map(MAP_SHEET)
