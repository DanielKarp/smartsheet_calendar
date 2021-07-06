import smartsheet


smart = smartsheet.Smartsheet()  # use 'SMARTSHEET_ACCESS_TOKEN' env variable
smart.errors_as_exceptions(True)

CHANGE_AGENT = 'dkarpele_smartsheet_calendar'
smart.with_change_agent(CHANGE_AGENT)

info = smart.Server.server_info().formats.color

print(info)
