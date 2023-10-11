from datetime import datetime, date, timedelta

def format_date(paramData):
	result = paramData.split('/')
	result = date(int(result[2]), int(result[1]), int(result[0]))
	return result
