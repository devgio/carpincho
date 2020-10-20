def paste_df(df, ws, top_left_cell):
	from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
	top_left_col, top_left_row = coordinate_from_string(top_left_cell)
	top_left_col = column_index_from_string(top_left_col)
	rows, cols = df.shape
	for r in range(rows):
		for c in range(cols):
			ws.cell(row=top_left_row+r, column=top_left_col+c).value = df[c][r]


def pull_date(file, cell_range, worksheet=0, max_or_min=max, dayfirst=True):
	import xlrd, openpyxl, os, datetime
	from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
	import dateutil.parser as date_parser
	
	top_left_cell, bottom_right_cell = (cell.upper() for cell in cell_range.split(':'))
	top_left_col, top_left_row = coordinate_from_string(top_left_cell)
	top_left_col = column_index_from_string(top_left_col)
	bottom_right_col, bottom_right_row = coordinate_from_string(bottom_right_cell)
	bottom_right_col = column_index_from_string(bottom_right_col)

	directory, file_name = os.path.split(file)
	file_name, file_type = os.path.splitext(file_name)

	if file_type.lower() == '.xls':
		wb = xlrd.open_workbook(file, logfile=open(os.devnull, 'w')) # Added the logfile=open(os.devnull, 'w') to supress warnings.
		ws = wb.sheets()[worksheet]
		top_left_col -= 1
		bottom_right_col -= 1
		top_left_row -= 1
		bottom_right_row -= 1
		max_col, max_row = ws.ncols-1, ws.nrows-1
	if file_type.lower() in ('.xlsx', '.xlsm'):
		wb = openpyxl.load_workbook(file)
		ws = wb.worksheets[worksheet]
		max_col, max_row = ws.max_column, ws.max_row


	dates = set()
	for col in range(top_left_col, bottom_right_col+1):
		if col <= max_col:
			for row in range(top_left_row, bottom_right_row+1):
				if row <= max_row:
					if file_type.lower() in ('.xlsx', '.xlsm'):
						cell_value = ws.cell(row=row, column=col).value
					else:
						cell_value = ws.cell(row, col).value
					date = None
					if file_type.lower() == '.xls':
						try:
							date = datetime.datetime(*xlrd.xldate_as_tuple(cell_value, wb.datemode))
						except: pass
					if not date:
						try:
							if type(cell_value) == datetime.datetime:
								date = cell_value
							else:
								date = date_parser.parse(cell_value, dayfirst=dayfirst, fuzzy=True)
						except: pass
					if type(date) == datetime.datetime:
						dates.add(date)
	try:
		return max_or_min(dates)
	except:
		return None








