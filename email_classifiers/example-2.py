from src.EmailClassifier import EmailClassifier
from src.excel_manipulation import pull_date
import openpyxl

def naming_function(file):
	date = pull_date(file, 'B16:B16').strftime('%Y-%m')
	return f'{date} MPS.xlsx'

def modifier_function(file):
	wb = openpyxl.load_workbook(file, data_only=True)
	ws = wb.active
	ws.title = 'MPS'
	wb.save(file)

EC = EmailClassifier(
	allowed_senders=('adam.smith@blackrock.com',), 
	subject_contains=('Sales', 'Stock',),
	filename_contains=('figures',),
	file_condition='isEXCEL',
	name_file_using=naming_function,
	modify_file_using=modifier_function,
	save_files_to='BlackRock sales',
	move_email_to=['Inbox', 'USA', 'BlackRock'],
	)
