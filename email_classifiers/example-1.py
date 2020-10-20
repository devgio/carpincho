from src.EmailClassifier import EmailClassifier
from src.excel_manipulation import pull_date

def naming_function(file):
	date = pull_date(file, 'B1:B50').strftime('%Y-%m')
	return f'{date} Monthly Intel Data.xlsx'

EC = EmailClassifier(
	allowed_senders=('@intel.com',),
	subject_contains=('Monthly Data', 'another posible subject section'),
	filename_contains=('intel data',),
	file_condition='isEXCEL',
	name_file_using = naming_function,
	save_files_to='Intel Monthly Data',
	move_email_to=['Inbox', 'UK', 'Intel'],
	)

